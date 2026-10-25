// ZIP I/O backing the OPC package layer.
//
// Wraps `fflate` because it is the only ZIP library that is small (~30KB),
// synchronous, and works in both Node and browsers without a polyfill.
// Writing goes through `zipSync`; reading walks the central directory here and
// inflates each entry with fflate.
//
// Reading treats the archive as untrusted input. `unzipSync` inflates every
// entry to the end even when the output overflows the size the central
// directory declares, so a few-KB entry declaring 10 bytes can still burn
// seconds of synchronous CPU. Here an entry is inflated at most one chunk of
// input past the point its output exceeds the declared size, which caps the
// work at the sum of the central directory's declared sizes plus about one
// chunk of overrun.
//
// What we preserve on round-trip:
//
//   - Entry order. Entries are returned in central-directory order.
//   - Approximate per-entry compression method. Real PPTX files store media
//     (PNG, JPEG, ...) with no compression (STORE) and deflate everything
//     else. We mirror that heuristic on write, using extension to decide.
//     We do NOT round-trip the *exact* DEFLATE level — every PPTX consumer
//     accepts any valid level.
//
// What we do NOT preserve:
//
//   - File timestamps (`mtime`). The reference desktop app sets these to the current time
//     anyway; preserving the original would be misleading.
//   - ZIP64 records. They are read, but never written, and rejected when the
//     caller passes load limits.
//   - Extra fields, file attributes, UTF-8 flags. These are not load-bearing
//     for PPTX as consumed by the reference desktop app.

import { Inflate, type ZipOptions, inflateSync, strFromU8, zipSync } from 'fflate';

export interface ZipEntry {
  /** Path inside the archive, forward-slash separated (`ppt/slides/slide1.xml`). */
  name: string;
  /** Decompressed bytes. */
  data: Uint8Array;
}

export interface ZipReadResult {
  /** Entries in the order the central directory listed them. */
  entries: ZipEntry[];
}

/**
 * Hard caps for opening an untrusted archive. Sizes are uncompressed bytes as
 * the central directory declares them; they are checked before anything is
 * inflated, and inflation stops once an entry outgrows its declared size, so
 * they bound the real output too. XML parts are entries ending in `.xml` or
 * `.rels`.
 */
export interface LoadLimits {
  /** Most entries the archive may list, directory entries included. */
  maxEntries: number;
  /** Most uncompressed bytes across all entries. */
  maxTotalUncompressedBytes: number;
  /** Most uncompressed bytes in a single XML part. */
  maxXmlPartBytes: number;
  /** Most uncompressed bytes across all XML parts. */
  maxTotalXmlBytes: number;
}

export interface ZipReadOptions {
  /**
   * Caps checked against the central directory. Also rejects ZIP64 records,
   * which a PPTX never needs and which can restate counts and offsets.
   */
  limits?: LoadLimits;
}

const STORE_BY_EXTENSION = new Set([
  // images — already compressed
  'png',
  'jpg',
  'jpeg',
  'gif',
  'bmp',
  'tif',
  'tiff',
  'webp',
  // media — already compressed
  'mp3',
  'mp4',
  'm4a',
  'm4v',
  'wav',
  'wmv',
  'mov',
  'avi',
  'webm',
  'ogg',
  'wma',
  // archives — already compressed
  'pdf',
  'zip',
  'gz',
  '7z',
  'rar',
]);

const extensionOf = (name: string): string => {
  const dot = name.lastIndexOf('.');
  return dot < 0 ? '' : name.slice(dot + 1).toLowerCase();
};

const XML_EXTENSIONS = new Set(['xml', 'rels']);

/**
 * Returns true if the entry should be stored uncompressed. PPTX writers
 * conventionally STORE already-compressed payloads to avoid wasting CPU on
 * negligible savings.
 */
export const shouldStore = (name: string): boolean => STORE_BY_EXTENSION.has(extensionOf(name));

const SIG_LOCAL_HEADER = 0x04034b50;
const SIG_CENTRAL_HEADER = 0x02014b50;
const SIG_EOCD = 0x06054b50;
const SIG_ZIP64_EOCD = 0x06064b50;
const SIG_ZIP64_LOCATOR = 0x07064b50;
const EOCD_SIZE = 22;
const MAX_COMMENT_LENGTH = 0xffff;
const ZIP64_LOCATOR_SIZE = 20;
const ZIP64_EOCD_MIN_SIZE = 56;
const CENTRAL_HEADER_SIZE = 46;
const LOCAL_HEADER_SIZE = 30;
const ZIP64_EXTRA_ID = 0x0001;
const ZIP64_SENTINEL_32 = 0xffffffff;
const FLAG_UTF8 = 0x0800;
const METHOD_STORE = 0;
const METHOD_DEFLATE = 8;
// A 16KB DEFLATE chunk inflates to at most ~17MB (~30ms), which bounds the
// work done past an entry's declared size before the overflow is noticed.
const INFLATE_CHUNK_SIZE = 16 * 1024;
// DEFLATE needs at least 2 bits per 258-byte match, so it expands at most
// ~1032:1; the slack covers block headers and the first literals.
const MAX_DEFLATE_RATIO = 1032;
const DEFLATE_RATIO_SLACK = 1024;

const invalid = (reason: string): Error => new Error(`invalid ZIP: ${reason}`);

const u16 = (d: Uint8Array, at: number): number => d[at]! | (d[at + 1]! << 8);
const u32 = (d: Uint8Array, at: number): number =>
  (d[at]! | (d[at + 1]! << 8) | (d[at + 2]! << 16) | (d[at + 3]! << 24)) >>> 0;
const u64 = (d: Uint8Array, at: number): number => {
  const value = u32(d, at) + u32(d, at + 4) * 2 ** 32;
  if (!Number.isSafeInteger(value)) throw invalid('ZIP64 value out of range');
  return value;
};

const requireRange = (d: Uint8Array, at: number, length: number, what: string): void => {
  if (at < 0 || length < 0 || at + length > d.length) throw invalid(`${what} is out of bounds`);
};

const findEocd = (d: Uint8Array): number => {
  const lowest = Math.max(0, d.length - EOCD_SIZE - MAX_COMMENT_LENGTH);
  for (let at = d.length - EOCD_SIZE; at >= lowest; at--) {
    if (u32(d, at) === SIG_EOCD) return at;
  }
  throw invalid('end of central directory not found');
};

interface CentralDirectory {
  offset: number;
  count: number;
}

const readCentralDirectoryLocation = (d: Uint8Array, rejectZip64: boolean): CentralDirectory => {
  const eocd = findEocd(d);
  const locator = eocd - ZIP64_LOCATOR_SIZE;
  // A ZIP64 locator overrides the 32-bit count and offset even when they are
  // not sentinels, so honour it; the per-entry size checks below hold either way.
  // One that does not lead to a ZIP64 EOCD is ignored, as fflate did.
  if (locator >= 0 && u32(d, locator) === SIG_ZIP64_LOCATOR) {
    if (rejectZip64) {
      throw invalid('ZIP64 end of central directory locator is not accepted under load limits');
    }
    const zip64Eocd = u32(d, locator + 8) + u32(d, locator + 12) * 2 ** 32;
    if (zip64Eocd + ZIP64_EOCD_MIN_SIZE <= d.length && u32(d, zip64Eocd) === SIG_ZIP64_EOCD) {
      return { count: u64(d, zip64Eocd + 32), offset: u64(d, zip64Eocd + 48) };
    }
  }
  return { count: u16(d, eocd + 10), offset: u32(d, eocd + 16) };
};

interface EntrySizes {
  compressedSize: number;
  uncompressedSize: number;
  localHeaderOffset: number;
}

const usesZip64 = (sizes: EntrySizes): boolean =>
  sizes.uncompressedSize === ZIP64_SENTINEL_32 ||
  sizes.compressedSize === ZIP64_SENTINEL_32 ||
  sizes.localHeaderOffset === ZIP64_SENTINEL_32;

// The ZIP64 extra field lists only the values whose 32-bit slot holds the
// sentinel, in this fixed order (APPNOTE 4.5.3).
const resolveZip64Sizes = (
  d: Uint8Array,
  extraStart: number,
  extraEnd: number,
  sizes: EntrySizes,
): EntrySizes => {
  if (!usesZip64(sizes)) return sizes;
  const needUncompressed = sizes.uncompressedSize === ZIP64_SENTINEL_32;
  const needCompressed = sizes.compressedSize === ZIP64_SENTINEL_32;
  const needOffset = sizes.localHeaderOffset === ZIP64_SENTINEL_32;
  for (let at = extraStart; at + 4 <= extraEnd; at += 4 + u16(d, at + 2)) {
    if (u16(d, at) !== ZIP64_EXTRA_ID) continue;
    const fieldEnd = at + 4 + u16(d, at + 2);
    if (fieldEnd > extraEnd) break;
    let cursor = at + 4;
    const next = (): number => {
      if (cursor + 8 > fieldEnd) throw invalid('ZIP64 extra field is truncated');
      const value = u64(d, cursor);
      cursor += 8;
      return value;
    };
    return {
      uncompressedSize: needUncompressed ? next() : sizes.uncompressedSize,
      compressedSize: needCompressed ? next() : sizes.compressedSize,
      localHeaderOffset: needOffset ? next() : sizes.localHeaderOffset,
    };
  }
  throw invalid('ZIP64 sizes without a ZIP64 extra field');
};

const overflowError = (name: string, declaredSize: number): Error =>
  invalid(`entry "${name}" inflates beyond its declared size of ${declaredSize} bytes`);

const corruptError = (name: string, error: unknown): Error => {
  const reason = error instanceof Error ? error.message : String(error);
  return invalid(`entry "${name}" is not valid DEFLATE data: ${reason}`);
};

// One-shot inflate into a buffer one byte longer than declared: fflate drops
// writes past a caller's buffer, so a full buffer means overflow. Streaming
// `Inflate` costs ~70µs per entry, which dominates packages of many tiny parts.
const inflateSmallEntry = (
  compressed: Uint8Array,
  declaredSize: number,
  name: string,
): Uint8Array => {
  // fflate returns the whole buffer for empty input; an empty stream yields nothing.
  if (compressed.length === 0) return new Uint8Array(0);
  let out: Uint8Array;
  try {
    out = inflateSync(compressed, { out: new Uint8Array(declaredSize + 1) });
  } catch (error) {
    // A stored block that overruns the buffer surfaces as a RangeError from `set`.
    if (error instanceof RangeError) throw overflowError(name, declaredSize);
    throw corruptError(name, error);
  }
  if (out.length > declaredSize) throw overflowError(name, declaredSize);
  return out;
};

// Fills a buffer of the declared size, as `unzipSync` did, so a large entry is
// held once.
const inflateLargeEntry = (
  compressed: Uint8Array,
  declaredSize: number,
  name: string,
): Uint8Array => {
  const out = new Uint8Array(declaredSize);
  let overflowed = false;
  let produced = 0;
  const inflater = new Inflate((chunk) => {
    if (produced + chunk.length > declaredSize) {
      overflowed = true;
      throw overflowError(name, declaredSize);
    }
    out.set(chunk, produced);
    produced += chunk.length;
  });
  for (let at = 0; at < compressed.length; at += INFLATE_CHUNK_SIZE) {
    const end = Math.min(at + INFLATE_CHUNK_SIZE, compressed.length);
    try {
      inflater.push(compressed.subarray(at, end), end === compressed.length);
    } catch (error) {
      if (overflowed) throw error;
      throw corruptError(name, error);
    }
  }
  return out.subarray(0, produced);
};

const inflateEntry = (compressed: Uint8Array, declaredSize: number, name: string): Uint8Array => {
  // Checked before allocating so a tiny entry cannot reserve gigabytes.
  if (declaredSize > compressed.length * MAX_DEFLATE_RATIO + DEFLATE_RATIO_SLACK) {
    throw invalid(
      `entry "${name}" declares ${declaredSize} bytes, more than ${compressed.length} compressed bytes can hold`,
    );
  }
  const out =
    compressed.length <= INFLATE_CHUNK_SIZE
      ? inflateSmallEntry(compressed, declaredSize, name)
      : inflateLargeEntry(compressed, declaredSize, name);
  if (out.length !== declaredSize) {
    throw invalid(
      `entry "${name}" inflates to ${out.length} bytes, not its declared ${declaredSize}`,
    );
  }
  return out;
};

const limitError = (subject: string, limit: keyof LoadLimits, max: number): Error =>
  invalid(`${subject}, over the ${limit} limit of ${max}`);

// A NaN limit would make every comparison false and silently disable it.
const requireLimits = (limits: LoadLimits): void => {
  for (const [key, value] of Object.entries(limits)) {
    if (!(value >= 0)) throw new RangeError(`load limit ${key} must be a non-negative number`);
  }
};

interface CentralRecord extends EntrySizes {
  name: string;
  method: number;
}

const readCentralDirectory = (
  input: Uint8Array,
  limits: LoadLimits | undefined,
): CentralRecord[] => {
  const { offset, count } = readCentralDirectoryLocation(input, limits !== undefined);
  if (limits !== undefined && count > limits.maxEntries) {
    throw limitError(`archive lists ${count} entries`, 'maxEntries', limits.maxEntries);
  }
  const records: CentralRecord[] = [];
  const seen = new Set<string>();
  let totalBytes = 0;
  let totalXmlBytes = 0;
  let cursor = offset;
  for (let i = 0; i < count; i++) {
    requireRange(input, cursor, CENTRAL_HEADER_SIZE, 'central directory');
    if (u32(input, cursor) !== SIG_CENTRAL_HEADER) throw invalid('bad central directory signature');
    const flags = u16(input, cursor + 8);
    const method = u16(input, cursor + 10);
    const nameLength = u16(input, cursor + 28);
    const extraLength = u16(input, cursor + 30);
    const commentLength = u16(input, cursor + 32);
    const nameStart = cursor + CENTRAL_HEADER_SIZE;
    const extraStart = nameStart + nameLength;
    const recordEnd = extraStart + extraLength + commentLength;
    requireRange(input, cursor, recordEnd - cursor, 'central directory');
    const name = strFromU8(input.subarray(nameStart, extraStart), (flags & FLAG_UTF8) === 0);
    const declared: EntrySizes = {
      compressedSize: u32(input, cursor + 20),
      uncompressedSize: u32(input, cursor + 24),
      localHeaderOffset: u32(input, cursor + 42),
    };
    cursor = recordEnd;

    if (seen.has(name)) throw invalid(`duplicate entry "${name}"`);
    seen.add(name);

    if (limits === undefined) {
      records.push({
        name,
        method,
        ...resolveZip64Sizes(input, extraStart, extraStart + extraLength, declared),
      });
      continue;
    }
    if (usesZip64(declared)) {
      throw invalid(`entry "${name}" uses ZIP64 sizes, which are not accepted under load limits`);
    }
    const size = declared.uncompressedSize;
    totalBytes += size;
    if (totalBytes > limits.maxTotalUncompressedBytes) {
      throw limitError(
        `entries declare ${totalBytes} uncompressed bytes`,
        'maxTotalUncompressedBytes',
        limits.maxTotalUncompressedBytes,
      );
    }
    if (XML_EXTENSIONS.has(extensionOf(name))) {
      if (size > limits.maxXmlPartBytes) {
        throw limitError(
          `entry "${name}" declares ${size} bytes`,
          'maxXmlPartBytes',
          limits.maxXmlPartBytes,
        );
      }
      totalXmlBytes += size;
      if (totalXmlBytes > limits.maxTotalXmlBytes) {
        throw limitError(
          `XML parts declare ${totalXmlBytes} bytes`,
          'maxTotalXmlBytes',
          limits.maxTotalXmlBytes,
        );
      }
    }
    records.push({ name, method, ...declared });
  }
  return records;
};

/**
 * Reads a ZIP archive (a `.pptx` file is a ZIP) and returns entries in
 * central-directory order. Throws if `input` is not a valid ZIP, if two
 * entries share a name, if an entry's data does not match the size its
 * central-directory record declares, or if the archive breaks
 * `options.limits`. The whole central directory is read before any entry is
 * inflated.
 */
export const readZip = (input: Uint8Array, options: ZipReadOptions = {}): ZipReadResult => {
  if (options.limits !== undefined) requireLimits(options.limits);
  const entries: ZipEntry[] = [];
  for (const record of readCentralDirectory(input, options.limits)) {
    const { name, method, compressedSize, uncompressedSize, localHeaderOffset } = record;
    requireRange(input, localHeaderOffset, LOCAL_HEADER_SIZE, `local header of "${name}"`);
    if (u32(input, localHeaderOffset) !== SIG_LOCAL_HEADER) {
      throw invalid(`bad local header signature for "${name}"`);
    }
    const dataStart =
      localHeaderOffset +
      LOCAL_HEADER_SIZE +
      u16(input, localHeaderOffset + 26) +
      u16(input, localHeaderOffset + 28);
    requireRange(input, dataStart, compressedSize, `data of "${name}"`);
    const compressed = input.subarray(dataStart, dataStart + compressedSize);

    if (method === METHOD_STORE) {
      if (compressedSize !== uncompressedSize) {
        throw invalid(`stored entry "${name}" has mismatched sizes`);
      }
      entries.push({ name, data: compressed.slice() });
    } else if (method === METHOD_DEFLATE) {
      entries.push({ name, data: inflateEntry(compressed, uncompressedSize, name) });
    } else {
      throw invalid(`entry "${name}" uses unsupported compression method ${method}`);
    }
  }
  return { entries };
};

/**
 * Writes a ZIP archive. Entries are written in the order given. Each entry's
 * compression method is chosen from its extension via `shouldStore`. Callers
 * who want a specific compression level can pass `levelOverride`.
 */
/** DEFLATE compression level. 0 stores uncompressed; 9 is best compression. */
export type CompressionLevel = 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;

export const writeZip = (
  entries: ReadonlyArray<ZipEntry>,
  options: { level?: CompressionLevel } = {},
): Uint8Array => {
  const defaultLevel: CompressionLevel = options.level ?? 6;
  const obj: Record<string, [Uint8Array, ZipOptions]> = {};
  for (const e of entries) {
    const level: CompressionLevel = shouldStore(e.name) ? 0 : defaultLevel;
    obj[e.name] = [e.data, { level }];
  }
  return zipSync(obj);
};

// ZIP I/O backing the OPC package layer.
//
// Wraps `fflate` because it is the only ZIP library that is small (~30KB),
// synchronous, and works in both Node and browsers without a polyfill. The
// wrapper is deliberately thin: it just translates between fflate's
// object-keyed API and an ordered array of entries, plus picks a sensible
// compression level per entry.
//
// What we preserve on round-trip:
//
//   - Entry order. fflate's `unzipSync` returns a plain object whose key
//     iteration order matches the central directory, so we use `Object.keys`
//     to recover the ordering. JS objects preserve insertion order for string
//     keys, so this is well-defined.
//   - Approximate per-entry compression method. Real PPTX files store media
//     (PNG, JPEG, ...) with no compression (STORE) and deflate everything
//     else. We mirror that heuristic on write, using extension to decide.
//     We do NOT round-trip the *exact* DEFLATE level — every PPTX consumer
//     accepts any valid level.
//
// What we do NOT preserve:
//
//   - File timestamps (`mtime`). PowerPoint sets these to the current time
//     anyway; preserving the original would be misleading.
//   - ZIP64 extensions. fflate auto-promotes when needed.
//   - Extra fields, file attributes, UTF-8 flags. These are not load-bearing
//     for PPTX as consumed by PowerPoint.

import { type ZipOptions, unzipSync, zipSync } from 'fflate';

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

/**
 * Returns true if the entry should be stored uncompressed. PPTX writers
 * conventionally STORE already-compressed payloads to avoid wasting CPU on
 * negligible savings.
 */
export const shouldStore = (name: string): boolean => STORE_BY_EXTENSION.has(extensionOf(name));

/**
 * Reads a ZIP archive (a `.pptx` file is a ZIP) and returns entries in
 * central-directory order. Throws if `input` is not a valid ZIP.
 */
export const readZip = (input: Uint8Array): ZipReadResult => {
  const obj = unzipSync(input);
  const entries: ZipEntry[] = [];
  for (const name of Object.keys(obj)) {
    const data = obj[name];
    if (data === undefined) continue;
    entries.push({ name, data });
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

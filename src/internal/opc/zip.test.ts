import { deflateSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { type LoadLimits, readZip, writeZip } from './zip.ts';

const utf8 = (s: string): Uint8Array => new TextEncoder().encode(s);

const concat = (chunks: ReadonlyArray<Uint8Array>): Uint8Array => {
  const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
  let at = 0;
  for (const c of chunks) {
    out.set(c, at);
    at += c.length;
  }
  return out;
};

const le = (value: number, bytes: number): number[] => {
  const out: number[] = [];
  for (let i = 0; i < bytes; i++) out.push(Math.floor(value / 2 ** (8 * i)) % 256);
  return out;
};

const SENTINEL_32 = 0xffffffff;

interface RawEntry {
  name: string;
  method: 0 | 8;
  /** Bytes as stored in the archive (already deflated when method is 8). */
  stored: Uint8Array;
  declaredSize: number;
  /** Put both sizes in a ZIP64 extra field behind 0xFFFFFFFF sentinels. */
  zip64Sizes?: boolean;
}

const localHeader = (e: RawEntry): Uint8Array => {
  const name = utf8(e.name);
  return new Uint8Array([
    ...le(0x04034b50, 4),
    ...le(20, 2),
    ...le(0x0800, 2),
    ...le(e.method, 2),
    ...le(0, 4),
    ...le(0, 4),
    ...le(e.stored.length, 4),
    ...le(e.declaredSize, 4),
    ...le(name.length, 2),
    ...le(0, 2),
    ...name,
  ]);
};

const centralRecord = (e: RawEntry, localOffset: number): Uint8Array => {
  const name = utf8(e.name);
  const extra = e.zip64Sizes
    ? [...le(0x0001, 2), ...le(16, 2), ...le(e.declaredSize, 8), ...le(e.stored.length, 8)]
    : [];
  return new Uint8Array([
    ...le(0x02014b50, 4),
    ...le(45, 2),
    ...le(45, 2),
    ...le(0x0800, 2),
    ...le(e.method, 2),
    ...le(0, 4),
    ...le(0, 4),
    ...le(e.zip64Sizes ? SENTINEL_32 : e.stored.length, 4),
    ...le(e.zip64Sizes ? SENTINEL_32 : e.declaredSize, 4),
    ...le(name.length, 2),
    ...le(extra.length, 2),
    ...le(0, 2),
    ...le(0, 2),
    ...le(0, 2),
    ...le(0, 4),
    ...le(localOffset, 4),
    ...name,
    ...extra,
  ]);
};

const eocd = (count: number, cdSize: number, cdOffset: number): Uint8Array =>
  new Uint8Array([
    ...le(0x06054b50, 4),
    ...le(0, 2),
    ...le(0, 2),
    ...le(count, 2),
    ...le(count, 2),
    ...le(cdSize, 4),
    ...le(cdOffset, 4),
    ...le(0, 2),
  ]);

const zip64Eocd = (count: number, cdSize: number, cdOffset: number): Uint8Array =>
  new Uint8Array([
    ...le(0x06064b50, 4),
    ...le(44, 8),
    ...le(45, 2),
    ...le(45, 2),
    ...le(0, 4),
    ...le(0, 4),
    ...le(count, 8),
    ...le(count, 8),
    ...le(cdSize, 8),
    ...le(cdOffset, 8),
  ]);

const zip64Locator = (zip64EocdOffset: number): Uint8Array =>
  new Uint8Array([...le(0x07064b50, 4), ...le(0, 4), ...le(zip64EocdOffset, 8), ...le(1, 4)]);

/** Local headers + data, returning each entry's local header offset. */
const writeLocals = (
  entries: ReadonlyArray<RawEntry>,
  start = 0,
): { bytes: Uint8Array; offsets: number[] } => {
  const chunks: Uint8Array[] = [];
  const offsets: number[] = [];
  let at = start;
  for (const e of entries) {
    const header = localHeader(e);
    offsets.push(at);
    chunks.push(header, e.stored);
    at += header.length + e.stored.length;
  }
  return { bytes: concat(chunks), offsets };
};

const buildZip = (
  entries: ReadonlyArray<RawEntry>,
  options: { zip64?: boolean } = {},
): Uint8Array => {
  const locals = writeLocals(entries);
  const cd = concat(entries.map((e, i) => centralRecord(e, locals.offsets[i]!)));
  const cdOffset = locals.bytes.length;
  if (!options.zip64) return concat([locals.bytes, cd, eocd(entries.length, cd.length, cdOffset)]);
  const zip64EocdOffset = cdOffset + cd.length;
  return concat([
    locals.bytes,
    cd,
    zip64Eocd(entries.length, cd.length, cdOffset),
    zip64Locator(zip64EocdOffset),
    eocd(0xffff, SENTINEL_32, SENTINEL_32),
  ]);
};

const deflated = (name: string, data: Uint8Array, declaredSize = data.length): RawEntry => ({
  name,
  method: 8,
  stored: deflateSync(data),
  declaredSize,
});

// A final fixed-Huffman DEFLATE block: one literal 0 followed by `matches`
// copies of (length 258, distance 1). Inflates to 1 + 258 * matches zero bytes
// at ~160:1, which lets a test hold a gigabyte-scale entry in a few MB.
const zeroBomb = (matches: number): Uint8Array => {
  const out = new Uint8Array(Math.ceil((3 + 8 + matches * 13 + 7) / 8) + 1);
  let bitPos = 0;
  const bits = (value: number, count: number): void => {
    for (let i = 0; i < count; i++, bitPos++) {
      if ((value >> i) & 1) out[bitPos >> 3]! |= 1 << (bitPos & 7);
    }
  };
  // Huffman codes are packed most-significant bit first.
  const code = (value: number, count: number): void => {
    for (let i = count - 1; i >= 0; i--, bitPos++) {
      if ((value >> i) & 1) out[bitPos >> 3]! |= 1 << (bitPos & 7);
    }
  };
  bits(1, 1); // BFINAL
  bits(1, 2); // BTYPE = fixed Huffman
  code(0b00110000, 8); // literal 0
  for (let i = 0; i < matches; i++) {
    code(0b11000101, 8); // length 258
    code(0, 5); // distance 1
  }
  code(0, 7); // end of block
  return out.subarray(0, Math.ceil(bitPos / 8));
};

describe('readZip on hostile input', () => {
  it('reads an archive it did not write itself', () => {
    const xml = utf8('<a/>'.repeat(1000));
    const zip = buildZip([
      deflated('ppt/a.xml', xml),
      { name: 'media/b.png', method: 0, stored: new Uint8Array([1, 2, 3]), declaredSize: 3 },
      deflated('empty.xml', new Uint8Array(0)),
    ]);
    const { entries } = readZip(zip);
    expect(entries.map((e) => e.name)).toEqual(['ppt/a.xml', 'media/b.png', 'empty.xml']);
    expect(entries[0]?.data).toEqual(xml);
    expect(entries[1]?.data).toEqual(new Uint8Array([1, 2, 3]));
    expect(entries[2]?.data).toEqual(new Uint8Array(0));
  });

  it('stops inflating an entry as soon as it outgrows its declared size', () => {
    // The tail is not valid DEFLATE: reaching it would fail with a different
    // error, so the message proves inflation stopped at the first chunk.
    const stored = concat([
      zeroBomb(40_000).subarray(0, 64 * 1024),
      new Uint8Array(1 << 20).fill(0xff),
    ]);
    const zip = buildZip([{ name: 'bomb.xml', method: 8, stored, declaredSize: 10 }]);
    expect(() => readZip(zip)).toThrow(/beyond its declared size of 10 bytes/);
  });

  it('rejects a 10-byte entry that really inflates to 1 GB', () => {
    const matches = Math.ceil(2 ** 30 / 258);
    const zip = buildZip([
      { name: 'bomb.xml', method: 8, stored: zeroBomb(matches), declaredSize: 10 },
    ]);
    expect(() => readZip(zip)).toThrow(/beyond its declared size/);
  });

  it('rejects a small entry that outgrows its declared size', () => {
    const zip = buildZip([deflated('small.xml', utf8('<a/>'.repeat(100)), 10)]);
    expect(() => readZip(zip)).toThrow(/"small.xml" inflates beyond its declared size of 10 bytes/);
  });

  it('rejects a small stored DEFLATE block that outgrows its declared size', () => {
    // BFINAL stored block: LEN 8, NLEN ~8, then 8 bytes.
    const block = new Uint8Array([0x01, 0x08, 0x00, 0xf7, 0xff, ...new Uint8Array(8)]);
    const zip = buildZip([{ name: 'raw.xml', method: 8, stored: block, declaredSize: 4 }]);
    expect(() => readZip(zip)).toThrow(/"raw.xml" inflates beyond its declared size of 4 bytes/);
  });

  it('rejects a declared size its compressed bytes cannot reach before allocating it', () => {
    const stored = deflateSync(utf8('<a/>'));
    const zip = buildZip([{ name: 'huge.xml', method: 8, stored, declaredSize: SENTINEL_32 - 1 }]);
    expect(() => readZip(zip)).toThrow(/^invalid ZIP: entry "huge.xml" declares 4294967294 bytes/);
  });

  it('accepts a maximally compressed entry', () => {
    // Zeros deflate at ~1030:1, right at the ratio inflateEntry allows.
    const size = 4 << 20;
    const zip = buildZip([deflated('zeros.bin', new Uint8Array(size))]);
    expect(readZip(zip).entries[0]?.data.length).toBe(size);
  });

  it('rejects an entry that inflates to less than its declared size', () => {
    const zip = buildZip([deflated('short.xml', utf8('<a/>'), 100)]);
    expect(() => readZip(zip)).toThrow(/inflates to 4 bytes, not its declared 100/);
  });

  it('reports corrupt DEFLATE data as an invalid ZIP naming the entry', () => {
    const zip = buildZip([
      { name: 'bad.xml', method: 8, stored: new Uint8Array([0xff, 0xff, 0xff]), declaredSize: 4 },
    ]);
    expect(() => readZip(zip)).toThrow(/^invalid ZIP: entry "bad.xml" is not valid DEFLATE data: /);
  });

  it('rejects a stored entry whose sizes disagree', () => {
    const zip = buildZip([
      { name: 'a.bin', method: 0, stored: new Uint8Array(8), declaredSize: 4 },
    ]);
    expect(() => readZip(zip)).toThrow(/mismatched sizes/);
  });

  it('enforces the sizes from the ZIP64 central directory a locator points at', () => {
    // The 32-bit EOCD describes a harmless directory; the ZIP64 locator in
    // front of it redirects to a second directory whose entry lies about its
    // size. A size check that reads one directory while the inflater reads the
    // other is exactly the gap this closes.
    const harmless = deflated('ok.xml', utf8('<ok/>'));
    const lying: RawEntry = {
      name: 'bomb.xml',
      method: 8,
      stored: deflateSync(new Uint8Array(1 << 20)),
      declaredSize: 10,
    };
    const locals = writeLocals([harmless, lying]);
    const lyingCd = centralRecord(lying, locals.offsets[1]!);
    const harmlessCd = centralRecord(harmless, locals.offsets[0]!);
    const lyingCdOffset = locals.bytes.length;
    const harmlessCdOffset = lyingCdOffset + lyingCd.length;
    const zip64EocdOffset = harmlessCdOffset + harmlessCd.length;
    const withoutLocator = concat([
      locals.bytes,
      lyingCd,
      harmlessCd,
      eocd(1, harmlessCd.length, harmlessCdOffset),
    ]);
    expect(readZip(withoutLocator).entries.map((e) => e.name)).toEqual(['ok.xml']);

    const withLocator = concat([
      locals.bytes,
      lyingCd,
      harmlessCd,
      zip64Eocd(1, lyingCd.length, lyingCdOffset),
      zip64Locator(zip64EocdOffset),
      eocd(1, harmlessCd.length, harmlessCdOffset),
    ]);
    expect(() => readZip(withLocator)).toThrow(/"bomb.xml" inflates beyond its declared size/);
  });

  it('ignores a ZIP64 locator signature that does not lead to a ZIP64 directory', () => {
    const entry = deflated('a.xml', utf8('<a/>'));
    const locals = writeLocals([entry]);
    const cd = centralRecord(entry, locals.offsets[0]!);
    const zip = concat([
      locals.bytes,
      cd,
      zip64Locator(0),
      eocd(1, cd.length, locals.bytes.length),
    ]);
    expect(readZip(zip).entries.map((e) => e.name)).toEqual(['a.xml']);
  });

  it('reads honest ZIP64 records and sizes', () => {
    const xml = utf8('<x/>'.repeat(100));
    const zip = buildZip(
      [{ ...deflated('a.xml', xml), zip64Sizes: true }, deflated('b.xml', utf8('<b/>'))],
      { zip64: true },
    );
    const { entries } = readZip(zip);
    expect(entries.map((e) => e.name)).toEqual(['a.xml', 'b.xml']);
    expect(entries[0]?.data).toEqual(xml);
  });

  it('enforces sizes declared in a ZIP64 extra field', () => {
    const zip = buildZip(
      [{ ...deflated('a.xml', new Uint8Array(1 << 20)), declaredSize: 10, zip64Sizes: true }],
      { zip64: true },
    );
    expect(() => readZip(zip)).toThrow(/beyond its declared size of 10 bytes/);
  });

  it('rejects duplicate entry names', () => {
    const zip = buildZip([deflated('a.xml', utf8('<a/>')), deflated('a.xml', utf8('<b/>'))]);
    expect(() => readZip(zip)).toThrow(/duplicate entry "a.xml"/);
  });

  it('rejects offsets and sizes that point outside the archive', () => {
    const good = writeZip([{ name: 'a.xml', data: utf8('<a/>') }]);
    expect(() => readZip(good.subarray(0, good.length - 1))).toThrow(/invalid ZIP/);
    const entry = deflated('a.xml', utf8('<a/>'));
    const zip = buildZip([entry]);
    const cdOffset = localHeader(entry).length + entry.stored.length;
    const badOffset = zip.slice();
    badOffset.set(le(0x7fffffff, 4), cdOffset + 42);
    expect(() => readZip(badOffset)).toThrow(/local header of "a.xml" is out of bounds/);
    const badSize = zip.slice();
    badSize.set(le(0x7fffffff, 4), cdOffset + 20);
    expect(() => readZip(badSize)).toThrow(/data of "a.xml" is out of bounds/);
  });

  it('rejects unsupported compression methods', () => {
    const entry: RawEntry = {
      name: 'a.bin',
      method: 0,
      stored: new Uint8Array(4),
      declaredSize: 4,
    };
    const zip = buildZip([entry]);
    const cdOffset = localHeader(entry).length + entry.stored.length;
    zip.set(le(12, 2), cdOffset + 10);
    expect(() => readZip(zip)).toThrow(/unsupported compression method 12/);
  });

  it('rejects input without an end of central directory', () => {
    expect(() => readZip(new Uint8Array(100))).toThrow(/end of central directory not found/);
    expect(() => readZip(new Uint8Array(0))).toThrow(/end of central directory not found/);
  });
});

describe('readZip with load limits', () => {
  const roomy: LoadLimits = {
    maxEntries: 100,
    maxTotalUncompressedBytes: 1 << 20,
    maxXmlPartBytes: 1 << 16,
    maxTotalXmlBytes: 1 << 18,
  };
  const xmlEntries = (count: number, size: number): RawEntry[] =>
    Array.from({ length: count }, (_, i) => deflated(`ppt/p${i}.xml`, new Uint8Array(size)));

  it('reads an archive within the limits exactly as without them', () => {
    const zip = buildZip([
      deflated('ppt/a.xml', utf8('<a/>'.repeat(1000))),
      deflated('ppt/_rels/a.xml.rels', utf8('<r/>')),
      { name: 'media/b.png', method: 0, stored: new Uint8Array([1, 2, 3]), declaredSize: 3 },
    ]);
    expect(readZip(zip, { limits: roomy })).toEqual(readZip(zip));
  });

  it('accepts sizes equal to each limit', () => {
    const zip = buildZip([...xmlEntries(2, 100), deflated('media/a.bin', new Uint8Array(100))]);
    const limits: LoadLimits = {
      maxEntries: 3,
      maxTotalUncompressedBytes: 300,
      maxXmlPartBytes: 100,
      maxTotalXmlBytes: 200,
    };
    expect(readZip(zip, { limits }).entries).toHaveLength(3);
  });

  it('rejects more entries than maxEntries', () => {
    const zip = buildZip(xmlEntries(4, 1));
    expect(() => readZip(zip, { limits: { ...roomy, maxEntries: 3 } })).toThrow(
      /^invalid ZIP: archive lists 4 entries, over the maxEntries limit of 3$/,
    );
  });

  it('rejects declared sizes over maxTotalUncompressedBytes', () => {
    const zip = buildZip([
      deflated('media/a.bin', new Uint8Array(600)),
      deflated('media/b.bin', new Uint8Array(600)),
    ]);
    expect(() => readZip(zip, { limits: { ...roomy, maxTotalUncompressedBytes: 1000 } })).toThrow(
      /^invalid ZIP: entries declare 1200 uncompressed bytes, over the maxTotalUncompressedBytes limit of 1000$/,
    );
  });

  it('caps .xml and .rels parts but not media under maxXmlPartBytes', () => {
    const limits = { ...roomy, maxXmlPartBytes: 100 };
    const media = deflated('ppt/media/a.bin', new Uint8Array(500));
    expect(readZip(buildZip([media]), { limits }).entries).toHaveLength(1);
    const rels = deflated('_rels/.RELS', new Uint8Array(101));
    expect(() => readZip(buildZip([rels]), { limits })).toThrow(
      /^invalid ZIP: entry "_rels\/.RELS" declares 101 bytes, over the maxXmlPartBytes limit of 100$/,
    );
  });

  it('rejects XML parts over maxTotalXmlBytes', () => {
    const zip = buildZip(xmlEntries(3, 100));
    expect(() => readZip(zip, { limits: { ...roomy, maxTotalXmlBytes: 250 } })).toThrow(
      /^invalid ZIP: XML parts declare 300 bytes, over the maxTotalXmlBytes limit of 250$/,
    );
  });

  it('checks the limits before inflating any entry', () => {
    // The first entry is corrupt; inflating it would report that instead.
    const zip = buildZip([
      { name: 'bad.xml', method: 8, stored: new Uint8Array([0xff, 0xff, 0xff]), declaredSize: 4 },
      deflated('big.xml', new Uint8Array(1000)),
    ]);
    expect(() => readZip(zip)).toThrow(/"bad.xml" is not valid DEFLATE data/);
    expect(() => readZip(zip, { limits: { ...roomy, maxXmlPartBytes: 999 } })).toThrow(
      /"big.xml" declares 1000 bytes, over the maxXmlPartBytes limit/,
    );
  });

  it('stops mid-inflate when an entry within the limits inflates past its declared size', () => {
    // The tail is not valid DEFLATE, so reaching it would fail differently.
    const stored = concat([
      zeroBomb(40_000).subarray(0, 64 * 1024),
      new Uint8Array(1 << 20).fill(0xff),
    ]);
    const zip = buildZip([{ name: 'bomb.xml', method: 8, stored, declaredSize: 10 }]);
    expect(() => readZip(zip, { limits: roomy })).toThrow(
      /^invalid ZIP: entry "bomb.xml" inflates beyond its declared size of 10 bytes$/,
    );
  });

  it('rejects a ZIP64 locator, which it reads without limits', () => {
    const zip = buildZip([deflated('a.xml', utf8('<a/>'))], { zip64: true });
    expect(readZip(zip).entries.map((e) => e.name)).toEqual(['a.xml']);
    expect(() => readZip(zip, { limits: roomy })).toThrow(
      /^invalid ZIP: ZIP64 end of central directory locator is not accepted under load limits$/,
    );
  });

  it('rejects ZIP64 sizes in an entry, which it reads without limits', () => {
    const zip = buildZip([{ ...deflated('a.xml', utf8('<a/>')), zip64Sizes: true }]);
    expect(readZip(zip).entries[0]?.data).toEqual(utf8('<a/>'));
    expect(() => readZip(zip, { limits: roomy })).toThrow(
      /^invalid ZIP: entry "a.xml" uses ZIP64 sizes, which are not accepted under load limits$/,
    );
  });

  it('rejects a limit that is negative or NaN but accepts Infinity', () => {
    const zip = buildZip([deflated('a.xml', utf8('<a/>'))]);
    expect(() => readZip(zip, { limits: { ...roomy, maxEntries: Number.NaN } })).toThrow(
      RangeError,
    );
    expect(() => readZip(zip, { limits: { ...roomy, maxTotalXmlBytes: -1 } })).toThrow(
      /maxTotalXmlBytes must be a non-negative number/,
    );
    const unbounded = { ...roomy, maxTotalUncompressedBytes: Number.POSITIVE_INFINITY };
    expect(readZip(zip, { limits: unbounded }).entries).toHaveLength(1);
  });
});

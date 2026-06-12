// Minimal hand-rolled PNG builder for tests. Produces a single-color RGB
// image of the requested dimensions, so tests and sample generators can embed
// recognizable images without pulling in an image-encoding dependency.
//
// Node-only (uses node:zlib) — fine for the test/ tree, which never ships.

import { deflateRawSync } from 'node:zlib';

const adler32 = (data: Uint8Array): number => {
  let a = 1;
  let b = 0;
  for (const byte of data) {
    a = (a + byte) % 65521;
    b = (b + a) % 65521;
  }
  return (b << 16) | a;
};

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

const crc32 = (data: Uint8Array): number => {
  let c = 0xffffffff;
  for (const byte of data) c = CRC_TABLE[(c ^ byte) & 0xff]! ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
};

const concat = (...arrs: Uint8Array[]): Uint8Array => {
  const total = arrs.reduce((n, a) => n + a.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const a of arrs) {
    out.set(a, off);
    off += a.length;
  }
  return out;
};

const chunk = (typeStr: string, data: Uint8Array): Uint8Array => {
  const buf = new Uint8Array(8 + data.length + 4);
  const dv = new DataView(buf.buffer);
  dv.setUint32(0, data.length);
  for (let i = 0; i < 4; i++) buf[4 + i] = typeStr.charCodeAt(i);
  buf.set(data, 8);
  dv.setUint32(8 + data.length, crc32(buf.subarray(4, 8 + data.length)) >>> 0);
  return buf;
};

export const buildPng = (
  width: number,
  height: number,
  rgb: readonly [number, number, number],
): Uint8Array => {
  const SIG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

  const ihdr = new Uint8Array(13);
  const dv = new DataView(ihdr.buffer);
  dv.setUint32(0, width);
  dv.setUint32(4, height);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 2; // color type: RGB

  const rowSize = 1 + width * 3;
  const raw = new Uint8Array(rowSize * height);
  for (let y = 0; y < height; y++) {
    raw[y * rowSize] = 0; // filter: None
    for (let x = 0; x < width; x++) {
      const idx = y * rowSize + 1 + x * 3;
      raw[idx] = rgb[0];
      raw[idx + 1] = rgb[1];
      raw[idx + 2] = rgb[2];
    }
  }

  // PNG IDAT is zlib-wrapped DEFLATE: zlib header + deflated body + adler32.
  const deflated = deflateRawSync(raw);
  const idatBody = new Uint8Array(2 + deflated.length + 4);
  idatBody[0] = 0x78; // zlib CMF
  idatBody[1] = 0x9c; // zlib FLG
  idatBody.set(deflated, 2);
  new DataView(idatBody.buffer, idatBody.byteOffset + 2 + deflated.length, 4).setUint32(
    0,
    adler32(raw) >>> 0,
  );

  return concat(SIG, chunk('IHDR', ihdr), chunk('IDAT', idatBody), chunk('IEND', new Uint8Array()));
};

// Tiny PNG encoder (8-bit RGBA, no interlacing, filter type 0). We already
// get our own render as PNG straight from resvg; this exists only to turn the
// ground-truth raster (decoded from PPM) and the computed diff into something
// a browser can show in the HTML report — without pulling in `sharp` or
// `pngjs`. Node's zlib supplies the DEFLATE; this file is never on a browser
// path, so `node:zlib` is fine.

import { deflateSync } from 'node:zlib';
import type { RgbaImage } from './image.ts';

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xed_b8_83_20 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

const crc32 = (bytes: Uint8Array): number => {
  let c = 0xff_ff_ff_ff;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]!) & 0xff]! ^ (c >>> 8);
  return (c ^ 0xff_ff_ff_ff) >>> 0;
};

const chunk = (type: string, body: Uint8Array): Uint8Array => {
  const typeBytes = new Uint8Array([
    type.charCodeAt(0),
    type.charCodeAt(1),
    type.charCodeAt(2),
    type.charCodeAt(3),
  ]);
  const out = new Uint8Array(12 + body.length);
  const view = new DataView(out.buffer);
  view.setUint32(0, body.length);
  out.set(typeBytes, 4);
  out.set(body, 8);
  const crcInput = new Uint8Array(4 + body.length);
  crcInput.set(typeBytes, 0);
  crcInput.set(body, 4);
  view.setUint32(8 + body.length, crc32(crcInput));
  return out;
};

export const encodePng = (img: RgbaImage): Uint8Array => {
  const { width, height, data } = img;
  // Prepend a filter byte (0 = none) to each scanline before DEFLATE.
  const raw = new Uint8Array(height * (1 + width * 4));
  for (let y = 0; y < height; y++) {
    const src = y * width * 4;
    const dst = y * (1 + width * 4);
    raw[dst] = 0;
    raw.set(data.subarray(src, src + width * 4), dst + 1);
  }
  const idat = deflateSync(raw);

  const ihdr = new Uint8Array(13);
  const hv = new DataView(ihdr.buffer);
  hv.setUint32(0, width);
  hv.setUint32(4, height);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // color type RGBA
  // 10/11/12 = compression / filter / interlace, all 0.

  const sig = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
  const parts = [
    sig,
    chunk('IHDR', ihdr),
    chunk('IDAT', new Uint8Array(idat)),
    chunk('IEND', new Uint8Array(0)),
  ];
  const total = parts.reduce((n, p) => n + p.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const p of parts) {
    out.set(p, off);
    off += p.length;
  }
  return out;
};

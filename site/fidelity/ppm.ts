// Decoder for binary (P6) PPM, the default output of `pdftoppm`. We use PPM
// rather than `-png` so the harness needs no PNG decoder dependency: the
// pixel math reads PPM directly, and only the HTML report re-encodes to PNG
// (see png.ts).

import { rgbToRgba } from './image.ts';
import type { RgbaImage } from './image.ts';

const isWhitespace = (b: number): boolean => b === 32 || b === 9 || b === 10 || b === 13;

// Read the next ASCII integer token from `buf` starting at `pos`, skipping
// leading whitespace and `#`-comment lines (both legal in the PPM header).
const readToken = (buf: Uint8Array, pos: number): { value: number; next: number } => {
  let i = pos;
  for (;;) {
    while (i < buf.length && isWhitespace(buf[i]!)) i++;
    if (buf[i] === 0x23) {
      // comment to end of line
      while (i < buf.length && buf[i] !== 10) i++;
      continue;
    }
    break;
  }
  let value = 0;
  let seen = false;
  while (i < buf.length && buf[i]! >= 0x30 && buf[i]! <= 0x39) {
    value = value * 10 + (buf[i]! - 0x30);
    seen = true;
    i++;
  }
  if (!seen) throw new Error('PPM: expected an integer token in header');
  return { value, next: i };
};

export const decodePpm = (buf: Uint8Array): RgbaImage => {
  if (buf[0] !== 0x50 || buf[1] !== 0x36) {
    throw new Error('PPM: not a P6 file');
  }
  const w = readToken(buf, 2);
  const h = readToken(buf, w.next);
  const max = readToken(buf, h.next);
  if (max.value !== 255) throw new Error(`PPM: unsupported maxval ${max.value}`);
  // Exactly one whitespace byte separates the header from the raster.
  const start = max.next + 1;
  const expected = w.value * h.value * 3;
  const rgb = buf.subarray(start, start + expected);
  if (rgb.length < expected) {
    throw new Error(`PPM: truncated raster (${rgb.length} < ${expected})`);
  }
  return rgbToRgba(rgb, w.value, h.value);
};

// Generates the pptx-kit favicon + apple-touch-icon + logo PNGs.
//
// Runs as a plain Node script — no external image dependency. Output goes
// to `site/static/`. The mark is a rounded orange square with a stylised
// "P" cut out in white (a nod to PowerPoint without using its trademark).
//
// Run: `node scripts/build-brand-assets.mjs` (from `site/`).

import { writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { deflateRawSync } from 'node:zlib';

const STATIC_DIR = fileURLToPath(new URL('../static/', import.meta.url));

const ORANGE = [232, 80, 28]; // pptx-kit --accent
const ORANGE_HOT = [255, 106, 54]; // --accent-hot

// ---------------------------------------------------------------------------
// Drawing primitives — operate on a flat RGBA Uint8Array (one byte per
// channel, row-major top-down).

function makeCanvas(size) {
  // Start transparent.
  return { size, data: new Uint8Array(size * size * 4) };
}

function setPixel(canvas, x, y, rgb, alpha = 255) {
  if (x < 0 || y < 0 || x >= canvas.size || y >= canvas.size) return;
  const i = (y * canvas.size + x) * 4;
  canvas.data[i] = rgb[0];
  canvas.data[i + 1] = rgb[1];
  canvas.data[i + 2] = rgb[2];
  canvas.data[i + 3] = alpha;
}

// Fill a rounded-rectangle region with a vertical 2-color gradient.
function fillRoundedRect(canvas, x0, y0, x1, y1, r, top, bottom) {
  const h = y1 - y0;
  for (let y = y0; y < y1; y++) {
    const t = (y - y0) / Math.max(h - 1, 1);
    const rgb = [
      Math.round(top[0] + (bottom[0] - top[0]) * t),
      Math.round(top[1] + (bottom[1] - top[1]) * t),
      Math.round(top[2] + (bottom[2] - top[2]) * t),
    ];
    for (let x = x0; x < x1; x++) {
      // Distance from the rounded-corner centers.
      let inside = true;
      if (x < x0 + r && y < y0 + r) {
        const dx = x - (x0 + r);
        const dy = y - (y0 + r);
        if (dx * dx + dy * dy > r * r) inside = false;
      } else if (x >= x1 - r && y < y0 + r) {
        const dx = x - (x1 - r - 1);
        const dy = y - (y0 + r);
        if (dx * dx + dy * dy > r * r) inside = false;
      } else if (x < x0 + r && y >= y1 - r) {
        const dx = x - (x0 + r);
        const dy = y - (y1 - r - 1);
        if (dx * dx + dy * dy > r * r) inside = false;
      } else if (x >= x1 - r && y >= y1 - r) {
        const dx = x - (x1 - r - 1);
        const dy = y - (y1 - r - 1);
        if (dx * dx + dy * dy > r * r) inside = false;
      }
      if (inside) setPixel(canvas, x, y, rgb);
    }
  }
}

// Subtract a region from the canvas — paints alpha 0 in the matched
// pixels. Used to cut the "P" shape out of the orange tile.
function cutPShape(canvas) {
  const s = canvas.size;
  const stroke = Math.max(2, Math.round(s * 0.14));
  const left = Math.round(s * 0.3);
  const top = Math.round(s * 0.22);
  const bottom = Math.round(s * 0.8);
  // Bowl is a right-half ring centered on the stem's right edge.
  const bowlCenterX = left + Math.round(stroke / 2);
  const bowlTop = top;
  const bowlBottom = Math.round(s * 0.55);
  const bowlCenterY = Math.round((bowlTop + bowlBottom) / 2);
  const bowlOuterR = Math.round((bowlBottom - bowlTop) / 2);
  const bowlInnerR = Math.max(2, bowlOuterR - stroke);
  const bowlRightX = bowlCenterX + bowlOuterR;

  for (let y = 0; y < s; y++) {
    for (let x = 0; x < s; x++) {
      let cut = false;
      // Vertical stem.
      if (x >= left && x < left + stroke && y >= top && y < bottom) cut = true;
      // Bowl: right-half ring of the circle, hugging the stem.
      const dx = x - bowlCenterX;
      const dy = y - bowlCenterY;
      const dist2 = dx * dx + dy * dy;
      if (
        x >= bowlCenterX &&
        x <= bowlRightX &&
        dist2 <= bowlOuterR * bowlOuterR &&
        dist2 >= bowlInnerR * bowlInnerR
      ) {
        cut = true;
      }
      if (cut) {
        const i = (y * s + x) * 4;
        canvas.data[i + 3] = 0;
      }
    }
  }
}

// ---------------------------------------------------------------------------
// PNG encoder — RGBA with one filter byte per row, zlib-wrapped DEFLATE.

const SIG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(data) {
  let c = 0xffffffff;
  for (let i = 0; i < data.length; i++) c = CRC_TABLE[(c ^ data[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(typeStr, data) {
  const buf = new Uint8Array(8 + data.length + 4);
  const dv = new DataView(buf.buffer);
  dv.setUint32(0, data.length);
  for (let i = 0; i < 4; i++) buf[4 + i] = typeStr.charCodeAt(i);
  buf.set(data, 8);
  const crc = crc32(buf.subarray(4, 8 + data.length));
  dv.setUint32(8 + data.length, crc);
  return buf;
}

function adler32(data) {
  let a = 1;
  let b = 0;
  for (let i = 0; i < data.length; i++) {
    a = (a + data[i]) % 65521;
    b = (b + a) % 65521;
  }
  return (b << 16) | a;
}

function encodePng(canvas) {
  const { size, data } = canvas;
  const rowSize = 1 + size * 4;
  const raw = new Uint8Array(rowSize * size);
  for (let y = 0; y < size; y++) {
    raw[y * rowSize] = 0; // filter: None
    raw.set(data.subarray(y * size * 4, (y + 1) * size * 4), y * rowSize + 1);
  }

  const ihdr = new Uint8Array(13);
  const dv = new DataView(ihdr.buffer);
  dv.setUint32(0, size);
  dv.setUint32(4, size);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // color type: RGBA
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  const deflated = deflateRawSync(raw);
  const adler = adler32(raw);
  const idat = new Uint8Array(2 + deflated.length + 4);
  idat[0] = 0x78;
  idat[1] = 0x9c;
  idat.set(deflated, 2);
  new DataView(idat.buffer, 2 + deflated.length, 4).setUint32(0, adler >>> 0);

  const out = [SIG, chunk('IHDR', ihdr), chunk('IDAT', idat), chunk('IEND', new Uint8Array())];
  const total = out.reduce((n, a) => n + a.length, 0);
  const png = new Uint8Array(total);
  let off = 0;
  for (const a of out) {
    png.set(a, off);
    off += a.length;
  }
  return png;
}

// ---------------------------------------------------------------------------
// Build the four sizes the site ships.

function buildIcon(size, opts = {}) {
  const canvas = makeCanvas(size);
  const radius = Math.round(size * 0.22);
  const inset = Math.max(1, Math.round(size * 0.04));
  fillRoundedRect(canvas, inset, inset, size - inset, size - inset, radius, ORANGE_HOT, ORANGE);
  if (opts.withMark !== false) cutPShape(canvas);
  return encodePng(canvas);
}

const targets = [
  { name: 'logo.png', size: 256 },
  { name: 'apple-touch-icon.png', size: 180 },
  { name: 'favicon-32.png', size: 32 },
  { name: 'favicon-16.png', size: 16, withMark: false }, // mark blurs at 16px; ship a flat tile.
];

for (const { name, size, withMark } of targets) {
  const png = buildIcon(size, { withMark });
  writeFileSync(STATIC_DIR + name, png);
  // biome-ignore lint/suspicious/noConsole: build output.
  console.log(`wrote ${name} (${size}×${size}, ${png.length} bytes)`);
}

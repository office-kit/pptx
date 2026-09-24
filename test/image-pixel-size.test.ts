import { expect, it } from 'vitest';
import * as pptx from '../src/api/index.ts';
import { readImagePixelSize } from '../src/internal/opc/index.ts';

// Header-only fixtures exercise dimensions independently of pixel decoders.
const gif = (version = '89a') => {
  const b = new Uint8Array(13);
  b.set(new TextEncoder().encode(`GIF${version}`));
  const v = new DataView(b.buffer);
  v.setUint16(6, 640, true);
  v.setUint16(8, 320, true);
  return b;
};
const chunk = (tag: string, data: Uint8Array) => {
  const b = new Uint8Array(8 + data.length + (data.length % 2));
  b.set(new TextEncoder().encode(tag));
  new DataView(b.buffer).setUint32(4, data.length, true);
  b.set(data, 8);
  return b;
};
const riff = (...chunks: Uint8Array[]) => {
  const b = new Uint8Array(12 + chunks.reduce((n, c) => n + c.length, 0));
  b.set(new TextEncoder().encode('RIFF'));
  new DataView(b.buffer).setUint32(4, b.length - 8, true);
  b.set(new TextEncoder().encode('WEBP'), 8);
  let offset = 12;
  for (const c of chunks) {
    b.set(c, offset);
    offset += c.length;
  }
  return b;
};
const lossy = new Uint8Array([0, 0, 0, 0x9d, 1, 0x2a, 0x80, 0xc2, 0x40, 0x81]);
const lossless = new Uint8Array(5);
lossless[0] = 0x2f;
new DataView(lossless.buffer).setUint32(1, 639 | (319 << 14) | (1 << 28), true);
const extended = new Uint8Array([2, 0, 0, 0, 0x7f, 2, 0, 0x3f, 1, 0]);
const formats = [
  ['GIF87a', gif('87a')],
  ['GIF89a', gif()],
  ['VP8', riff(chunk('VP8 ', lossy))],
  ['VP8L', riff(chunk('VP8L', lossless))],
  ['VP8X', riff(chunk('VP8X', extended), chunk('ANMF', new Uint8Array(16)))],
  ['unknown padded chunk', riff(chunk('JUNK', new Uint8Array(3)), chunk('VP8L', lossless))],
] as const;

it.each(formats)('reads %s canvas dimensions from an offset byte view', (_, bytes) => {
  const padded = new Uint8Array(bytes.length + 12);
  padded.set(bytes, 7);
  expect(readImagePixelSize(padded.subarray(7, 7 + bytes.length))).toEqual({
    width: 640,
    height: 320,
  });
  for (let n = 0; n < (bytes[0] === 71 ? 13 : bytes.length); n++) {
    expect(readImagePixelSize(bytes.subarray(0, n))).toBeNull();
  }
});

it('rejects invalid GIF and WebP dimension headers', () => {
  const badGif = gif();
  badGif[6] = 0;
  badGif[7] = 0;
  const badSignature = lossy.slice();
  badSignature[3] = 0;
  const interFrame = lossy.slice();
  interFrame[0] = 1;
  const zeroWidth = lossy.slice();
  zeroWidth[6] = 0;
  zeroWidth[7] = 0;
  const badVersion = lossless.slice();
  badVersion[4]! |= 0x20;
  const badLossless = lossless.slice();
  badLossless[0] = 0;
  const overflow = extended.slice();
  overflow.fill(255, 4);
  const oversizedChunk = riff(chunk('VP8L', lossless));
  new DataView(oversizedChunk.buffer).setUint32(16, 0xffffffff, true);
  for (const b of [
    badGif,
    gif('88a'),
    riff(),
    oversizedChunk,
    riff(chunk('VP8 ', badSignature)),
    riff(chunk('VP8 ', interFrame)),
    riff(chunk('VP8 ', zeroWidth)),
    riff(chunk('VP8L', badVersion)),
    riff(chunk('VP8L', badLossless)),
    riff(chunk('VP8X', overflow)),
    riff(chunk('VP8X', extended.subarray(0, 9))),
  ]) {
    expect(readImagePixelSize(b)).toBeNull();
  }
});

it.each(formats)(
  '%s supports insertion, replacement, Fill/Fit and save/reload',
  async (_, bytes) => {
    const p = pptx.createPresentation(),
      s = pptx.addBlankSlide(p);
    const box = { x: pptx.emu(0), y: pptx.emu(0), w: pptx.inches(2), h: pptx.inches(2) };
    const pic = pptx.addSlideImage(s, bytes, { ...box, fit: 'contain' });
    expect(pptx.getShapeBounds(pic)).toEqual({ ...box, y: pptx.inches(0.5), h: pptx.inches(1) });
    pptx.setShapeBounds(pic, box);
    pptx.setShapeImage(pic, bytes, { fit: 'contain' });
    expect(pptx.getShapeBounds(pic)?.h).toBe(pptx.inches(1));
    pptx.setShapeBounds(pic, box);
    pptx.setShapeImageFit(pic, 'fill');
    expect(pptx.getShapeImageCrop(pic)).toEqual({ left: 0.25, right: 0.25, top: 0, bottom: 0 });
    pptx.setShapeImageFit(pic, 'fit');
    const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
    const restored = pptx.getSlideShapes(pptx.getSlides(saved)[0]!)[0]!;
    expect(pptx.getShapeBounds(restored)).toEqual(box);
    expect(pptx.getShapeImageCrop(restored)).toEqual({
      left: 0,
      right: 0,
      top: -0.5,
      bottom: -0.5,
    });
  },
);

import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  addSlide,
  findSlideLayout,
  getSlideSize,
  loadPresentation,
  inches,
  pt,
  setShapeImageFill,
  setShapeImageCrop,
  getShapeImageCrop,
  savePresentation,
  getSlides,
  getSlideShapes,
  setShapeImageFillLayout,
  getShapeImageIntrinsicSize,
  type ImageTileAlignment,
  type ImageTileFlip,
} from '../src/api/index.ts';
import { SHAPE_ELEMENT } from '../src/api/_internal-symbols.ts';
import { NS, attr, firstChildElement, qname } from '../src/internal/xml/index.ts';
import { readImageResolution } from '../src/internal/opc/image-format.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import { renderSlideToRgba } from '../packages/preview/src/node.ts';
import { buildPng } from './lib/build-png.ts';

async function fixture() {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/blank.pptx', import.meta.url)),
  );
  const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
  const shape = addSlideShape(slide, {
    preset: 'rect',
    x: inches(1),
    y: inches(1),
    w: inches(2),
    h: inches(2),
  });
  const png = buildPng(24, 24, (x, y) => [x < 12 ? 255 : 0, y < 12 ? 255 : 0, 0]);
  setShapeImageFill(shape, png);
  return { pres, slide, shape, png };
}

it('uses natural pixels at 96 DPI, then the explicit fill DPI override', async () => {
  const { shape } = await fixture();
  expect(getShapeImageIntrinsicSize(shape)).toEqual({ width: inches(0.25), height: inches(0.25) });
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml))!;
  const fill = firstChildElement(spPr, qname('a', 'blipFill', NS.dml))!;
  fill.attrs.push(attr(qname('', 'dpi', ''), '144'));
  expect(getShapeImageIntrinsicSize(shape)).toEqual({
    width: inches(1 / 6),
    height: inches(1 / 6),
  });
});

it('reads PNG physical resolution and rejects truncated chunks', async () => {
  const { png } = await fixture();
  const chunk = new Uint8Array(21);
  const view = new DataView(chunk.buffer);
  view.setUint32(0, 9);
  chunk.set(new TextEncoder().encode('pHYs'), 4);
  view.setUint32(8, 5669);
  view.setUint32(12, 11338);
  chunk[16] = 1;
  const bytes = new Uint8Array(png.length + chunk.length);
  bytes.set(png.subarray(0, 33));
  bytes.set(chunk, 33);
  bytes.set(png.subarray(33), 54);
  expect(readImageResolution(bytes)).toEqual({ x: 5669 * 0.0254, y: 11338 * 0.0254 });
  expect(readImageResolution(bytes.subarray(0, 50))).toBeNull();
  bytes[49] = 0;
  expect(readImageResolution(bytes)).toBeNull();
});

it('reads JFIF inch and centimeter densities', () => {
  const bytes = new Uint8Array([
    0xff, 0xd8, 0xff, 0xe0, 0, 16, 74, 70, 73, 70, 0, 1, 2, 1, 0, 144, 0, 72, 0, 0, 0xff, 0xd9,
  ]);
  expect(readImageResolution(bytes)).toEqual({ x: 144, y: 72 });
  bytes[13] = 2;
  expect(readImageResolution(bytes)).toEqual({ x: 144 * 2.54, y: 72 * 2.54 });
  bytes[13] = 0;
  expect(readImageResolution(bytes)).toBeNull();
  expect(readImageResolution(bytes.subarray(0, 16))).toBeNull();
});

describe.each(['none', 'x', 'y', 'xy'] as ImageTileFlip[])('tile mirror %s', (flip) => {
  it('repeats actual image pixels and reflects alternate tiles on the requested axes', async () => {
    const { pres, slide, shape } = await fixture();
    setShapeImageFillLayout(shape, { mode: 'tile', flip });
    const { image } = renderSlideToRgba(pres, slide, {
      width: Math.round(getSlideSize(pres)!.width / 9525),
    });
    const pixel = (x: number, y: number) =>
      Array.from(image.data.slice((y * image.width + x) * 4, (y * image.width + x) * 4 + 3));
    expect(pixel(102, 102)).toEqual([255, 255, 0]);
    expect(pixel(126, 102)).toEqual([flip === 'x' || flip === 'xy' ? 0 : 255, 255, 0]);
    expect(pixel(102, 126)).toEqual([255, flip === 'y' || flip === 'xy' ? 0 : 255, 0]);
    expect(pixel(150, 150)).toEqual([255, 255, 0]);
    expect(pixel(294, 150)).toEqual([255, 255, 255]);
  });
});

it.each([
  ['tl', 100, 104],
  ['t', 184, 104],
  ['tr', 268, 104],
  ['l', 100, 188],
  ['ctr', 184, 188],
  ['r', 268, 188],
  ['bl', 100, 272],
  ['b', 184, 272],
  ['br', 268, 272],
] as [ImageTileAlignment, number, number][])(
  'anchors %s before applying offsets',
  async (alignment, x, y) => {
    const { pres, slide, shape } = await fixture();
    setShapeImageFillLayout(shape, { mode: 'tile', alignment, offsetX: pt(3), offsetY: pt(6) });
    expect(renderSlideToSvg(pres, slide)).toContain(
      `patternUnits="userSpaceOnUse" x="${x.toFixed(2)}" y="${y.toFixed(2)}" width="24.00" height="24.00"`,
    );
  },
);

it('scales each tile independently and omits zero-sized tiles', async () => {
  const { pres, slide, shape } = await fixture();
  setShapeImageFillLayout(shape, { mode: 'tile', scaleX: 0.5, scaleY: 0.25 });
  expect(renderSlideToSvg(pres, slide)).toContain('width="12.00" height="6.00"');
  setShapeImageFillLayout(shape, { mode: 'tile', scaleX: 0 });
  expect(renderSlideToSvg(pres, slide)).not.toContain('<pattern');
});

it.each(['none', 'x', 'y', 'xy'] as const)(
  'crops each tile before %s mirroring and save/reload',
  async (flip) => {
    const { pres, slide, shape } = await fixture();
    setShapeImageCrop(shape, { left: 0.5 });
    setShapeImageFillLayout(shape, { mode: 'tile', flip });
    const restored = await loadPresentation(await savePresentation(pres));
    const savedSlide = getSlides(restored)[0]!;
    const savedShape = getSlideShapes(savedSlide).at(-1)!;
    expect(getShapeImageCrop(savedShape)).toEqual({ left: 0.5, top: 0, right: 0, bottom: 0 });
    for (const [presentation, target] of [
      [pres, slide],
      [restored, savedSlide],
    ] as const) {
      const { image } = renderSlideToRgba(presentation, target, { width: 960 });
      const pixel = (x: number, y: number) =>
        Array.from(image.data.slice((y * image.width + x) * 4, (y * image.width + x) * 4 + 3));
      expect(pixel(102, 102)).toEqual([0, 255, 0]);
      expect(pixel(114, 102)).toEqual([0, 255, 0]);
      expect(pixel(102, 114)).toEqual([0, 0, 0]);
      expect(pixel(102, 126)).toEqual([0, flip === 'y' || flip === 'xy' ? 0 : 255, 0]);
    }
  },
);

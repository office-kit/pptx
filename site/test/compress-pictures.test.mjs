import assert from 'node:assert/strict';
import test from 'node:test';
import {
  addBlankSlide,
  addSlideImage,
  addSlideShape,
  createPresentation,
  getShapeBounds,
  getShapeId,
  setShapeBounds,
  groupShapes,
  inches,
} from '@office-kit/pptx';
import {
  collectPictures,
  PICTURE_QUALITIES,
  planCompression,
} from '../src/lib/editor/core/compress-pictures.ts';

const INCH = 914400;
const plan = (input) => planCompression({ crop: null, ppi: null, deleteCropped: false, ...input });

test('quality choices follow Mac PowerPoint’s Compress Pictures sheet', () => {
  assert.deepEqual(
    PICTURE_QUALITIES.map((quality) => [quality.label, quality.ppi]),
    [
      ['High Fidelity (Maximum ppi)', null],
      ['HD (330 ppi)', 330],
      ['Print (220 ppi)', 220],
      ['On-screen (150 ppi)', 150],
      ['Email (96 ppi)', 96],
      ['Use Original Quality', null],
    ],
  );
});

test('downsamples to the target resolution of the frame and never upsamples', () => {
  // 2000 px shown 2 inches wide is 1000 ppi; Print keeps 220 ppi.
  const print = plan({
    pixelWidth: 2000,
    pixelHeight: 1000,
    frameWidth: 2 * INCH,
    frameHeight: INCH,
    ppi: 220,
  });
  assert.deepEqual(print, {
    source: { x: 0, y: 0, width: 2000, height: 1000 },
    width: 440,
    height: 220,
    crop: null,
  });
  // Already below the target: nothing to do.
  assert.equal(
    plan({
      pixelWidth: 100,
      pixelHeight: 100,
      frameWidth: 2 * INCH,
      frameHeight: 2 * INCH,
      ppi: 96,
    }),
    null,
  );
  // High Fidelity without cropped-area deletion keeps the picture.
  assert.equal(
    plan({ pixelWidth: 2000, pixelHeight: 1000, frameWidth: INCH, frameHeight: INCH }),
    null,
  );
  // A stretched frame keeps the denser axis at the target.
  const stretched = plan({
    pixelWidth: 1000,
    pixelHeight: 1000,
    frameWidth: 4 * INCH,
    frameHeight: INCH,
    ppi: 150,
  });
  assert.equal(stretched.width, 600);
  assert.equal(stretched.height, 600);
});

test('a crop counts only the visible pixels and deletion removes the rest', () => {
  const crop = { left: 0.25, top: 0, right: 0.25, bottom: 0.5 };
  // Visible part: 1000 × 500 px in a 2 × 1 inch frame = 500 ppi.
  const kept = plan({
    pixelWidth: 2000,
    pixelHeight: 1000,
    frameWidth: 2 * INCH,
    frameHeight: INCH,
    crop,
    ppi: 250,
  });
  assert.deepEqual(kept, {
    source: { x: 0, y: 0, width: 2000, height: 1000 },
    width: 1000,
    height: 500,
    crop,
  });
  const cut = plan({
    pixelWidth: 2000,
    pixelHeight: 1000,
    frameWidth: 2 * INCH,
    frameHeight: INCH,
    crop,
    ppi: 250,
    deleteCropped: true,
  });
  assert.deepEqual(cut, {
    source: { x: 500, y: 0, width: 1000, height: 500 },
    width: 500,
    height: 250,
    crop: null,
  });
  // Deleting crops alone (High Fidelity) still rewrites the picture.
  const only = plan({
    pixelWidth: 2000,
    pixelHeight: 1000,
    frameWidth: 2 * INCH,
    frameHeight: INCH,
    crop,
    deleteCropped: true,
  });
  assert.deepEqual(only, {
    source: { x: 500, y: 0, width: 1000, height: 500 },
    width: 1000,
    height: 500,
    crop: null,
  });
});

test('a negative crop (frame past the picture) survives deletion, rescaled', () => {
  const result = plan({
    pixelWidth: 1000,
    pixelHeight: 1000,
    frameWidth: INCH,
    frameHeight: INCH,
    crop: { left: 0.5, top: 0, right: -0.25, bottom: 0 },
    deleteCropped: true,
  });
  assert.deepEqual(result.source, { x: 500, y: 0, width: 500, height: 1000 });
  assert.deepEqual(result.crop, { left: 0, top: 0, right: -0.5, bottom: 0 });
});

test('collects pictures inside groups with the group scale', () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  // prettier-ignore
  const png = new Uint8Array([0x89,0x50,0x4e,0x47,0x0d,0x0a,0x1a,0x0a,0,0,0,0x0d,0x49,0x48,0x44,0x52,0,0,0,1,0,0,0,1,8,6,0,0,0,0x1f,0x15,0xc4,0x89,0,0,0,0x0d,0x49,0x44,0x41,0x54,0x78,0x9c,0x63,0,1,0,0,5,0,1,0x0d,0x0a,0x2d,0xb4,0,0,0,0,0x49,0x45,0x4e,0x44,0xae,0x42,0x60,0x82]);
  const loose = addSlideImage(slide, png, {
    x: inches(0),
    y: inches(0),
    w: inches(1),
    h: inches(1),
  });
  const grouped = addSlideImage(slide, png, {
    x: inches(2),
    y: inches(0),
    w: inches(1),
    h: inches(1),
  });
  const box = addSlideShape(slide, {
    preset: 'rect',
    x: inches(4),
    y: inches(0),
    w: inches(1),
    h: inches(1),
  });
  const group = groupShapes([grouped, box]);
  const bounds = getShapeBounds(group);
  setShapeBounds(group, { ...bounds, w: bounds.w * 2 });
  const pictures = collectPictures(pres);
  assert.deepEqual(
    pictures.map((item) => [item.ids, item.scaleX, item.scaleY]),
    [
      [[getShapeId(loose)], 1, 1],
      [[getShapeId(grouped), getShapeId(group)], 2, 1],
    ],
  );
});

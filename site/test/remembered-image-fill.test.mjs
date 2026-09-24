import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createPresentation,
  addBlankSlide,
  addSlideShape,
  inches,
  setShapeImageFill,
  setShapeImageFillLayout,
  getShapeImageFillLayout,
  setShapeImageOpacity,
  getShapeImageOpacity,
  setShapeImageCrop,
  getShapeImageCrop,
  getShapeImageFillBytes,
  setShapeFill,
  savePresentation,
  loadPresentation,
  getSlides,
  getSlideShapes,
} from '@office-kit/pptx';
import {
  switchRememberedImageLayout,
  readRememberedImageFill,
  restoreRememberedImageFill,
} from '../src/lib/editor/core/remembered-image-fill.ts';

test('switching away from a picture fill and back restores its media, placement, transparency and crop', async () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shape = addSlideShape(slide, {
    preset: 'rect',
    x: inches(1),
    y: inches(1),
    w: inches(2),
    h: inches(2),
  });
  const bytes = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
  setShapeImageFill(shape, bytes);
  setShapeImageFillLayout(shape, {
    mode: 'tile',
    scaleX: 0.6,
    scaleY: 0.8,
    alignment: 'br',
    flip: 'xy',
    rotateWithShape: false,
  });
  setShapeImageOpacity(shape, 0.65);
  setShapeImageCrop(shape, { left: 0.2, bottom: 0.1 });
  const remembered = readRememberedImageFill(shape);
  assert.ok(remembered);
  const layout = getShapeImageFillLayout(shape);
  const crop = getShapeImageCrop(shape);
  setShapeFill(shape, { color: 'accent1' });
  assert.equal(readRememberedImageFill(shape), undefined);
  restoreRememberedImageFill(shape, remembered);
  const loaded = await loadPresentation(await savePresentation(pres));
  const restored = getSlideShapes(getSlides(loaded)[0])[0];
  assert.deepEqual(getShapeImageFillBytes(restored), bytes);
  assert.deepEqual(getShapeImageFillLayout(restored), layout);
  assert.deepEqual(getShapeImageCrop(restored), crop);
  assert.equal(getShapeImageOpacity(restored), 0.65);
});

test('restoring remembered stretch fills preserves empty effects and offsets', () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shape = addSlideShape(slide, {
    preset: 'rect',
    x: inches(1),
    y: inches(1),
    w: inches(2),
    h: inches(2),
  });
  setShapeImageFill(shape, new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]));
  setShapeImageFillLayout(shape, { mode: 'stretch', left: 0.25, right: -0.2 });
  const remembered = readRememberedImageFill(shape);
  setShapeFill(shape, { color: 'accent2' });
  restoreRememberedImageFill(shape, remembered);
  assert.deepEqual(getShapeImageFillLayout(shape), remembered.layout);
  assert.equal(getShapeImageOpacity(shape), null);
  assert.equal(getShapeImageCrop(shape), null);
});

test('tile and stretch settings survive mode changes while rotation remains shared', () => {
  const remembered = {};
  const tile = {
    mode: 'tile',
    scaleX: 0.6,
    scaleY: 0.8,
    alignment: 'br',
    flip: 'xy',
    rotateWithShape: true,
  };
  const stretch = switchRememberedImageLayout(tile, 'stretch', remembered);
  assert.deepEqual(stretch, { mode: 'stretch', rotateWithShape: true });
  stretch.left = 0.25;
  stretch.rotateWithShape = false;
  const restoredTile = switchRememberedImageLayout(stretch, 'tile', remembered);
  assert.deepEqual(restoredTile, { ...tile, rotateWithShape: false });
  assert.deepEqual(switchRememberedImageLayout(restoredTile, 'stretch', remembered), stretch);
});

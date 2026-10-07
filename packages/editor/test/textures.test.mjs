import assert from 'node:assert/strict';
import test from 'node:test';
import {
  addBlankSlide,
  addSlideShape,
  createPresentation,
  getShapeImageFillLayout,
  getShapeImageIntrinsicSize,
  getShapeImageOpacity,
  inches,
  setShapeImageFill,
  setShapeImageFillLayout,
  setShapeImageOpacity,
} from '@office-kit/pptx';
import { TEXTURES, TEXTURE_SIZE, encodeTexturePng, texturePixels } from '../src/core/textures.ts';
import { TEXTURE_TILE_LAYOUT, insertRememberedTextureFill } from '../src/core/remembered-fill.ts';

const N = TEXTURE_SIZE;
const difference = (pixels, a, b) =>
  Math.abs(pixels[a] - pixels[b]) +
  Math.abs(pixels[a + 1] - pixels[b + 1]) +
  Math.abs(pixels[a + 2] - pixels[b + 2]);
/** Mean colour step between column x and x + 1 (wrapping), or row y and y + 1 when `rows`. */
const step = (pixels, at, rows) => {
  let total = 0;
  for (let i = 0; i < N; i++) {
    const [ax, ay, bx, by] = rows ? [i, at, i, (at + 1) % N] : [at, i, (at + 1) % N, i];
    total += difference(pixels, (ay * N + ax) * 4, (by * N + bx) * 4);
  }
  return total / N;
};

test('the gallery lists PowerPoint’s twenty-four textures in order', () => {
  assert.deepEqual(
    TEXTURES.map((texture) => texture.name),
    [
      'Papyrus',
      'Canvas',
      'Denim',
      'Woven mat',
      'Water droplets',
      'Paper bag',
      'Fish fossil',
      'Sand',
      'Green marble',
      'White marble',
      'Brown marble',
      'Granite',
      'Newsprint',
      'Recycled paper',
      'Parchment',
      'Stationery',
      'Blue tissue paper',
      'Pink tissue paper',
      'Purple mesh',
      'Bouquet',
      'Cork',
      'Walnut',
      'Oak',
      'Medium wood',
    ],
  );
});

for (const { id } of TEXTURES) {
  test(`${id} is a deterministic, opaque, seamless 128 × 128 tile`, () => {
    const pixels = texturePixels(id);
    assert.equal(pixels.length, 128 * 128 * 4);
    assert.deepEqual(texturePixels(id), pixels);
    for (let at = 3; at < pixels.length; at += 4) assert.equal(pixels[at], 255);
    for (const rows of [false, true]) {
      // The wrap (last column/row to the first) must be no harsher than the
      // harshest step inside the tile; a non-repeating texture shows a seam
      // there. The 10% allows for speckled textures, where the seam is just
      // another random step that may happen to be the largest.
      const interior = Array.from({ length: N - 1 }, (_, at) => step(pixels, at, rows));
      const seam = step(pixels, N - 1, rows);
      const limit = Math.max(...interior) * 1.1;
      assert.ok(seam <= limit, `${rows ? 'row' : 'column'} seam ${seam} > ${limit}`);
    }
  });
}

test('textures are distinct from each other', () => {
  const keys = new Set(TEXTURES.map(({ id }) => texturePixels(id).join()));
  assert.equal(keys.size, TEXTURES.length);
});

test('texture PNGs carry 128 px and ~144 DPI, and encode deterministically', async () => {
  const png = await encodeTexturePng('canvas');
  assert.deepEqual(await encodeTexturePng('canvas'), png);
  const view = new DataView(png.buffer, png.byteOffset);
  assert.equal(new TextDecoder().decode(png.subarray(12, 16)), 'IHDR');
  assert.equal(view.getUint32(16), 128);
  assert.equal(view.getUint32(20), 128);

  // The library reads the pHYs density: 128 px at 144 DPI is 64 pt.
  const slide = addBlankSlide(createPresentation());
  const shape = addSlideShape(slide, {
    preset: 'rect',
    x: inches(1),
    y: inches(1),
    w: inches(2),
    h: inches(2),
  });
  setShapeImageFill(shape, png);
  setShapeImageFillLayout(shape, { mode: 'tile' });
  const size = getShapeImageIntrinsicSize(shape);
  assert.ok(Math.abs(size.width - inches(128 / 144)) < 100);
  assert.ok(Math.abs(size.height - inches(128 / 144)) < 100);
});

test('a texture tiles with PowerPoint’s defaults, keeps transparency and remembers the stretch placement', async () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shape = addSlideShape(slide, {
    preset: 'rect',
    x: inches(1),
    y: inches(1),
    w: inches(2),
    h: inches(2),
  });
  setShapeImageFill(shape, await encodeTexturePng('sand'));
  setShapeImageFillLayout(shape, { mode: 'stretch', left: 0.25 });
  setShapeImageOpacity(shape, 0.6);
  const remembered = {};
  insertRememberedTextureFill(pres, shape, await encodeTexturePng('oak'), remembered);
  assert.deepEqual(getShapeImageFillLayout(shape), TEXTURE_TILE_LAYOUT);
  assert.equal(getShapeImageOpacity(shape), 0.6);
  assert.equal(remembered.imageLayouts.stretch.left, 0.25);
});

import assert from 'node:assert/strict';
import test from 'node:test';
import { BUILTIN_PICTURE_STYLES } from '@office-kit/pptx';
import { pictureStyleSwatches } from '../src/lib/editor/core/picture-style-swatches.ts';

const svg = (image) => decodeURIComponent(image.slice('data:image/svg+xml,'.length));

test('draws one swatch per built-in picture style, in gallery order, once', () => {
  const swatches = pictureStyleSwatches();
  assert.deepEqual([...swatches.keys()], [...BUILTIN_PICTURE_STYLES]);
  assert.equal(pictureStyleSwatches(), swatches);
  for (const image of swatches.values()) assert.match(svg(image), /<image /);
  // Each style draws differently: frames, shadows, reflections, 3-D.
  assert.equal(new Set(swatches.values()).size, BUILTIN_PICTURE_STYLES.length);
  assert.match(svg(swatches.get('Reflected Rounded Rectangle')), /data-pptx-reflection/);
  assert.match(svg(swatches.get('Rotated, White')), /transform="matrix\(/);
  assert.match(svg(swatches.get('Bevel Rectangle')), /feDiffuseLighting/);
});

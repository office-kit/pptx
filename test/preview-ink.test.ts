import { expect, it } from 'vitest';
import { createRequire } from 'node:module';
const { Resvg } = createRequire(new URL('../packages/preview/package.json', import.meta.url))(
  '@resvg/resvg-js',
);
import { renderInkToSvg } from '../packages/preview/src/ink.ts';
import { renderInkToPng } from '../packages/preview/src/node.ts';

it('paints round ink marks on a transparent background in the selected color', () => {
  const svg = renderInkToSvg([
    { points: [{ x: 100000, y: 200000 }], widthEmu: 95250, color: '#12ab34' },
  ]);
  const image = new Resvg(svg).render();
  expect([image.width, image.height]).toEqual([10, 10]);
  const rgba: Uint8Array = image.pixels;
  expect(Array.from(rgba.slice((5 * 10 + 5) * 4, (5 * 10 + 5) * 4 + 4))).toEqual([
    18, 171, 52, 255,
  ]);
  expect(rgba[3]).toBe(0);
});

it('bounds raster dimensions for tiny and extreme-aspect-ratio strokes', () => {
  for (const points of [
    [{ x: 0, y: 0 }],
    [
      { x: 0, y: 0 },
      { x: 1000000000, y: 0 },
    ],
    [
      { x: 0, y: 0 },
      { x: 0, y: 1000000000 },
    ],
  ]) {
    const png = Buffer.from(renderInkToPng([{ points, widthEmu: 1, color: '#FF0000' }]));
    expect(png.readUInt32BE(16)).toBeGreaterThanOrEqual(1);
    expect(png.readUInt32BE(20)).toBeGreaterThanOrEqual(1);
    expect(Math.max(png.readUInt32BE(16), png.readUInt32BE(20))).toBeLessThanOrEqual(2048);
  }
});

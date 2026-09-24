import { expect, it } from 'vitest';
import { renderSlideToRgba } from '../packages/preview/src/node.ts';
import { buildPng } from './lib/build-png.ts';
import {
  createPresentation,
  addBlankSlide,
  addSlideShape,
  addSlideLine,
  addSlideImage,
  setShapeImageCrop,
  setShapeImageOpacity,
  setShapeStrokeOpacity,
  inches,
  setShapeFill,
  setShapeNoFill,
  setShapeNoStroke,
  setShapeStroke,
  setShapeStrokeCompound,
} from '../src/api/index.ts';

it.each(['shape', 'connector', 'picture'])(
  'double and triple %s outlines expose the actual background through their gaps',
  (geometry) => {
    const p = createPresentation();
    const slide = addBlankSlide(p);
    const bg = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(10),
      h: inches(7.5),
    });
    setShapeFill(bg, '#0000FF');
    const shape =
      geometry === 'connector'
        ? addSlideLine(slide, {
            from: { x: inches(1), y: inches(1) },
            to: { x: inches(4), y: inches(1) },
          })
        : geometry === 'picture'
          ? addSlideImage(slide, buildPng(2, 2, [0, 0, 255]), {
              x: inches(1),
              y: inches(1),
              w: inches(3),
              h: inches(2),
            })
          : addSlideShape(slide, {
              preset: 'rect',
              x: inches(1),
              y: inches(1),
              w: inches(3),
              h: inches(2),
            });
    setShapeNoFill(shape);
    if (geometry === 'picture') {
      setShapeImageCrop(shape, { top: 0.25, left: 0.25 });
      setShapeImageOpacity(shape, 0.25);
    }
    setShapeStroke(shape, { color: '#FF0000', widthEmu: 285750 });
    for (const [kind, samples] of [
      [
        'dbl',
        [
          [86, true],
          [96, false],
          [106, true],
        ],
      ],
      [
        'tri',
        [
          [84, true],
          [87, false],
          [90, false],
          [92, true],
          [96, true],
          [99, true],
          [102, false],
          [104, false],
          [108, true],
        ],
      ],
    ] as const) {
      setShapeStrokeCompound(shape, kind);
      const { image } = renderSlideToRgba(p, slide);
      const pixel = (y: number) =>
        Array.from(image.data.slice((y * image.width + 200) * 4, (y * image.width + 200) * 4 + 4));
      for (const [y, red] of samples)
        expect(pixel(y), `${kind} at y=${y}`).toEqual(red ? [255, 0, 0, 255] : [0, 0, 255, 255]);
    }
    setShapeStrokeCompound(shape, 'dbl');
    setShapeStrokeOpacity(shape, 0.5);
    const { image } = renderSlideToRgba(p, slide);
    const offset = (86 * image.width + 200) * 4;
    expect(image.data[offset]).toBeGreaterThanOrEqual(127);
    expect(image.data[offset]).toBeLessThanOrEqual(128);
    expect(image.data[offset + 2]).toBeGreaterThanOrEqual(127);
    expect(image.data[offset + 2]).toBeLessThanOrEqual(128);
  },
);

it('picture outlines retain solid/asymmetric bands and can be removed', () => {
  const p = createPresentation();
  const slide = addBlankSlide(p);
  const bg = addSlideShape(slide, {
    preset: 'rect',
    x: inches(0),
    y: inches(0),
    w: inches(10),
    h: inches(7.5),
  });
  setShapeFill(bg, '#0000FF');
  const picture = addSlideImage(slide, buildPng(2, 2, [0, 255, 0]), {
    x: inches(1),
    y: inches(1),
    w: inches(3),
    h: inches(2),
  });
  setShapeImageCrop(picture, { top: 0.25, right: 0.25 });
  setShapeStroke(picture, { color: '#FF0000', widthEmu: 304800 });
  for (const compound of ['sng', 'thickThin', 'thinThick', 'none'] as const) {
    if (compound === 'none') setShapeNoStroke(picture);
    else setShapeStrokeCompound(picture, compound);
    const { image } = renderSlideToRgba(p, slide);
    const pixel = (y: number) =>
      Array.from(image.data.slice((y * image.width + 240) * 4, (y * image.width + 240) * 4 + 4));
    expect(pixel(91), `${compound} outside`).toEqual(
      compound === 'none' || compound === 'thinThick' ? [0, 0, 255, 255] : [255, 0, 0, 255],
    );
    expect(pixel(100), `${compound} inside`).toEqual(
      compound === 'none' || compound === 'thickThin' ? [0, 255, 0, 255] : [255, 0, 0, 255],
    );
    expect(pixel(120)).toEqual([0, 255, 0, 255]);
  }
});

it.each(['rect', 'ellipse'] as const)(
  'asymmetric %s outlines place the thick band on the requested side',
  (preset) => {
    const p = createPresentation();
    const slide = addBlankSlide(p);
    const bg = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(10),
      h: inches(7.5),
    });
    setShapeFill(bg, '#0000FF');
    const shape = addSlideShape(slide, {
      preset,
      x: inches(1),
      y: inches(1),
      w: inches(3),
      h: inches(2),
    });
    setShapeStroke(shape, { color: '#FF0000', widthEmu: 304800 });
    for (const filled of [false, true]) {
      if (filled) setShapeFill(shape, '#00FF00');
      else setShapeNoFill(shape);
      for (const compound of ['thickThin', 'thinThick'] as const) {
        setShapeStrokeCompound(shape, compound);
        const { image } = renderSlideToRgba(p, slide);
        const pixel = (y: number) =>
          Array.from(
            image.data.slice((y * image.width + 240) * 4, (y * image.width + 240) * 4 + 4),
          );
        const red = [255, 0, 0, 255];
        const blue = [0, 0, 255, 255];
        const interior = filled ? [0, 255, 0, 255] : blue;
        expect(pixel(83), `${compound} outer edge`).toEqual(red);
        expect(pixel(91), `${compound} outside centerline`).toEqual(
          compound === 'thickThin' ? red : blue,
        );
        expect(pixel(100), `${compound} inside centerline`).toEqual(
          compound === 'thickThin' ? interior : red,
        );
        expect(pixel(109), `${compound} inner edge`).toEqual(red);
        expect(pixel(120), `${compound} fill`).toEqual(interior);
      }
    }
  },
);

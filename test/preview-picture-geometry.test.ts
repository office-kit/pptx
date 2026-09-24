import { expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideImage,
  addSlideShape,
  createPresentation,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeFill,
  setShapeImageCrop,
  setShapeImageCropShape,
  setShapeImageFill,
  setShapeNoStroke,
  setShapeStroke,
  setShapeRotation,
} from '../src/api/index.ts';
import { renderSlideToRgba } from '../packages/preview/src/node.ts';
import { buildPng } from './lib/build-png.ts';

it.each([
  ['ellipse', 0],
  ['roundRect', 0],
  ['ellipse', 90],
  ['roundRect', 90],
] as const)(
  '%s at %s degrees clips pictures and image fills to their outline',
  async (preset, rotation) => {
    for (const kind of ['picture', 'image-fill']) {
      let p = createPresentation();
      let slide = addBlankSlide(p);
      const bg = addSlideShape(slide, {
        preset: 'rect',
        x: inches(0),
        y: inches(0),
        w: inches(10),
        h: inches(7.5),
      });
      setShapeFill(bg, '#0000FF');
      const bounds = { x: inches(1), y: inches(1), w: inches(3), h: inches(2) };
      const png = buildPng(2, 2, [0, 255, 0]);
      const shape =
        kind === 'picture'
          ? addSlideImage(slide, png, bounds)
          : addSlideShape(slide, { ...bounds, preset });
      if (kind === 'picture') setShapeImageCrop(shape, { top: 0.25, left: 0.25 });
      else setShapeImageFill(shape, png);
      setShapeStroke(shape, { color: '#FF0000', widthEmu: 38100 });
      setShapeRotation(shape, rotation);
      if (kind === 'picture') {
        setShapeImageCropShape(shape, preset);
        p = await loadPresentation(await savePresentation(p));
        slide = getSlides(p)[0]!;
      }
      const { image } = renderSlideToRgba(p, slide);
      const offset = (x: number, y: number, width: number) => {
        const px = rotation === 90 ? 240 - (y - 192) : x;
        const py = rotation === 90 ? 192 + (x - 240) : y;
        return (py * width + px) * 4;
      };
      const pixel = (x: number, y: number) => {
        const start = offset(x, y, image.width);
        return Array.from(image.data.slice(start, start + 4));
      };
      expect(pixel(98, 98), `${kind} outside ${preset}`).toEqual([0, 0, 255, 255]);
      expect(pixel(240, 96), `${kind} top outline`).toEqual([255, 0, 0, 255]);
      expect(pixel(240, 192), `${kind} image center`).toEqual([0, 255, 0, 255]);
      if (kind === 'image-fill') {
        setShapeNoStroke(shape);
        const { image: noBorder } = renderSlideToRgba(p, slide);
        const corner = offset(98, 98, noBorder.width);
        expect(Array.from(noBorder.data.slice(corner, corner + 4))).toEqual([0, 0, 255, 255]);
      }
    }
  },
);

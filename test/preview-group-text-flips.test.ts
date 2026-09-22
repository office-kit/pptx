import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideShape,
  createPresentation,
  getGroupChildren,
  getSlides,
  getSlideShapes,
  groupShapes,
  inches,
  loadPresentation,
  savePresentation,
  setShapeFlip,
  setShapeRotation,
  ungroupShapes,
} from '../src/api/index.ts';
import { renderSlideSvg } from '../packages/preview/src/render-slide.ts';

type Matrix = [number, number, number, number, number, number];
const identity = (): Matrix => [1, 0, 0, 1, 0, 0];
const multiply = (a: Matrix, b: Matrix): Matrix => [
  a[0] * b[0] + a[2] * b[1],
  a[1] * b[0] + a[3] * b[1],
  a[0] * b[2] + a[2] * b[3],
  a[1] * b[2] + a[3] * b[3],
  a[0] * b[4] + a[2] * b[5] + a[4],
  a[1] * b[4] + a[3] * b[5] + a[5],
];
const translation = (x: number, y: number): Matrix => [1, 0, 0, 1, x, y];

// Evaluate the SVG transforms at the text box, independently of the renderer.
const textCorners = (svg: string): number[] => {
  const stack: Matrix[] = [identity()];
  for (const [tag] of svg.matchAll(/<\/?g\b[^>]*>|<foreignObject\b[^>]*>/g)) {
    if (tag.startsWith('</g')) {
      stack.pop();
      continue;
    }
    let matrix = stack.at(-1)!;
    const transform = /transform="([^"]*)"/.exec(tag)?.[1] ?? '';
    for (const [, op, args] of transform.matchAll(/(\w+)\(([^)]*)\)/g)) {
      const [x = 0, y = 0, z = 0] = args!.split(/[ ,]+/).map(Number);
      let next: Matrix;
      if (op === 'translate') next = translation(x, y);
      else if (op === 'scale') next = [x, 0, 0, y, 0, 0];
      else if (op === 'rotate') {
        const radians = (x * Math.PI) / 180;
        const c = Math.cos(radians),
          s = Math.sin(radians);
        next = multiply(multiply(translation(y, z), [c, s, -s, c, 0, 0]), translation(-y, -z));
      } else throw new Error(`Unexpected transform: ${op}`);
      matrix = multiply(matrix, next);
    }
    if (tag.startsWith('<g')) {
      stack.push(matrix);
      continue;
    }
    const attr = (name: string) => Number(new RegExp(`${name}="([^"]+)"`).exec(tag)![1]);
    const x = attr('x'),
      y = attr('y'),
      w = attr('width'),
      h = attr('height');
    return [
      [x, y],
      [x + w, y],
      [x, y + h],
    ].flatMap(([px, py]) => [
      matrix[0] * px! + matrix[2] * py! + matrix[4],
      matrix[1] * px! + matrix[3] * py! + matrix[5],
    ]);
  }
  throw new Error('No text box rendered');
};

const flips = [
  { horizontal: false, vertical: false },
  { horizontal: true, vertical: false },
  { horizontal: false, vertical: true },
  { horizontal: true, vertical: true },
];

describe('text orientation in flipped groups', () => {
  for (const outerFlip of flips)
    for (const innerFlip of flips) {
      it(`matches ungrouped text for outer ${JSON.stringify(outerFlip)}, inner ${JSON.stringify(innerFlip)}`, async () => {
        const pres = createPresentation();
        const slide = addBlankSlide(pres);
        const shape = (x: number, text?: string) =>
          addSlideShape(slide, {
            preset: 'rect',
            x: inches(x),
            y: inches(2),
            w: inches(2),
            h: inches(1),
            ...(text === undefined ? {} : { text }),
          });
        const label = shape(1, '日本語 English');
        setShapeRotation(label, 31);
        setShapeFlip(label, { vertical: true });
        const inner = groupShapes([label, shape(4)]);
        setShapeRotation(inner, 23);
        setShapeFlip(inner, innerFlip);
        const outer = groupShapes([inner, shape(7)]);
        setShapeRotation(outer, 47);
        setShapeFlip(outer, outerFlip);
        const loaded = await loadPresentation(await savePresentation(pres));
        const loadedSlide = getSlides(loaded)[0]!;
        const grouped = textCorners(renderSlideSvg(loaded, loadedSlide));
        const loadedOuter = getSlideShapes(loadedSlide)[0]!;
        const loadedInner = getGroupChildren(loadedOuter)[0]!;
        ungroupShapes(loadedOuter);
        ungroupShapes(loadedInner);
        const ungrouped = textCorners(renderSlideSvg(loaded, loadedSlide));
        grouped.forEach((coordinate, index) =>
          expect(Math.abs(coordinate - ungrouped[index]!)).toBeLessThan(0.03),
        );
      });
    }
});

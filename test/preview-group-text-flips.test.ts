import { describe, expect, it } from 'vitest';
import {
  type ChartSpec,
  addBlankSlide,
  addSlideShape,
  addSlideChart,
  addSlideTable,
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
import { attrsOf } from './lib/svg-query.ts';
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
const allTextCorners = (svg: string): number[][] => {
  const corners: number[][] = [];
  const stack: Matrix[] = [identity()];
  for (const [tag] of svg.matchAll(/<\/?g\b[^>]*>|<(?:foreignObject|text)\b[^>]*>/g)) {
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
      w = tag.startsWith('<text') ? 1 : attr('width'),
      h = tag.startsWith('<text') ? 1 : attr('height');
    corners.push(
      [
        [x, y],
        [x + w, y],
        [x, y + h],
      ].flatMap(([px, py]) => [
        matrix[0] * px! + matrix[2] * py! + matrix[4],
        matrix[1] * px! + matrix[3] * py! + matrix[5],
      ]),
    );
  }
  if (corners.length === 0) throw new Error('No text box rendered');
  return corners;
};
const textCorners = (svg: string): number[] => allTextCorners(svg)[0]!;

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

describe('table text orientation in flipped groups', () => {
  for (const textLayout of ['foreignObject', 'svg'] as const)
    for (const outerFlip of flips)
      for (const tableFlip of flips) {
        it(`${textLayout}: readable table text with group ${JSON.stringify(outerFlip)}, table ${JSON.stringify(tableFlip)}`, async () => {
          const pres = createPresentation();
          const slide = addBlankSlide(pres);
          const table = addSlideTable(slide, {
            x: inches(1),
            y: inches(2),
            w: inches(4),
            h: inches(2),
            rows: [
              ['日本語 English', 'B'],
              ['C', 'D'],
            ],
          });
          setShapeRotation(table, 31);
          setShapeFlip(table, tableFlip);
          const sibling = addSlideShape(slide, {
            preset: 'rect',
            x: inches(7),
            y: inches(2),
            w: inches(1),
            h: inches(1),
          });
          const group = groupShapes([table, sibling]);
          setShapeRotation(group, 47);
          setShapeFlip(group, outerFlip);
          const loaded = await loadPresentation(await savePresentation(pres));
          const loadedSlide = getSlides(loaded)[0]!;
          const grouped = textCorners(renderSlideSvg(loaded, loadedSlide, { textLayout }));
          const [x, y, rightX, rightY, bottomX, bottomY] = grouped as [
            number,
            number,
            number,
            number,
            number,
            number,
          ];
          // Positive signed area means the rendered glyph axes are not mirrored.
          expect((rightX - x) * (bottomY - y) - (rightY - y) * (bottomX - x)).toBeGreaterThan(0);
          ungroupShapes(getSlideShapes(loadedSlide)[0]!);
          const ungrouped = textCorners(renderSlideSvg(loaded, loadedSlide, { textLayout }));
          grouped.forEach((coordinate, index) =>
            expect(Math.abs(coordinate - ungrouped[index]!)).toBeLessThan(0.03),
          );
        });
      }
});

// One spec per chart kind. `ChartSpec` is a union, and a pie has no axes to
// title, so the axis labels ride only with the kinds that plot them.
const labelledSpec = (
  kind: 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'bubble' | 'radar',
): ChartSpec => {
  const common = {
    title: '日本語 English',
    categories: ['A', 'B'],
    legend: { position: 'r' as const },
  };
  const axes = {
    valueAxisTitle: '値 Value',
    categoryAxisTitle: '分類 Category',
    categoryAxisLabelRotationDeg: 30,
  };
  const series = { name: '系列 Series', values: [2, 4] };
  if (kind === 'scatter')
    return { ...common, ...axes, kind, series: [{ ...series, xValues: [1, 2] }] };
  if (kind === 'bubble')
    return {
      ...common,
      ...axes,
      kind,
      series: [{ ...series, xValues: [1, 2], bubbleSizes: [3, 4] }],
    };
  if (kind === 'pie') return { ...common, kind, series: [series] };
  if (kind === 'doughnut') return { ...common, kind, series: [series] };
  if (kind === 'radar') return { ...common, ...axes, kind, series: [series] };
  return { ...common, ...axes, kind, series: [series] };
};

describe('chart label orientation in flipped groups', () => {
  for (const kind of [
    'column',
    'bar',
    'line',
    'area',
    'pie',
    'doughnut',
    'scatter',
    'bubble',
    'radar',
  ] as const)
    for (const groupFlip of flips)
      for (const chartFlip of flips) {
        it(`${kind}: group ${JSON.stringify(groupFlip)}, chart ${JSON.stringify(chartFlip)}`, async () => {
          const pres = createPresentation();
          const slide = addBlankSlide(pres);
          const chart = addSlideChart(slide, {
            x: inches(1),
            y: inches(1),
            w: inches(5),
            h: inches(4),
            spec: labelledSpec(kind),
          });
          setShapeRotation(chart, 31);
          setShapeFlip(chart, chartFlip);
          const sibling = addSlideShape(slide, {
            preset: 'rect',
            x: inches(7),
            y: inches(1),
            w: inches(1),
            h: inches(1),
          });
          const group = groupShapes([chart, sibling]);
          setShapeRotation(group, 47);
          setShapeFlip(group, groupFlip);
          const loaded = await loadPresentation(await savePresentation(pres));
          const loadedSlide = getSlides(loaded)[0]!;
          const groupedSvg = renderSlideSvg(loaded, loadedSlide);
          const reflected =
            (groupFlip.horizontal !== groupFlip.vertical) !==
            (chartFlip.horizontal !== chartFlip.vertical);
          expect(attrsOf(groupedSvg, 'text').at(-1)?.['text-anchor'] ?? 'start').toBe(
            reflected ? 'end' : 'start',
          );
          const grouped = allTextCorners(groupedSvg);
          expect(grouped.length).toBeGreaterThan(1);
          for (const coords of grouped) {
            const [x, y, rx, ry, bx, by] = coords as [
              number,
              number,
              number,
              number,
              number,
              number,
            ];
            expect((rx - x) * (by - y) - (ry - y) * (bx - x)).toBeGreaterThan(0);
          }
          ungroupShapes(getSlideShapes(loadedSlide)[0]!);
          const ungrouped = allTextCorners(renderSlideSvg(loaded, loadedSlide));
          expect(ungrouped).toHaveLength(grouped.length);
          grouped.forEach((coords, i) =>
            coords.forEach((coordinate, j) =>
              expect(Math.abs(coordinate - ungrouped[i]![j]!)).toBeLessThan(0.03),
            ),
          );
        });
      }
});

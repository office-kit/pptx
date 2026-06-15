// Corpus of parity cases. Each case authors the *same* visual slide twice —
// once through PptxGenJS (the battle-tested reference whose output opens
// cleanly in PowerPoint) and once through pptx-kit — so the harness can diff
// the two drawing trees and turn "is our output good?" into a number.
//
// `pgjs` receives a fresh PptxGenJS slide. `kit` receives a fresh pptx-kit
// presentation plus a blank slide already added to it; author onto `slide`.

import type { PresentationData, SlideData } from '../../src/api/_internal-symbols.ts';
import type { PgjsDeck, PgjsSlide } from './pptxgenjs-types.ts';
import {
  addSlideChart,
  addSlideImage,
  addSlideLine,
  addSlideShape,
  addSlideTable,
  addSlideTextBox,
  getTableCell,
  inches,
  pt,
  setParagraphAlignment,
  setShapeFill,
  setShapeBullets,
  setShapeRunFormat,
  setShapeRunHyperlink,
  setShapeStroke,
  setShapeStrokeArrow,
  setShapeStrokeDash,
  setTableCellAlignment,
  setTableCellFill,
  setTableCellTextFormat,
} from '../../src/api/index.ts';
import { buildPng } from '../lib/build-png.ts';

// A 1x1 transparent-ish PNG reused so image cases compare identical bytes.
const PNG = buildPng(64, 48, [40, 90, 180]);

export interface CorpusCase {
  id: string;
  /** PptxGenJS authoring. `slide` is a PptxGenJS Slide; `pptx` the deck. */
  pgjs: (slide: PgjsSlide, pptx: PgjsDeck) => void;
  /** pptx-kit authoring onto the already-added blank `slide`. */
  kit: (pres: PresentationData, slide: SlideData) => void;
  /** Bytes the image case feeds to PptxGenJS (data URI) and pptx-kit. */
  png?: Uint8Array;
}

const dataUri = (png: Uint8Array): string =>
  `data:image/png;base64,${Buffer.from(png).toString('base64')}`;

export const CASES: CorpusCase[] = [
  {
    id: 'text-plain',
    pgjs: (s) => {
      s.addText('Hello World', { x: 1, y: 1, w: 4, h: 1 });
    },
    kit: (_p, slide) => {
      addSlideTextBox(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(1),
        text: 'Hello World',
      });
    },
  },
  {
    id: 'text-formatted',
    pgjs: (s) => {
      s.addText('Bold Red', {
        x: 1,
        y: 1,
        w: 4,
        h: 1,
        bold: true,
        fontSize: 24,
        color: 'C00000',
        fontFace: 'Calibri',
      });
    },
    kit: (_p, slide) => {
      const box = addSlideTextBox(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(1),
        text: 'Bold Red',
      });
      setShapeRunFormat(box, 0, 0, { bold: true, size: 24, color: '#C00000', font: 'Calibri' });
    },
  },
  {
    id: 'text-align-center',
    pgjs: (s) => {
      s.addText('Centered', { x: 1, y: 1, w: 6, h: 1, align: 'center' });
    },
    kit: (_p, slide) => {
      const box = addSlideTextBox(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(6),
        h: inches(1),
        text: 'Centered',
      });
      setParagraphAlignment(box, 0, 'ctr');
    },
  },
  {
    id: 'shape-rect-fill',
    pgjs: (s) => {
      s.addShape('rect', { x: 1, y: 1, w: 3, h: 2, fill: { color: '00B050' } });
    },
    kit: (_p, slide) => {
      const sh = addSlideShape(slide, {
        preset: 'rect',
        x: inches(1),
        y: inches(1),
        w: inches(3),
        h: inches(2),
      });
      setShapeFill(sh, '#00B050');
    },
  },
  {
    id: 'shape-roundrect-stroke',
    pgjs: (s) => {
      s.addShape('roundRect', {
        x: 1,
        y: 1,
        w: 3,
        h: 2,
        fill: { color: 'FFFFFF' },
        line: { color: '1F4E79', width: 2 },
      });
    },
    kit: (_p, slide) => {
      const sh = addSlideShape(slide, {
        preset: 'roundRect',
        x: inches(1),
        y: inches(1),
        w: inches(3),
        h: inches(2),
      });
      setShapeFill(sh, '#FFFFFF');
      setShapeStroke(sh, { color: '#1F4E79', widthEmu: pt(2) });
    },
  },
  {
    id: 'line-plain',
    pgjs: (s) => {
      s.addShape('line', { x: 1, y: 2, w: 5, h: 0, line: { color: '222222', width: 2 } });
    },
    kit: (_p, slide) => {
      const ln = addSlideLine(slide, {
        from: { x: inches(1), y: inches(2) },
        to: { x: inches(6), y: inches(2) },
      });
      setShapeStroke(ln, { color: '#222222', widthEmu: pt(2) });
    },
  },
  {
    id: 'table-basic',
    pgjs: (s) => {
      s.addTable(
        [
          ['A', 'B'],
          ['1', '2'],
        ],
        { x: 1, y: 1, w: 4, colW: [2, 2] },
      );
    },
    kit: (_p, slide) => {
      // PptxGenJS auto-sizes rows to ~0.5in each (no rowH given); request a
      // matching total so the graphic-frame extent and per-row heights line up.
      const table = addSlideTable(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(1),
        rows: [
          ['A', 'B'],
          ['1', '2'],
        ],
      });
      // Match PptxGenJS's 12pt cell-text default so the runs compare equal.
      for (let r = 0; r < 2; r++) {
        for (let col = 0; col < 2; col++) {
          setTableCellTextFormat(getTableCell(table, r, col), { size: 12 });
        }
      }
    },
  },
  {
    id: 'shape-ellipse',
    pgjs: (s) => {
      s.addShape('ellipse', { x: 2, y: 1, w: 3, h: 2, fill: { color: 'C00000' } });
    },
    kit: (_p, slide) => {
      const sh = addSlideShape(slide, {
        preset: 'ellipse',
        x: inches(2),
        y: inches(1),
        w: inches(3),
        h: inches(2),
      });
      setShapeFill(sh, '#C00000');
    },
  },
  {
    id: 'shape-triangle',
    pgjs: (s) => {
      s.addShape('triangle', { x: 2, y: 1, w: 3, h: 2, fill: { color: '548235' } });
    },
    kit: (_p, slide) => {
      const sh = addSlideShape(slide, {
        preset: 'triangle',
        x: inches(2),
        y: inches(1),
        w: inches(3),
        h: inches(2),
      });
      setShapeFill(sh, '#548235');
    },
  },
  {
    id: 'shape-text',
    pgjs: (s) => {
      s.addText('Label', {
        shape: 'roundRect',
        x: 1,
        y: 1,
        w: 4,
        h: 1.5,
        fill: { color: '2E75B6' },
      });
    },
    kit: (_p, slide) => {
      const sh = addSlideShape(slide, {
        preset: 'roundRect',
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(1.5),
        text: 'Label',
      });
      setShapeFill(sh, '#2E75B6');
    },
  },
  {
    id: 'text-italic',
    pgjs: (s) => {
      s.addText('Italic', { x: 1, y: 1, w: 4, h: 1, italic: true });
    },
    kit: (_p, slide) => {
      const box = addSlideTextBox(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(1),
        text: 'Italic',
      });
      setShapeRunFormat(box, 0, 0, { italic: true });
    },
  },
  {
    id: 'text-underline',
    pgjs: (s) => {
      s.addText('Underline', { x: 1, y: 1, w: 4, h: 1, underline: { style: 'sng' } });
    },
    kit: (_p, slide) => {
      const box = addSlideTextBox(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(1),
        text: 'Underline',
      });
      setShapeRunFormat(box, 0, 0, { underline: true });
    },
  },
  {
    id: 'text-multiline',
    pgjs: (s) => {
      s.addText('First\nSecond', { x: 1, y: 1, w: 6, h: 2 });
    },
    kit: (_p, slide) => {
      addSlideTextBox(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(6),
        h: inches(2),
        text: 'First\nSecond',
      });
    },
  },
  {
    id: 'line-dashed',
    pgjs: (s) => {
      s.addShape('line', {
        x: 1,
        y: 2,
        w: 5,
        h: 0,
        line: { color: '1F4E79', width: 2, dashType: 'dash' },
      });
    },
    kit: (_p, slide) => {
      const ln = addSlideLine(slide, {
        from: { x: inches(1), y: inches(2) },
        to: { x: inches(6), y: inches(2) },
      });
      setShapeStroke(ln, { color: '#1F4E79', widthEmu: pt(2) });
      setShapeStrokeDash(ln, 'dash');
    },
  },
  {
    id: 'line-arrow',
    pgjs: (s) => {
      s.addShape('line', {
        x: 1,
        y: 2,
        w: 5,
        h: 0,
        line: { color: 'C00000', width: 3, endArrowType: 'triangle' },
      });
    },
    kit: (_p, slide) => {
      const ln = addSlideLine(slide, {
        from: { x: inches(1), y: inches(2) },
        to: { x: inches(6), y: inches(2) },
      });
      setShapeStroke(ln, { color: '#C00000', widthEmu: pt(3) });
      // PptxGenJS `endArrowType` is the line's END point — OOXML `<a:tailEnd>`.
      setShapeStrokeArrow(ln, 'tail', { type: 'triangle' });
    },
  },
  {
    id: 'image-large',
    png: PNG,
    pgjs: (s) => {
      s.addImage({ data: dataUri(PNG), x: 0.5, y: 0.5, w: 6, h: 4.5 });
    },
    kit: (_p, slide) => {
      addSlideImage(slide, PNG, { x: inches(0.5), y: inches(0.5), w: inches(6), h: inches(4.5) });
    },
  },
  {
    id: 'chart-column',
    pgjs: (s) => {
      s.addChart('bar', [{ name: 'Revenue', labels: ['Q1', 'Q2'], values: [120, 180] }], {
        x: 1,
        y: 1,
        w: 6,
        h: 4,
      });
    },
    kit: (_p, slide) => {
      addSlideChart(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(6),
        h: inches(4),
        spec: {
          kind: 'column',
          categories: ['Q1', 'Q2'],
          series: [{ name: 'Revenue', values: [120, 180] }],
        },
      });
    },
  },
  {
    id: 'chart-line',
    pgjs: (s) => {
      s.addChart(
        'line',
        [{ name: 'Visits', labels: ['Mon', 'Tue', 'Wed'], values: [10, 22, 18] }],
        { x: 1, y: 1, w: 6, h: 4 },
      );
    },
    kit: (_p, slide) => {
      addSlideChart(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(6),
        h: inches(4),
        spec: {
          kind: 'line',
          categories: ['Mon', 'Tue', 'Wed'],
          series: [{ name: 'Visits', values: [10, 22, 18] }],
        },
      });
    },
  },
  {
    id: 'chart-pie',
    pgjs: (s) => {
      s.addChart(
        'pie',
        [{ name: 'Share', labels: ['Web', 'Mobile', 'Desktop'], values: [55, 30, 15] }],
        { x: 1, y: 1, w: 6, h: 4 },
      );
    },
    kit: (_p, slide) => {
      addSlideChart(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(6),
        h: inches(4),
        spec: {
          kind: 'pie',
          categories: ['Web', 'Mobile', 'Desktop'],
          series: [{ name: 'Share', values: [55, 30, 15] }],
        },
      });
    },
  },
  {
    id: 'chart-multiseries',
    pgjs: (s) => {
      s.addChart(
        'bar',
        [
          { name: 'A', labels: ['Q1', 'Q2', 'Q3'], values: [4, 8, 5] },
          { name: 'B', labels: ['Q1', 'Q2', 'Q3'], values: [2, 5, 9] },
        ],
        { x: 1, y: 1, w: 6, h: 4 },
      );
    },
    kit: (_p, slide) => {
      addSlideChart(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(6),
        h: inches(4),
        spec: {
          kind: 'column',
          categories: ['Q1', 'Q2', 'Q3'],
          series: [
            { name: 'A', values: [4, 8, 5] },
            { name: 'B', values: [2, 5, 9] },
          ],
        },
      });
    },
  },
  {
    id: 'chart-bar-horizontal',
    pgjs: (s) => {
      s.addChart('bar', [{ name: 'S', labels: ['A', 'B', 'C'], values: [3, 7, 5] }], {
        x: 1,
        y: 1,
        w: 6,
        h: 4,
        barDir: 'bar',
      });
    },
    kit: (_p, slide) => {
      addSlideChart(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(6),
        h: inches(4),
        spec: {
          kind: 'bar',
          categories: ['A', 'B', 'C'],
          series: [{ name: 'S', values: [3, 7, 5] }],
        },
      });
    },
  },
  {
    id: 'chart-area',
    pgjs: (s) => {
      s.addChart('area', [{ name: 'S', labels: ['A', 'B', 'C'], values: [3, 7, 5] }], {
        x: 1,
        y: 1,
        w: 6,
        h: 4,
      });
    },
    kit: (_p, slide) => {
      addSlideChart(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(6),
        h: inches(4),
        spec: {
          kind: 'area',
          categories: ['A', 'B', 'C'],
          series: [{ name: 'S', values: [3, 7, 5] }],
        },
      });
    },
  },
  {
    id: 'chart-doughnut',
    pgjs: (s) => {
      s.addChart(
        'doughnut',
        [{ name: 'Share', labels: ['Web', 'Mobile', 'Desktop'], values: [55, 30, 15] }],
        {
          x: 1,
          y: 1,
          w: 6,
          h: 4,
        },
      );
    },
    kit: (_p, slide) => {
      addSlideChart(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(6),
        h: inches(4),
        spec: {
          kind: 'doughnut',
          categories: ['Web', 'Mobile', 'Desktop'],
          series: [{ name: 'Share', values: [55, 30, 15] }],
        },
      });
    },
  },
  {
    id: 'text-hyperlink',
    pgjs: (s) => {
      s.addText('Open', {
        x: 1,
        y: 1,
        w: 4,
        h: 1,
        hyperlink: { url: 'https://example.com/' },
      });
    },
    kit: (_p, slide) => {
      const box = addSlideTextBox(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(1),
        text: 'Open',
      });
      // PptxGenJS auto-underlines linked text and mirrors the link onto the
      // shape (cNvPr); pptx-kit links the runs (which makes the text
      // clickable). Author the underline to match; the shape-level mirror is
      // folded by the comparator.
      setShapeRunHyperlink(box, 0, 0, 'https://example.com/');
      setShapeRunFormat(box, 0, 0, { underline: true });
    },
  },
  {
    id: 'text-strike',
    pgjs: (s) => {
      s.addText('Struck', { x: 1, y: 1, w: 4, h: 1, strike: true });
    },
    kit: (_p, slide) => {
      const box = addSlideTextBox(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(1),
        text: 'Struck',
      });
      setShapeRunFormat(box, 0, 0, { strike: true });
    },
  },
  {
    id: 'shape-diamond',
    pgjs: (s) => {
      s.addShape('diamond', { x: 2, y: 1, w: 3, h: 2, fill: { color: '8064A2' } });
    },
    kit: (_p, slide) => {
      const sh = addSlideShape(slide, {
        preset: 'diamond',
        x: inches(2),
        y: inches(1),
        w: inches(3),
        h: inches(2),
      });
      setShapeFill(sh, '#8064A2');
    },
  },
  {
    id: 'table-cell-fill',
    pgjs: (s) => {
      s.addTable(
        [
          [{ text: 'A', options: { fill: { color: 'FFFF00' } } }, 'B'],
          ['1', '2'],
        ] as unknown as string[][],
        { x: 1, y: 1, w: 4, colW: [2, 2] },
      );
    },
    kit: (_p, slide) => {
      const table = addSlideTable(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(1),
        rows: [
          ['A', 'B'],
          ['1', '2'],
        ],
      });
      setTableCellFill(getTableCell(table, 0, 0), '#FFFF00');
      for (let r = 0; r < 2; r++) {
        for (let col = 0; col < 2; col++) {
          setTableCellTextFormat(getTableCell(table, r, col), { size: 12 });
        }
      }
    },
  },
  {
    id: 'text-bullets',
    pgjs: (s) => {
      s.addText('One\nTwo\nThree', { x: 1, y: 1, w: 5, h: 3, bullet: true });
    },
    kit: (_p, slide) => {
      const box = addSlideTextBox(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(5),
        h: inches(3),
        text: 'One\nTwo\nThree',
      });
      setShapeBullets(box, 'bullet');
    },
  },
  {
    id: 'shape-star5',
    pgjs: (s) => {
      s.addShape('star5', { x: 2, y: 1, w: 3, h: 2, fill: { color: 'FFD966' } });
    },
    kit: (_p, slide) => {
      const sh = addSlideShape(slide, {
        preset: 'star5',
        x: inches(2),
        y: inches(1),
        w: inches(3),
        h: inches(2),
      });
      setShapeFill(sh, '#FFD966');
    },
  },
  {
    id: 'shape-pentagon',
    pgjs: (s) => {
      s.addShape('pentagon', { x: 2, y: 1, w: 3, h: 2, fill: { color: '70AD47' } });
    },
    kit: (_p, slide) => {
      const sh = addSlideShape(slide, {
        preset: 'pentagon',
        x: inches(2),
        y: inches(1),
        w: inches(3),
        h: inches(2),
      });
      setShapeFill(sh, '#70AD47');
    },
  },
  {
    id: 'table-cell-align',
    pgjs: (s) => {
      s.addTable(
        [
          [{ text: 'A', options: { align: 'center' } }, 'B'],
          ['1', '2'],
        ] as unknown as string[][],
        { x: 1, y: 1, w: 4, colW: [2, 2] },
      );
    },
    kit: (_p, slide) => {
      const table = addSlideTable(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(1),
        rows: [
          ['A', 'B'],
          ['1', '2'],
        ],
      });
      setTableCellAlignment(getTableCell(table, 0, 0), 'ctr');
      for (let r = 0; r < 2; r++) {
        for (let col = 0; col < 2; col++) {
          setTableCellTextFormat(getTableCell(table, r, col), { size: 12 });
        }
      }
    },
  },
  {
    id: 'image-png',
    png: PNG,
    pgjs: (s) => {
      s.addImage({ data: dataUri(PNG), x: 1, y: 1, w: 2, h: 1.5 });
    },
    kit: (_p, slide) => {
      addSlideImage(slide, PNG, { x: inches(1), y: inches(1), w: inches(2), h: inches(1.5) });
    },
  },
];

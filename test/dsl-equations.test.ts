// Equations reach TSX through `paragraphs`, the core `ParagraphSpec` that
// `Text` and table cells already take; no separate element.

import { expect, it } from 'vitest';
import * as api from '../src/api/index.ts';
import { Presentation, Slide, Table, Text, compile } from '../packages/dsl/src/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

const t = (text: string): api.MathNodeSpec => ({ kind: 'text', text });
const logistic: api.MathNodeSpecs = [
  t('σ(z)='),
  {
    kind: 'fraction',
    numerator: [t('1')],
    denominator: [t('1+'), { kind: 'superscript', base: [t('e')], sup: [t('−z')] }],
  },
];

it('writes equations from Text and table-cell paragraphs at the element format', async () => {
  const tree = Presentation({
    children: Slide({
      children: [
        Text({
          x: 1,
          y: 1,
          width: 8,
          height: 1,
          size: 28,
          color: '#1F4E79',
          paragraphs: [{ runs: [{ text: 'Logistic: ' }, { equation: logistic }] }],
        }),
        Table({
          x: 1,
          y: 3,
          width: 6,
          height: 1,
          rows: [['Loss', { paragraphs: [{ runs: [{ equation: [t('L='), t('∑')] }] }] }]],
        }),
      ],
    }),
  });
  const pres = await api.loadPresentation(await api.savePresentation(await compile(tree)));
  const slide = api.getSlides(pres)[0]!;
  const [text, table] = api.getSlideShapes(slide);
  const [, math] = api.getShapeParagraphElements(text!, 0);
  if (math?.kind !== 'math') throw new Error('expected an equation');
  expect(math.text).toBe('σ(z)=11+e−z');
  expect(math.math[0]).toMatchObject({
    format: { size: 28, color: '#1F4E79', font: 'Cambria Math' },
  });
  const cell = api.getTableCells(table!)[0]![1]!;
  expect(api.getTableCellParagraphs(cell)[0]!.elements.map((e) => e.kind)).toEqual(['math']);
  expect(renderSlideToSvg(pres, slide).match(/data-pptx-math=""/g)).toHaveLength(2);
});

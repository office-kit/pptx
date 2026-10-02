import assert from 'node:assert/strict';
import test from 'node:test';
import {
  addBlankSlide,
  addSlideTable,
  createPresentation,
  getTableCellParagraphs,
  getTableCells,
  setPresentationTheme,
  setTableCellTextFormat,
  inches,
} from '@office-kit/pptx';
import { resolveEditingTextColor } from '../src/lib/editor/core/inline-text-html.ts';

test('table editing resolves scheme run colors through the slide color map', () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const table = addSlideTable(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(1),
    rows: [['Text']],
  });
  const cell = getTableCells(table)[0][0];
  setPresentationTheme(pres, { dark1: '#123456' });
  setTableCellTextFormat(cell, { color: 'scheme:tx1', underline: 'dotted' });

  const rawColor = getTableCellParagraphs(cell)[0].elements[0].format.color;
  assert.equal(rawColor, 'tx1');
  assert.equal(resolveEditingTextColor(pres, table, rawColor), '#123456');
  assert.equal(resolveEditingTextColor(pres, table, '#123456'), '#123456');
});

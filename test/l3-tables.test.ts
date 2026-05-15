// Table authoring via graphicFrame.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTable,
  findSlideLayout,
  getShapeKind,
  getShapePosition,
  getShapeSize,
  getSlideXmlString,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

describe('L3: addSlideTable', () => {
  it('emits a graphicFrame with a:tbl matching rows × cols', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    const tbl = addSlideTable(slide, {
      x: inches(1), y: inches(1), w: inches(8), h: inches(3),
      rows: [
        ['Name', 'Score', 'Notes'],
        ['Alice', '42', 'first'],
        ['Bob', '7', 'second'],
      ],
    });
    expect(getShapeKind(tbl)).toBe('graphicFrame');
    expect(getShapePosition(tbl)).toEqual({ x: inches(1), y: inches(1) });
    expect(getShapeSize(tbl)).toEqual({ w: inches(8), h: inches(3) });

    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).toContain('<a:tbl>');
    expect(xml).toContain('<a:tblPr');
    expect(xml).toContain('<a:tblGrid>');
    expect((xml.match(/<a:gridCol/g) ?? []).length).toBe(3);
    expect((xml.match(/<a:tr /g) ?? []).length).toBe(3);
    expect((xml.match(/<a:tc>/g) ?? []).length).toBe(9);
    expect(xml).toContain('Alice');
    expect(xml).toContain('Bob');
    expect(xml).toContain('Score');
  });

  it('respects custom column widths and row heights', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    addSlideTable(slide, {
      x: inches(1), y: inches(1), w: inches(6), h: inches(2),
      colWidths: [inches(1), inches(2), inches(3)],
      rowHeights: [inches(0.5), inches(1.5)],
      rows: [['A', 'B', 'C'], ['1', '2', '3']],
    });
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).toMatch(/<a:gridCol w="914400"\/>/);
    expect(xml).toMatch(/<a:gridCol w="1828800"\/>/);
    expect(xml).toMatch(/<a:gridCol w="2743200"\/>/);
  });

  it('rejects ragged rows', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    expect(() =>
      addSlideTable(slide, {
        x: inches(0), y: inches(0), w: inches(4), h: inches(2),
        rows: [['a', 'b'], ['c']],
      }),
    ).toThrow(/row 1 has 1 cells; expected 2/);
  });

  it('rejects empty input', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    expect(() =>
      addSlideTable(slide, {
        x: inches(0), y: inches(0), w: inches(4), h: inches(2), rows: [],
      }),
    ).toThrow(/at least one row/);
    expect(() =>
      addSlideTable(slide, {
        x: inches(0), y: inches(0), w: inches(4), h: inches(2), rows: [[]],
      }),
    ).toThrow(/at least one column/);
  });

  skipIfNoXmllint('table-bearing slide validates against pml.xsd', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    addSlideTable(slide, {
      x: inches(1), y: inches(1), w: inches(8), h: inches(3),
      rows: [['H1', 'H2'], ['r1c1', 'r1c2'], ['r2c1', 'r2c2']],
    });
    expectSchemaValid(getSlideXmlString(getSlides(pres).at(-1)!), 'pml');
  });

  it('table content round-trips through save/reload', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    addSlideTable(slide, {
      x: inches(1), y: inches(1), w: inches(6), h: inches(2),
      rows: [['k', 'v'], ['a', '1'], ['b', '2']],
    });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const xml = getSlideXmlString(getSlides(reloaded).at(-1)!);
    for (const expected of ['k', 'v', 'a', '1', 'b', '2']) {
      expect(xml).toContain(expected);
    }
  });
});

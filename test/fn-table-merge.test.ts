// mergeTableCells — the write counterpart to getTableCellSpan. Merges a
// rectangular block into the top-left anchor (gridSpan / rowSpan) and
// marks the covered cells hMerge / vMerge per ECMA-376 §21.1.3.18.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  _internalPackageOf,
  addSlideTable,
  getSlides,
  getSlideTables,
  getTableCell,
  getTableCellSpan,
  inches,
  loadPresentation,
  mergeTableCells,
  savePresentation,
} from '../src/api/index.ts';
import { partName } from '../src/internal/opc/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const decode = (b: Uint8Array): string => new TextDecoder().decode(b);
const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

const buildTable = (slide: ReturnType<typeof getSlides>[number]) =>
  addSlideTable(slide, {
    rows: [
      ['a', 'b', 'c'],
      ['d', 'e', 'f'],
      ['g', 'h', 'i'],
    ],
    x: inches(0),
    y: inches(0),
    w: inches(6),
    h: inches(3),
  });

describe('fn API: mergeTableCells', () => {
  it('merges a horizontal 1×2 block (gridSpan on anchor, hMerge on cover)', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tbl = buildTable(slide);
    mergeTableCells(tbl, { row: 0, col: 0, rowSpan: 1, colSpan: 2 });

    const anchor = getTableCellSpan(getTableCell(tbl, 0, 0));
    expect(anchor.gridSpan).toBe(2);
    expect(anchor.rowSpan).toBe(1);
    expect(anchor.hMerge).toBe(false);

    const covered = getTableCellSpan(getTableCell(tbl, 0, 1));
    expect(covered.hMerge).toBe(true);
    expect(covered.vMerge).toBe(false);
  });

  it('merges a vertical 2×1 block (rowSpan on anchor, vMerge on cover)', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tbl = buildTable(slide);
    mergeTableCells(tbl, { row: 0, col: 0, rowSpan: 2, colSpan: 1 });

    const anchor = getTableCellSpan(getTableCell(tbl, 0, 0));
    expect(anchor.rowSpan).toBe(2);
    expect(anchor.gridSpan).toBe(1);

    const covered = getTableCellSpan(getTableCell(tbl, 1, 0));
    expect(covered.vMerge).toBe(true);
    expect(covered.hMerge).toBe(false);
  });

  it('merges a 2×2 block with both markers on the bottom-right cover', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tbl = buildTable(slide);
    mergeTableCells(tbl, { row: 1, col: 1, rowSpan: 2, colSpan: 2 });

    const anchor = getTableCellSpan(getTableCell(tbl, 1, 1));
    expect(anchor.gridSpan).toBe(2);
    expect(anchor.rowSpan).toBe(2);

    // top-right cover of the anchor row → hMerge only.
    const topRight = getTableCellSpan(getTableCell(tbl, 1, 2));
    expect(topRight.hMerge).toBe(true);
    expect(topRight.vMerge).toBe(false);

    // bottom-left cover → vMerge only.
    const bottomLeft = getTableCellSpan(getTableCell(tbl, 2, 1));
    expect(bottomLeft.vMerge).toBe(true);
    expect(bottomLeft.hMerge).toBe(false);

    // bottom-right cover → both.
    const bottomRight = getTableCellSpan(getTableCell(tbl, 2, 2));
    expect(bottomRight.hMerge).toBe(true);
    expect(bottomRight.vMerge).toBe(true);
  });

  it('round-trips a merge through save → load', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tbl = buildTable(slide);
    mergeTableCells(tbl, { row: 0, col: 0, rowSpan: 2, colSpan: 2 });

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reloadedTable = getSlideTables(getSlides(reloaded)[0]!)[0]!;
    const anchor = getTableCellSpan(getTableCell(reloadedTable, 0, 0));
    expect(anchor.gridSpan).toBe(2);
    expect(anchor.rowSpan).toBe(2);
    expect(getTableCellSpan(getTableCell(reloadedTable, 1, 1)).vMerge).toBe(true);
  });

  it('rejects out-of-range blocks', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tbl = buildTable(slide);
    expect(() => mergeTableCells(tbl, { row: 2, col: 2, rowSpan: 2, colSpan: 1 })).toThrow(
      /exceed/,
    );
    expect(() => mergeTableCells(tbl, { row: 0, col: 2, rowSpan: 1, colSpan: 2 })).toThrow(
      /exceed/,
    );
  });

  it('rejects a 1×1 "merge" and non-positive spans', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tbl = buildTable(slide);
    expect(() => mergeTableCells(tbl, { row: 0, col: 0, rowSpan: 1, colSpan: 1 })).toThrow(
      /1×1 block is not a merge/,
    );
    expect(() => mergeTableCells(tbl, { row: 0, col: 0, rowSpan: 0, colSpan: 2 })).toThrow(/≥ 1/);
  });

  it('rejects overlapping merges and leaves the table untouched', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tbl = buildTable(slide);
    mergeTableCells(tbl, { row: 0, col: 0, rowSpan: 1, colSpan: 2 });
    // (0,1) is now hMerge — a second merge touching it must be rejected.
    expect(() => mergeTableCells(tbl, { row: 0, col: 1, rowSpan: 2, colSpan: 1 })).toThrow(
      /already part of a merge/,
    );
    // The rejected call must not have set rowSpan on (0,1).
    expect(getTableCellSpan(getTableCell(tbl, 0, 1)).rowSpan).toBe(1);
  });

  skipIfNoXmllint('a slide with a merged table validates against pml.xsd', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tbl = buildTable(slide);
    mergeTableCells(tbl, { row: 1, col: 1, rowSpan: 2, colSpan: 2 });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const pkg = _internalPackageOf(reloaded);
    const slidePart = pkg.getPart(partName('/ppt/slides/slide1.xml'));
    expect(slidePart).not.toBeNull();
    expectSchemaValid(decode(slidePart!.data), 'pml');
  });
});

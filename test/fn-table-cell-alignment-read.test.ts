// getTableCellAlignment — read back a cell's first-paragraph algn.
// Counterpart to setTableCellAlignment.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  _internalPackageOf,
  addSlideTable,
  getSlidePartName,
  getSlideTables,
  getSlides,
  getTableCell,
  getTableCellAlignment,
  inches,
  loadPresentation,
  savePresentation,
  setTableCellAlignment,
} from '../src/api/index.ts';
import { partName } from '../src/internal/opc/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getTableCellAlignment', () => {
  it('returns null on a freshly-built cell', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addSlideTable(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      rows: [['a', 'b']],
    });
    const cell = getTableCell(table, 0, 0)!;
    expect(getTableCellAlignment(cell)).toBeNull();
  });

  it('reflects each ParagraphAlignment value', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addSlideTable(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      rows: [['a', 'b', 'c', 'd']],
    });

    const c0 = getTableCell(table, 0, 0)!;
    setTableCellAlignment(c0, 'l');
    expect(getTableCellAlignment(c0)).toBe('l');

    const c1 = getTableCell(table, 0, 1)!;
    setTableCellAlignment(c1, 'ctr');
    expect(getTableCellAlignment(c1)).toBe('ctr');

    const c2 = getTableCell(table, 0, 2)!;
    setTableCellAlignment(c2, 'r');
    expect(getTableCellAlignment(c2)).toBe('r');

    const c3 = getTableCell(table, 0, 3)!;
    setTableCellAlignment(c3, 'just');
    expect(getTableCellAlignment(c3)).toBe('just');
  });

  it('reads a friendly name back as its spec token, and types it that way', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const table = addSlideTable(getSlides(pres)[0]!, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      rows: [['a']],
    });
    const cell = getTableCell(table, 0, 0)!;
    setTableCellAlignment(cell, 'center');
    const align = getTableCellAlignment(cell);
    // @ts-expect-error -- 'center' is what you set; 'ctr' is what you read.
    expect(align === 'center').toBe(false);
    expect(align).toBe('ctr');
  });

  it('reads an algn value outside ST_TextAlignType as unset', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addSlideTable(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      rows: [['a']],
    });
    setTableCellAlignment(getTableCell(table, 0, 0)!, 'center');

    // Corrupt the attribute the way a hand-edited or third-party file could.
    const saved = await loadPresentation(await savePresentation(pres));
    const part = _internalPackageOf(saved).getPart(partName(getSlidePartName(slide)))!;
    const xml = new TextDecoder().decode(part.data);
    expect(xml).toContain('algn="ctr"');
    part.data = new TextEncoder().encode(xml.replace('algn="ctr"', 'algn="bogus"'));

    const reloaded = await loadPresentation(await savePresentation(saved));
    const [reloadedTable] = getSlideTables(getSlides(reloaded)[0]!);
    expect(getTableCellAlignment(getTableCell(reloadedTable!, 0, 0)!)).toBeNull();
  });
});

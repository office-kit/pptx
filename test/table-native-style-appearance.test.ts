import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTable,
  createPresentation,
  getSlides,
  getSlideShapes,
  getTableCells,
  getTableCellAppearanceEffective,
  inches,
  loadPresentation,
  savePresentation,
  setTableStyleFlags,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';

it('resolves the Medium Style 2 definition saved by Mac PowerPoint', async () => {
  // Captured from Mac PowerPoint tableStyles.xml; validated against dml-main.xsd.
  const style = await readFile(new URL('./fixtures/native-table-style.xml', import.meta.url));
  const original = createPresentation();
  const table = addSlideTable(addBlankSlide(original), {
    x: inches(1),
    y: inches(1),
    w: inches(6),
    h: inches(3),
    rows: [
      ['A', 'B', 'C'],
      ['D', 'E', 'F'],
      ['G', 'H', 'I'],
    ],
  });
  setTableStyleFlags(table, { firstRow: true, bandRow: true });
  const { entries } = readZip(await savePresentation(original));
  const pres = await loadPresentation(
    writeZip(
      entries.map((entry) =>
        entry.name === 'ppt/tableStyles.xml' ? { ...entry, data: style } : entry,
      ),
    ),
  );
  const cells = getTableCells(getSlideShapes(getSlides(pres)[0]!)[0]!);
  const header = getTableCellAppearanceEffective(pres, cells[0]![1]!);
  const band = getTableCellAppearanceEffective(pres, cells[1]![1]!);
  const body = getTableCellAppearanceEffective(pres, cells[2]![1]!);
  expect(header.fill).toEqual({ kind: 'solid', color: '#4F81BD' });
  expect(band.fill.kind).toBe('solid');
  expect(body.fill.kind).toBe('solid');
  expect(header.fill).not.toEqual(band.fill);
  expect(band.fill).not.toEqual(body.fill);
  expect(header.borders.bottom).toMatchObject({ color: '#FFFFFF', widthEmu: 38100 });
  expect(band.borders.bottom).toMatchObject({ color: '#FFFFFF', widthEmu: 12700 });
  const roundTrip = readZip(await savePresentation(pres));
  expect(roundTrip.entries.find((entry) => entry.name === 'ppt/tableStyles.xml')?.data).toEqual(
    new Uint8Array(style),
  );
});

it('resolves the built-in Medium Style 2 when tableStyles.xml only has its default GUID', () => {
  const pres = createPresentation();
  const table = addSlideTable(addBlankSlide(pres), {
    x: inches(1),
    y: inches(1),
    w: inches(6),
    h: inches(3),
    rows: [
      ['A', 'B'],
      ['C', 'D'],
      ['E', 'F'],
    ],
  });
  setTableStyleFlags(table, { firstRow: true, bandRow: true });
  const cells = getTableCells(getSlideShapes(getSlides(pres)[0]!)[0]!);
  const header = getTableCellAppearanceEffective(pres, cells[0]![0]!);
  const band = getTableCellAppearanceEffective(pres, cells[1]![0]!);
  expect(header.fill).toEqual({ kind: 'solid', color: '#4F81BD' });
  expect(band.fill).toEqual({ kind: 'solid', color: '#D0D8E8' });
  expect(header.borders.bottom).toMatchObject({ color: '#FFFFFF', widthEmu: 38100 });
});

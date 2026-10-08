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
  mergeTableCells,
  loadPresentation,
  savePresentation,
  setTableStyleFlags,
  setTableStyleId,
  getTableStyleId,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';

it('resolves the Medium Style 2 definition saved by the reference desktop app on Mac', async () => {
  // Captured from the tableStyles.xml the reference desktop app wrote on Mac; validated against
  // dml-main.xsd.
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

it('retains the built-in No Style, Table Grid appearance across flags and saving', async () => {
  const pres = createPresentation();
  const table = addSlideTable(addBlankSlide(pres), {
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
  const styleId = '{5940675A-B579-460E-94D1-54222C63F5DA}';
  setTableStyleId(table, styleId);
  const cells = getTableCells(table);
  for (let flags = 0; flags < 64; flags++) {
    setTableStyleFlags(table, {
      firstRow: Boolean(flags & 1),
      lastRow: Boolean(flags & 2),
      firstCol: Boolean(flags & 4),
      lastCol: Boolean(flags & 8),
      bandRow: Boolean(flags & 16),
      bandCol: Boolean(flags & 32),
    });
    for (const cell of cells.flat()) {
      const appearance = getTableCellAppearanceEffective(pres, cell);
      expect(appearance.fill).toEqual({ kind: 'none' });
      for (const side of ['left', 'right', 'top', 'bottom'] as const)
        expect(appearance.borders[side]).toMatchObject({ color: '#000000', widthEmu: 12700 });
      expect(appearance.borders.tlToBr).toBeNull();
      expect(appearance.borders.blToTr).toBeNull();
    }
  }
  const loaded = await loadPresentation(await savePresentation(pres));
  const loadedTable = getSlideShapes(getSlides(loaded)[0]!)[0]!;
  expect(getTableStyleId(loadedTable)).toBe(styleId);
  expect(getTableCellAppearanceEffective(loaded, getTableCells(loadedTable)[1]![1]!)).toEqual(
    getTableCellAppearanceEffective(pres, cells[1]![1]!),
  );
});

it('applies Total Row to a vertical merge ending at the last row, as in the reference desktop app on Mac', async () => {
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
  mergeTableCells(table, { row: 0, col: 0, rowSpan: 3, colSpan: 1 });
  setTableStyleFlags(table, { firstRow: false, bandRow: true, lastRow: true });
  const cell = getTableCells(table)[0]![0]!;
  expect(getTableCellAppearanceEffective(pres, cell).fill).toEqual({
    kind: 'solid',
    color: '#4F81BD',
  });
  setTableStyleFlags(table, { lastRow: false });
  expect(getTableCellAppearanceEffective(pres, cell).fill).toEqual({
    kind: 'solid',
    color: '#D0D8E8',
  });
  setTableStyleFlags(table, { lastRow: true });
  const loaded = await loadPresentation(await savePresentation(pres));
  const loadedCell = getTableCells(getSlideShapes(getSlides(loaded)[0]!)[0]!)[0]![0]!;
  expect(getTableCellAppearanceEffective(loaded, loadedCell).fill).toEqual({
    kind: 'solid',
    color: '#4F81BD',
  });
});

it('applies Last Column to a horizontal merge reaching the right edge, as in the reference desktop app on Mac', async () => {
  const pres = createPresentation();
  const table = addSlideTable(addBlankSlide(pres), {
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
  mergeTableCells(table, { row: 0, col: 0, rowSpan: 1, colSpan: 3 });
  mergeTableCells(table, { row: 2, col: 0, rowSpan: 1, colSpan: 2 });
  setTableStyleFlags(table, { firstRow: false, firstCol: false, bandRow: true, lastCol: true });
  const cells = getTableCells(table);
  expect(getTableCellAppearanceEffective(pres, cells[0]![0]!).fill).toEqual({
    kind: 'solid',
    color: '#4F81BD',
  });
  expect(getTableCellAppearanceEffective(pres, cells[2]![0]!).fill).toEqual({
    kind: 'solid',
    color: '#D0D8E8',
  });
  setTableStyleFlags(table, { lastCol: false });
  expect(getTableCellAppearanceEffective(pres, cells[0]![0]!).fill).toEqual({
    kind: 'solid',
    color: '#D0D8E8',
  });
  setTableStyleFlags(table, { lastCol: true });
  const loaded = await loadPresentation(await savePresentation(pres));
  const loadedCell = getTableCells(getSlideShapes(getSlides(loaded)[0]!)[0]!)[0]![0]!;
  expect(getTableCellAppearanceEffective(loaded, loadedCell).fill).toEqual({
    kind: 'solid',
    color: '#4F81BD',
  });
});

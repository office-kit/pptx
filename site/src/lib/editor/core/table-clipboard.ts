import {
  getTableCells,
  getTableCellText,
  getTableCellPosition,
  getTableCellSpan,
  insertTableColumn,
  insertTableRow,
  setTableCellText,
  type SlideShapeData,
} from '@office-kit/pptx';
import { tableCellsInRange, tableSelectionBlock } from './table-selection.ts';
import type { CellSelection } from './selection.ts';

/** Spreadsheet text uses quoted fields for embedded tabs, newlines and quotes. */
export function parseTableClipboard(text: string): string[][] | null {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let quoted = false;
  let closed = false;
  for (let i = 0; i < text.length; i++) {
    const char = text[i]!;
    if (quoted) {
      if (char === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          quoted = false;
          closed = true;
        }
      } else field += char;
      continue;
    }
    if (char === '\t' || char === '\r' || char === '\n') {
      row.push(field);
      field = '';
      closed = false;
      if (char !== '\t') {
        rows.push(row);
        row = [];
        if (char === '\r' && text[i + 1] === '\n') i++;
      }
    } else if (closed) return null;
    else if (char === '"' && field === '') quoted = true;
    else field += char;
  }
  if (quoted) return null;
  if (row.length || field.length || closed || !rows.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

export function tableHasMergedCells(table: SlideShapeData): boolean {
  return getTableCells(table).some((row) =>
    row.some((cell) => {
      const span = getTableCellSpan(cell);
      return span.hMerge || span.vMerge || span.gridSpan > 1 || span.rowSpan > 1;
    }),
  );
}

/** Check before starting a transaction so rejected pastes leave text and history intact. */
export function canPasteTableCells(
  table: SlideShapeData,
  row: number,
  col: number,
  values: string[][],
): boolean {
  const cells = getTableCells(table);
  const width = values.reduce((max, value) => Math.max(max, value.length), 0);
  if (
    (row + values.length > cells.length || col + width > (cells[0]?.length ?? 0)) &&
    tableHasMergedCells(table)
  )
    return false;
  return values.every((valuesRow, r) =>
    valuesRow.every((_, c) => {
      const cell = cells[row + r]?.[col + c];
      if (!cell) return true;
      const span = getTableCellSpan(cell);
      return !span.hMerge && !span.vMerge && span.gridSpan === 1 && span.rowSpan === 1;
    }),
  );
}

/** Apply a preflighted paste using the existing table mutation API. */
export function pasteTableCells(
  table: SlideShapeData,
  row: number,
  col: number,
  values: string[][],
): void {
  const cells = getTableCells(table);
  const width = values.reduce((max, value) => Math.max(max, value.length), 0);
  for (let c = cells[0]?.length ?? 0; c < col + width; c++) insertTableColumn(table);
  for (let r = cells.length; r < row + values.length; r++) insertTableRow(table);
  const targets = getTableCells(table);
  values.forEach((valuesRow, r) =>
    valuesRow.forEach((value, c) => {
      setTableCellText(targets[row + r]![col + c]!, value, { preserveFormatting: true });
    }),
  );
}

export function copyTableCellValues(table: SlideShapeData, selection: CellSelection): string[][] {
  const cells = getTableCells(table);
  const selected = tableCellsInRange(cells, tableSelectionBlock(selection));
  if (!selected.size) return [];
  let top = cells.length,
    left = cells[0]!.length,
    bottom = 0,
    right = 0;
  for (const cell of selected) {
    const { row, col } = getTableCellPosition(cell);
    top = Math.min(top, row);
    left = Math.min(left, col);
    bottom = Math.max(bottom, row);
    right = Math.max(right, col);
  }
  return cells.slice(top, bottom + 1).map((row) =>
    row.slice(left, right + 1).map((cell) => {
      const span = getTableCellSpan(cell);
      return span.hMerge || span.vMerge ? '' : getTableCellText(cell);
    }),
  );
}

export function serializeTableClipboard(values: readonly (readonly string[])[]): string {
  return values
    .map((row) =>
      row
        .map((value) => (/[\t\r\n"]/.test(value) ? `"${value.replaceAll('"', '""')}"` : value))
        .join('\t'),
    )
    .join('\n');
}

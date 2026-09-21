import {
  getTableCells,
  getTableCellSpan,
  type TableCellData,
  type SlideShapeData,
} from '@office-kit/pptx';

/** Move past a merged cell and resolve covered destinations to their visible anchor. */
export function neighboringTableCell(
  table: SlideShapeData,
  row: number,
  col: number,
  dr: number,
  dc: number,
): { row: number; col: number } | null {
  const cells = getTableCells(table);
  const current = cells[row]?.[col];
  if (!current) return null;
  const span = getTableCellSpan(current);
  const targetRow = row + (dr > 0 ? span.rowSpan : dr);
  const targetCol = col + (dc > 0 ? span.gridSpan : dc);
  if (!cells[targetRow]?.[targetCol]) return null;
  for (let r = 0; r <= targetRow; r++) {
    for (let c = 0; c <= targetCol; c++) {
      const cell = cells[r]?.[c];
      if (!cell) continue;
      const candidate = getTableCellSpan(cell);
      if (
        !candidate.hMerge &&
        !candidate.vMerge &&
        r + candidate.rowSpan > targetRow &&
        c + candidate.gridSpan > targetCol
      )
        return { row: r, col: c };
    }
  }
  return null;
}

/** Include every physical cell of merged anchors intersecting the selected rectangle. */
export function tableCellsInRange(
  cells: readonly (readonly TableCellData[])[],
  block: { row: number; col: number; rowSpan: number; colSpan: number },
): Set<TableCellData> {
  const result = new Set<TableCellData>();
  for (let r = 0; r < cells.length; r++) {
    for (let c = 0; c < cells[r]!.length; c++) {
      const cell = cells[r]![c]!;
      const span = getTableCellSpan(cell);
      if (
        span.hMerge ||
        span.vMerge ||
        r >= block.row + block.rowSpan ||
        r + span.rowSpan <= block.row ||
        c >= block.col + block.colSpan ||
        c + span.gridSpan <= block.col
      )
        continue;
      for (let y = r; y < Math.min(r + span.rowSpan, cells.length); y++)
        for (let x = c; x < Math.min(c + span.gridSpan, cells[y]!.length); x++)
          result.add(cells[y]![x]!);
    }
  }
  return result;
}

export function tableSelectionBlock(selection: {
  row: number;
  col: number;
  end?: { row: number; col: number };
}) {
  const end = selection.end ?? selection;
  return {
    row: Math.min(selection.row, end.row),
    col: Math.min(selection.col, end.col),
    rowSpan: Math.abs(selection.row - end.row) + 1,
    colSpan: Math.abs(selection.col - end.col) + 1,
  };
}

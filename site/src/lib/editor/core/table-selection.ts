import { getTableCells, getTableCellSpan, type SlideShapeData } from '@office-kit/pptx';

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

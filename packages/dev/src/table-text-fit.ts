import * as pptx from '@office-kit/pptx';
import { measureTableCellHeight } from '@office-kit/pptx-preview';

/** Grow rows to fit horizontal cell text while retaining the rotated top edge. */
export function fitTableText(
  presentation: pptx.PresentationData,
  table: pptx.SlideShapeData,
  equalRows?: { start: number; count: number },
  anchor: 'top' | 'bottom' = 'top',
): void {
  const box = pptx.getShapeBoundsResolved(presentation, table);
  const original = pptx.getTableRowHeights(table);
  const sum = original.reduce((a, b) => a + b, 0);
  if (!box || sum <= 0) return;
  const heights = original.map((value) => (value * box.h) / sum);
  const constraints: { row: number; count: number; height: number }[] = [];
  pptx.getTableCells(table).forEach((cells, row) =>
    cells.forEach((cell, column) => {
      const span = pptx.getTableCellSpan(cell);
      const height = measureTableCellHeight(presentation, table, row, column);
      if (height !== null) constraints.push({ row, count: span.rowSpan, height });
    }),
  );
  // Satisfy single rows before merged regions; a merged cell constrains its
  // combined row height, not each physical grid row separately.
  constraints.sort((a, b) => a.count - b.count);
  let changed = false;
  for (const { row, count, height } of constraints) {
    const current = heights.slice(row, row + count).reduce((a, b) => a + b, 0);
    if (height <= current || row + count > heights.length) continue;
    changed = true;
    const overlap = equalRows
      ? Math.max(
          0,
          Math.min(row + count, equalRows.start + equalRows.count) - Math.max(row, equalRows.start),
        )
      : 0;
    if (equalRows && overlap) {
      // A distributed group must remain equal even when only one of its cells
      // needs more room. Each intersecting row contributes to this constraint.
      const extra = (height - current) / overlap;
      for (let r = equalRows.start; r < equalRows.start + equalRows.count; r++)
        heights[r] = heights[r]! + extra;
    } else {
      const extra = (height - current) / count;
      for (let r = row; r < row + count; r++) heights[r] = heights[r]! + extra;
    }
  }
  if (!changed) return;
  const next = heights.map(Math.ceil);
  const extent = next.reduce((a, b) => a + b, 0);
  const growth = extent - box.h;
  const angle = (pptx.getShapeRotation(table) * Math.PI) / 180;
  const direction = anchor === 'top' ? 1 : -1;
  next.forEach((height, row) => pptx.setTableRowHeight(table, row, pptx.emu(height)));
  pptx.setShapeBounds(table, {
    x: pptx.emu(Math.round(box.x - (Math.sin(angle) * direction * growth) / 2)),
    y: pptx.emu(Math.round(box.y + ((Math.cos(angle) * direction - 1) * growth) / 2)),
    w: box.w,
    h: pptx.emu(extent),
  });
}

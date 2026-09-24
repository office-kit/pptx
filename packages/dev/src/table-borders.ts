import * as pptx from '@office-kit/pptx';

export type TableBorderMode =
  | 'all'
  | 'none'
  | 'outside'
  | 'inside'
  | 'horizontal'
  | 'vertical'
  | 'left'
  | 'right'
  | 'top'
  | 'bottom'
  | 'tlToBr'
  | 'blToTr';
export function applyTableBorders(
  presentation: pptx.PresentationData,
  table: pptx.SlideShapeData,
  mode: TableBorderMode,
  pen: Partial<pptx.TableCellBorder>,
  target?: { row: number; column: number; rows?: number; columns?: number },
) {
  const cells = pptx.getTableCells(table);
  const { rows, cols } = pptx.getTableDimensions(table);
  const owners = Array.from({ length: rows }, () => Array<number>(cols).fill(-1));
  const regions: Array<{ r: number; c: number; bottom: number; right: number }> = [];
  for (let r = 0; r < rows; r++)
    for (let c = 0; c < cols; c++) {
      const span = pptx.getTableCellSpan(cells[r]![c]!);
      if (span.hMerge || span.vMerge) continue;
      const region = {
        r,
        c,
        bottom: Math.min(rows, r + span.rowSpan),
        right: Math.min(cols, c + span.gridSpan),
      };
      const index = regions.push(region) - 1;
      for (let rr = r; rr < region.bottom; rr++)
        for (let cc = c; cc < region.right; cc++) owners[rr]![cc] = index;
    }
  const selected = target
    ? regions.find((region) => region.r === target.row && region.c === target.column)
    : undefined;
  if (target && !selected) throw new Error('Select the merged cell anchor.');
  const area =
    target?.rows !== undefined || target?.columns !== undefined
      ? {
          r: target.row,
          c: target.column,
          bottom: target.row + target.rows!,
          right: target.column + target.columns!,
        }
      : (selected ?? { r: 0, c: 0, bottom: rows, right: cols });
  if (
    ![area.r, area.c, area.bottom, area.right].every(Number.isInteger) ||
    area.r < 0 ||
    area.c < 0 ||
    area.bottom <= area.r ||
    area.right <= area.c ||
    area.bottom > rows ||
    area.right > cols
  )
    throw new Error('Invalid table cell range.');
  for (const region of regions) {
    const intersects =
      region.r < area.bottom &&
      region.bottom > area.r &&
      region.c < area.right &&
      region.right > area.c;
    if (
      intersects &&
      (region.r < area.r ||
        region.c < area.c ||
        region.bottom > area.bottom ||
        region.right > area.right)
    )
      throw new Error('Select the entire merged cell.');
  }
  type Side = keyof pptx.TableCellBorders;
  const edits = new Map<pptx.TableCellData, Partial<Record<Side, Partial<pptx.TableCellBorder>>>>();
  const write = (r: number, c: number, side: Side, border: Partial<pptx.TableCellBorder>) => {
    const cell = cells[r]?.[c];
    if (!cell) return;
    edits.set(cell, { ...edits.get(cell), [side]: border });
  };
  // Materialize merged perimeters before editing segments. Empty line properties
  // retain inherited styling while preventing a changed anchor from leaking to
  // untouched segments of that same merged edge.
  for (const region of regions) {
    if (region.bottom - region.r === 1 && region.right - region.c === 1) continue;
    const anchor = pptx.getTableCellBorders(presentation, cells[region.r]![region.c]!);
    const preserve = (r: number, c: number, side: Side) =>
      write(
        r,
        c,
        side,
        pptx.getTableCellBorders(presentation, cells[r]![c]!)[side] ?? anchor[side] ?? {},
      );
    for (let r = region.r; r < region.bottom; r++) {
      preserve(r, region.c, 'left');
      preserve(r, region.right - 1, 'right');
    }
    for (let c = region.c; c < region.right; c++) {
      preserve(region.r, c, 'top');
      preserve(region.bottom - 1, c, 'bottom');
    }
  }
  const border = mode === 'none' ? { noFill: true } : pen;
  const vertical = (r: number, c: number) => {
    if (c > 0 && c < cols && owners[r]![c - 1] === owners[r]![c]) return;
    write(r, c - 1, 'right', border);
    write(r, c, 'left', border);
  };
  const horizontal = (r: number, c: number) => {
    if (r > 0 && r < rows && owners[r - 1]![c] === owners[r]![c]) return;
    write(r - 1, c, 'bottom', border);
    write(r, c, 'top', border);
  };
  for (let r = area.r; r < area.bottom; r++)
    for (let c = area.c; c <= area.right; c++) {
      const outside = c === area.c || c === area.right;
      if (
        mode === 'all' ||
        mode === 'none' ||
        (mode === 'outside' && outside) ||
        ((mode === 'inside' || mode === 'vertical') && !outside) ||
        (mode === 'left' && c === area.c) ||
        (mode === 'right' && c === area.right)
      )
        vertical(r, c);
    }
  for (let r = area.r; r <= area.bottom; r++)
    for (let c = area.c; c < area.right; c++) {
      const outside = r === area.r || r === area.bottom;
      if (
        mode === 'all' ||
        mode === 'none' ||
        (mode === 'outside' && outside) ||
        ((mode === 'inside' || mode === 'horizontal') && !outside) ||
        (mode === 'top' && r === area.r) ||
        (mode === 'bottom' && r === area.bottom)
      )
        horizontal(r, c);
    }
  for (const region of regions)
    if (
      region.r >= area.r &&
      region.c >= area.c &&
      region.bottom <= area.bottom &&
      region.right <= area.right
    ) {
      if (mode === 'none') {
        write(region.r, region.c, 'tlToBr', border);
        write(region.r, region.c, 'blToTr', border);
      } else if (mode === 'tlToBr' || mode === 'blToTr') write(region.r, region.c, mode, pen);
    }
  for (const [cell, sides] of edits) pptx.setTableCellBorders(cell, sides);
}

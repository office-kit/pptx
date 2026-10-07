// The cells the Table Design and Table Layout tabs act on: the selected cell
// range, or every cell when the table itself is selected (as in PowerPoint).
import {
  getShapeBoundsResolved,
  getTableCells,
  getTableCellSpan,
  getTableColumnWidths,
  getTableRowHeights,
  getTableSize,
  isTableShape,
  setShapeBounds,
  type PresentationData,
  type SlideShapeData,
  type TableCellData,
} from '@office-kit/pptx';
import type { EditorController } from './controller.svelte.ts';
import { tableCellsInRange, tableSelectionBlock } from './table-selection.ts';

export interface TableTarget {
  readonly table: SlideShapeData;
  readonly tableId: number;
  readonly slideIndex: number;
  readonly cells: readonly (readonly TableCellData[])[];
  readonly block: { row: number; col: number; rowSpan: number; colSpan: number };
  readonly selected: ReadonlySet<TableCellData>;
  /** The top-left selected cell, whose state the toggles show. */
  readonly anchor: TableCellData;
  readonly widths: readonly number[];
  readonly heights: readonly number[];
  /** Some cell spans rows or columns; row/column insertion needs them split first. */
  readonly merged: boolean;
}

export function tableTarget(editor: EditorController): TableTarget | null {
  const doc = editor.doc;
  const sel = doc.selection;
  let tableId: number;
  if (sel.kind === 'cell') tableId = sel.shapeId;
  else if (sel.kind === 'shape' && sel.shapeIds.length === 1) tableId = sel.shapeIds[0]!;
  else return null;
  const table = doc.shapeById(sel.slideIndex, tableId);
  if (!table || !isTableShape(table)) return null;
  const cells = getTableCells(table);
  const widths = getTableColumnWidths(table);
  const heights = getTableRowHeights(table);
  const block =
    sel.kind === 'cell'
      ? tableSelectionBlock(sel)
      : { row: 0, col: 0, rowSpan: cells.length, colSpan: widths.length };
  const anchor = cells[block.row]?.[block.col];
  if (!anchor) return null;
  const merged = cells.some((row) =>
    row.some((cell) => {
      const span = getTableCellSpan(cell);
      return span.gridSpan > 1 || span.rowSpan > 1 || span.hMerge || span.vMerge;
    }),
  );
  return {
    table,
    tableId,
    slideIndex: sel.slideIndex,
    cells,
    block,
    selected: tableCellsInRange(cells, block),
    anchor,
    widths,
    heights,
    merged,
  };
}

/**
 * Grows or shrinks the graphic frame to the sum of its rows and columns, as
 * PowerPoint keeps them equal; the library leaves `<a:xfrm>` to the caller.
 */
export function fitTableFrame(pres: PresentationData, table: SlideShapeData): void {
  const bounds = getShapeBoundsResolved(pres, table);
  if (!bounds) return;
  const size = getTableSize(table);
  setShapeBounds(table, { ...bounds, w: size.width, h: size.height });
}

// Row, column and merge commands on the selected table cells, as the cell
// context menu offers them. Each command is one undoable edit.
import {
  getTableCells,
  getTableCellSpan,
  insertTableColumn,
  insertTableRow,
  mergeTableCells,
  removeShape,
  removeTableColumn,
  removeTableRow,
  splitTableCell,
  type SlideShapeData,
  type TableCellData,
} from '@office-kit/pptx';
import type { EditorController } from './controller.svelte.ts';
import type { CellSelection } from './selection.ts';
import { tableHasMergedCells } from './table-clipboard.ts';
import { tableSelectionBlock } from './table-selection.ts';
import { t } from '../i18n/i18n.svelte.ts';

export type TableAxis = 'row' | 'column';

interface CellTarget {
  readonly selection: CellSelection;
  readonly table: SlideShapeData;
  readonly cells: readonly (readonly TableCellData[])[];
  readonly block: ReturnType<typeof tableSelectionBlock>;
}

export function selectedCellTarget(editor: EditorController): CellTarget | null {
  const selection = editor.doc.selection;
  if (selection.kind !== 'cell') return null;
  const table = editor.doc.shapeById(selection.slideIndex, selection.shapeId);
  if (!table) return null;
  return {
    selection,
    table,
    cells: getTableCells(table),
    block: tableSelectionBlock(selection),
  };
}

/** A rectangle of two or more plain (unmerged) cells. */
export function canMergeTableBlock(
  cells: readonly (readonly TableCellData[])[],
  block: ReturnType<typeof tableSelectionBlock>,
): boolean {
  const columns = cells[0]?.length ?? 0;
  if (
    block.rowSpan * block.colSpan < 2 ||
    block.row + block.rowSpan > cells.length ||
    block.col + block.colSpan > columns
  )
    return false;
  return cells.slice(block.row, block.row + block.rowSpan).every((row) =>
    row.slice(block.col, block.col + block.colSpan).every((cell) => {
      const span = getTableCellSpan(cell);
      return !span.hMerge && !span.vMerge && span.rowSpan === 1 && span.gridSpan === 1;
    }),
  );
}

export function mergeSelectedCells(editor: EditorController): void {
  const target = selectedCellTarget(editor);
  if (!target || !canMergeTableBlock(target.cells, target.block)) return;
  const { selection, table, block } = target;
  editor.doc.transact(t('Merge cells'), () => {
    mergeTableCells(table, block, { coveredText: 'append' });
    editor.doc.selectCell(selection.slideIndex, selection.shapeId, block.row, block.col);
  });
}

/** The selected cell when it spans several grid cells (the editor splits merged cells only). */
export function mergedSelectedCell(editor: EditorController): TableCellData | null {
  const target = selectedCellTarget(editor);
  const cell = target?.cells[target.selection.row]?.[target.selection.col];
  if (!cell) return null;
  const span = getTableCellSpan(cell);
  return span.gridSpan > 1 || span.rowSpan > 1 ? cell : null;
}

export function splitSelectedCell(editor: EditorController): void {
  const cell = mergedSelectedCell(editor);
  if (cell) editor.doc.transact(t('Split cell'), () => splitTableCell(cell));
}

/** Rows and columns can only be added or removed while no cell is merged. */
export function canEditTableLines(editor: EditorController): boolean {
  const target = selectedCellTarget(editor);
  return target !== null && !tableHasMergedCells(target.table);
}

/** Inserts as many rows or columns as are selected, before or after the selection. */
export function insertTableLines(editor: EditorController, axis: TableAxis, after: boolean): void {
  const target = selectedCellTarget(editor);
  if (!target || tableHasMergedCells(target.table)) return;
  const { selection, table, block } = target;
  const count = axis === 'row' ? block.rowSpan : block.colSpan;
  const start = axis === 'row' ? block.row : block.col;
  const at = after ? start + count : start;
  editor.doc.transact(t(axis === 'row' ? 'Insert Rows' : 'Insert Columns'), () => {
    for (let i = 0; i < count; i++) {
      if (axis === 'row') insertTableRow(table, at);
      else insertTableColumn(table, at);
    }
    // The selected cells keep their content; inserting before them shifts them.
    const shift = after ? 0 : count;
    editor.doc.selectCell(
      selection.slideIndex,
      selection.shapeId,
      block.row + (axis === 'row' ? shift : 0),
      block.col + (axis === 'column' ? shift : 0),
    );
  });
}

/** Deletes the selected rows or columns; deleting all of them deletes the table, as PowerPoint does. */
export function deleteTableLines(editor: EditorController, axis: TableAxis): void {
  const target = selectedCellTarget(editor);
  if (!target || tableHasMergedCells(target.table)) return;
  const { selection, table, block, cells } = target;
  const total = axis === 'row' ? cells.length : (cells[0]?.length ?? 0);
  const count = axis === 'row' ? block.rowSpan : block.colSpan;
  const start = axis === 'row' ? block.row : block.col;
  if (count >= total) {
    deleteTable(editor);
    return;
  }
  editor.doc.transact(t(axis === 'row' ? 'Delete Rows' : 'Delete Columns'), () => {
    for (let index = start + count - 1; index >= start; index--) {
      if (axis === 'row') removeTableRow(table, index);
      else removeTableColumn(table, index);
    }
    const last = total - count - 1;
    editor.doc.selectCell(
      selection.slideIndex,
      selection.shapeId,
      axis === 'row' ? Math.min(start, last) : block.row,
      axis === 'column' ? Math.min(start, last) : block.col,
    );
  });
}

export function deleteTable(editor: EditorController): void {
  const target = selectedCellTarget(editor);
  if (!target) return;
  editor.doc.transact(t('Delete Table'), () => {
    removeShape(target.table);
    editor.doc.clearShapeSelection();
  });
}

/** Extends the cell selection across the selected rows or columns. */
export function selectTableLines(editor: EditorController, axis: TableAxis): void {
  const target = selectedCellTarget(editor);
  if (!target) return;
  const { selection, cells, block } = target;
  const lastRow = cells.length - 1;
  const lastCol = (cells[0]?.length ?? 1) - 1;
  const doc = editor.doc;
  if (axis === 'row') {
    doc.selectCell(selection.slideIndex, selection.shapeId, block.row, 0);
    doc.selectCell(
      selection.slideIndex,
      selection.shapeId,
      block.row + block.rowSpan - 1,
      lastCol,
      true,
    );
  } else {
    doc.selectCell(selection.slideIndex, selection.shapeId, 0, block.col);
    doc.selectCell(
      selection.slideIndex,
      selection.shapeId,
      lastRow,
      block.col + block.colSpan - 1,
      true,
    );
  }
}

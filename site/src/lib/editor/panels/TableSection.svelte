<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import { selectedShapeId } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import { mergeTableCells, splitTableCell, getTableCells, isTableShape, getTableCellText, getTableCellSpan, getTableCellFill, getTableColumnWidths, getTableRowHeights, getTableCellAlignment, getTableCellAnchor, setTableCellText, setTableCellFill, setTableCellAlignment, setTableCellAnchor, setTableRowHeight, setTableColumnWidth, insertTableRow, insertTableColumn, removeTableRow, removeTableColumn, inches } from '@office-kit/pptx';

  const editor = getEditor();
  const doc = editor.doc;
  const tableState = $derived.by(() => {
    doc.version;
    const sel = doc.selection;
    const id = selectedShapeId(sel);
    if (id === null || (sel.kind === 'shape' && sel.shapeIds.length !== 1)) return null;
    const table = doc.shapeById(sel.slideIndex, id);
    if (!table || !isTableShape(table)) return null;
    const cells = getTableCells(table);
    const row = sel.kind === 'cell' ? sel.row : 0;
    const col = sel.kind === 'cell' ? sel.col : 0;
    const cell = cells[row]?.[col];
    return { table, id, cells, row, col, cell, widths: getTableColumnWidths(table), heights: getTableRowHeights(table), merged: cells.some((r) => r.some((c) => { const s = getTableCellSpan(c); return s.gridSpan > 1 || s.rowSpan > 1 || s.hMerge || s.vMerge; })) };
  });
  let rangeEnd = $state<{ id: number; slide: number; anchorRow: number; anchorCol: number; row: number; col: number } | null>(null);
  const block = $derived.by(() => {
    if (!tableState) return null;
    const end = rangeEnd?.id === tableState.id && rangeEnd.slide === doc.selection.slideIndex && rangeEnd.anchorRow === tableState.row && rangeEnd.anchorCol === tableState.col ? rangeEnd : tableState;
    const row = Math.min(tableState.row, end.row);
    const col = Math.min(tableState.col, end.col);
    return { row, col, rowSpan: Math.abs(tableState.row - end.row) + 1, colSpan: Math.abs(tableState.col - end.col) + 1 };
  });
  const canMerge = $derived.by(() => {
    if (!tableState || !block || block.rowSpan * block.colSpan < 2 || block.row + block.rowSpan > tableState.cells.length || block.col + block.colSpan > tableState.widths.length) return false;
    return tableState.cells.slice(block.row, block.row + block.rowSpan).every(row => row.slice(block.col, block.col + block.colSpan).every(cell => {
      const span = getTableCellSpan(cell);
      return !span.hMerge && !span.vMerge && span.rowSpan === 1 && span.gridSpan === 1;
    }));
  });
  function select(row: number, col: number, extend: boolean) {
    if (!tableState) return;
    if (extend) rangeEnd = { id: tableState.id, slide: doc.selection.slideIndex, anchorRow: tableState.row, anchorCol: tableState.col, row, col };
    else { rangeEnd = null; doc.selectCell(doc.selection.slideIndex, tableState.id, row, col); }
  }
  function merge() {
    const s = tableState;
    const b = block;
    if (!s || !b || !canMerge) return;
    doc.transact(t('Merge cells'), () => {
      mergeTableCells(s.table, b, { coveredText: 'append' });
      doc.selectCell(doc.selection.slideIndex, s.id, b.row, b.col);
    });
    rangeEnd = null;
  }
  function split() {
    const cell = tableState?.cell;
    if (cell) doc.transact(t('Split cell'), () => splitTableCell(cell));
    rangeEnd = null;
  }
  function structure(axis: 'row' | 'column', remove: boolean) {
    rangeEnd = null;
    if (!tableState) return;
    const s = tableState;
    doc.transact(t(remove ? 'Delete table row or column' : 'Insert table row or column'), () => {
      if (axis === 'row') {
        if (remove) removeTableRow(s.table, s.row);
        else insertTableRow(s.table, s.row + 1);
      } else {
        if (remove) removeTableColumn(s.table, s.col);
        else insertTableColumn(s.table, s.col + 1);
      }
      doc.selectCell(doc.selection.slideIndex, s.id, axis === 'row' ? Math.min(s.row, s.cells.length - (remove ? 2 : 1)) : s.row, axis === 'column' ? Math.min(s.col, s.widths.length - (remove ? 2 : 1)) : s.col);
    });
  }
</script>

{#if tableState}
  <section class="table-controls" aria-label={t('Table options')}>
    <strong>{t('Table options')}</strong>
    <div class="cell-grid ok-scroll">
      <table aria-label={t('Table cells')}><tbody>
        {#each tableState.cells as row, r}<tr>
          {#each row as cell, c}
            {@const span = getTableCellSpan(cell)}
            {#if !span.hMerge && !span.vMerge}
              <td rowspan={span.rowSpan} colspan={span.gridSpan}><button class="ok-btn" aria-label={`${t('Cell')} ${r + 1}, ${c + 1}`} aria-pressed={!!block && r >= block.row && r < block.row + block.rowSpan && c >= block.col && c < block.col + block.colSpan} onclick={(e) => select(r, c, e.shiftKey)}>{getTableCellText(cell) || '—'}</button></td>
            {/if}
          {/each}
        </tr>{/each}
      </tbody></table>
    </div>
    <small>{t('Shift-click another cell to select a range')}</small>
    {#if tableState.cell}
      <div class="actions">
        <button class="ok-btn" disabled={!canMerge} onclick={merge}>{t('Merge cells')}</button>
        <button class="ok-btn" disabled={getTableCellSpan(tableState.cell).gridSpan === 1 && getTableCellSpan(tableState.cell).rowSpan === 1} onclick={split}>{t('Split cell')}</button>
      </div>
      <label>{t('Cell text')}<textarea class="ok-input" rows="3" value={getTableCellText(tableState.cell)} onchange={(e) => { const s = tableState; if (s?.cell) doc.transact(t('Edit cell text'), () => setTableCellText(s.cell!, e.currentTarget.value, { preserveFormatting: true })); }}></textarea></label>
      <label>{t('Cell fill')}<input type="color" value={getTableCellFill(tableState.cell) ?? '#ffffff'} onchange={(e) => { const cell = tableState?.cell; if (cell) doc.transact(t('Cell fill'), () => setTableCellFill(cell, e.currentTarget.value)); }} /></label>
      <label>{t('Horizontal alignment')}<select aria-label={t('Horizontal alignment')} class="ok-input" value={getTableCellAlignment(tableState.cell) ?? 'l'} onchange={(e) => { const cell = tableState?.cell; const v = e.currentTarget.value; if (cell && (v === 'l' || v === 'ctr' || v === 'r')) doc.transact(t('Horizontal alignment'), () => setTableCellAlignment(cell, v)); }}>
        <option value="l">{t('Left')}</option><option value="ctr">{t('Center')}</option><option value="r">{t('Right')}</option>
      </select></label>
      <label>{t('Vertical alignment')}<select aria-label={t('Vertical alignment')} class="ok-input" value={getTableCellAnchor(tableState.cell) ?? 'top'} onchange={(e) => { const cell = tableState?.cell; const v = e.currentTarget.value; if (cell && (v === 'top' || v === 'center' || v === 'bottom')) doc.transact(t('Vertical alignment'), () => setTableCellAnchor(cell, v)); }}>
        <option value="top">{t('Top')}</option><option value="center">{t('Center')}</option><option value="bottom">{t('Bottom')}</option>
      </select></label>
      <label>{t('Row height (inches)')}<input class="ok-input" type="number" min="0.01" step="0.01" required value={tableState.heights[tableState.row]! / inches(1)} onchange={(e) => { const s = tableState; if (s && e.currentTarget.reportValidity()) doc.transact(t('Resize table row'), () => setTableRowHeight(s.table, s.row, inches(e.currentTarget.valueAsNumber))); }} /></label>
      <label>{t('Column width (inches)')}<input class="ok-input" type="number" min="0.01" step="0.01" required value={tableState.widths[tableState.col]! / inches(1)} onchange={(e) => { const s = tableState; if (s && e.currentTarget.reportValidity()) doc.transact(t('Resize table column'), () => setTableColumnWidth(s.table, s.col, inches(e.currentTarget.valueAsNumber))); }} /></label>
      <div class="actions">
        <button class="ok-btn" disabled={tableState.merged} onclick={() => structure('row', false)}>{t('Insert row below')}</button>
        <button class="ok-btn" disabled={tableState.merged} onclick={() => structure('column', false)}>{t('Insert column right')}</button>
        <button class="ok-btn" disabled={tableState.merged || tableState.cells.length <= 1} onclick={() => structure('row', true)}>{t('Delete row')}</button>
        <button class="ok-btn" disabled={tableState.merged || tableState.widths.length <= 1} onclick={() => structure('column', true)}>{t('Delete column')}</button>
      </div>
      {#if tableState.merged}<small>{t('Split merged cells before changing rows or columns')}</small>{/if}
    {/if}
  </section>
{/if}

<style>
  .table-controls { padding: 12px; display: flex; flex-direction: column; gap: 10px; border-bottom: 1px solid var(--ok-border); }
  strong { font-size: 12px; }
  label { display: grid; gap: 4px; font-size: 11px; }
  textarea, input, select { width: 100%; box-sizing: border-box; }
  .cell-grid { overflow: auto; max-height: 220px; }
  table { border-collapse: collapse; width: 100%; }
  td { border: 1px solid var(--ok-border); padding: 1px; }
  td button { width: 100%; min-width: 50px; max-width: 130px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  button[aria-pressed='true'] { background: var(--ok-accent); color: white; }
  .actions { display: grid; grid-template-columns: 1fr 1fr; gap: 4px; }
  small { color: var(--ok-text-2); }
</style>

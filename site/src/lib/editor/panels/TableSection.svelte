<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import { selectedShapeId } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import { getTableCells, isTableShape, getTableCellText, getTableCellSpan, getTableCellFill, getTableColumnWidths, getTableRowHeights, getTableCellAlignment, getTableCellAnchor, setTableCellText, setTableCellFill, setTableCellAlignment, setTableCellAnchor, setTableRowHeight, setTableColumnWidth, insertTableRow, insertTableColumn, removeTableRow, removeTableColumn, inches } from '@office-kit/pptx';

  const editor = getEditor();
  const doc = editor.doc;
  const state = $derived.by(() => {
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
  function select(row: number, col: number) {
    if (state) doc.selectCell(doc.selection.slideIndex, state.id, row, col);
  }
  function structure(axis: 'row' | 'column', remove: boolean) {
    if (!state) return;
    const s = state;
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

{#if state}
  <section class="table-controls" aria-label={t('Table options')}>
    <strong>{t('Table options')}</strong>
    <div class="cell-grid ok-scroll">
      <table aria-label={t('Table cells')}><tbody>
        {#each state.cells as row, r}<tr>
          {#each row as cell, c}
            {@const span = getTableCellSpan(cell)}
            {#if !span.hMerge && !span.vMerge}
              <td rowspan={span.rowSpan} colspan={span.gridSpan}><button class="ok-btn" aria-label={`${t('Cell')} ${r + 1}, ${c + 1}`} aria-pressed={state.row === r && state.col === c} onclick={() => select(r, c)}>{getTableCellText(cell) || '—'}</button></td>
            {/if}
          {/each}
        </tr>{/each}
      </tbody></table>
    </div>
    {#if state.cell}
      <label>{t('Cell text')}<textarea class="ok-input" rows="3" value={getTableCellText(state.cell)} onchange={(e) => { const s = state; if (s?.cell) doc.transact(t('Edit cell text'), () => setTableCellText(s.cell!, e.currentTarget.value, { preserveFormatting: true })); }}></textarea></label>
      <label>{t('Cell fill')}<input type="color" value={getTableCellFill(state.cell) ?? '#ffffff'} onchange={(e) => { const cell = state?.cell; if (cell) doc.transact(t('Cell fill'), () => setTableCellFill(cell, e.currentTarget.value)); }} /></label>
      <label>{t('Horizontal alignment')}<select aria-label={t('Horizontal alignment')} class="ok-input" value={getTableCellAlignment(state.cell) ?? 'l'} onchange={(e) => { const cell = state?.cell; const v = e.currentTarget.value; if (cell && (v === 'l' || v === 'ctr' || v === 'r')) doc.transact(t('Horizontal alignment'), () => setTableCellAlignment(cell, v)); }}>
        <option value="l">{t('Left')}</option><option value="ctr">{t('Center')}</option><option value="r">{t('Right')}</option>
      </select></label>
      <label>{t('Vertical alignment')}<select aria-label={t('Vertical alignment')} class="ok-input" value={getTableCellAnchor(state.cell) ?? 'top'} onchange={(e) => { const cell = state?.cell; const v = e.currentTarget.value; if (cell && (v === 'top' || v === 'center' || v === 'bottom')) doc.transact(t('Vertical alignment'), () => setTableCellAnchor(cell, v)); }}>
        <option value="top">{t('Top')}</option><option value="center">{t('Center')}</option><option value="bottom">{t('Bottom')}</option>
      </select></label>
      <label>{t('Row height (inches)')}<input class="ok-input" type="number" min="0.01" step="0.01" required value={state.heights[state.row]! / inches(1)} onchange={(e) => { const s = state; if (s && e.currentTarget.reportValidity()) doc.transact(t('Resize table row'), () => setTableRowHeight(s.table, s.row, inches(e.currentTarget.valueAsNumber))); }} /></label>
      <label>{t('Column width (inches)')}<input class="ok-input" type="number" min="0.01" step="0.01" required value={state.widths[state.col]! / inches(1)} onchange={(e) => { const s = state; if (s && e.currentTarget.reportValidity()) doc.transact(t('Resize table column'), () => setTableColumnWidth(s.table, s.col, inches(e.currentTarget.valueAsNumber))); }} /></label>
      <div class="actions">
        <button class="ok-btn" disabled={state.merged} onclick={() => structure('row', false)}>{t('Insert row below')}</button>
        <button class="ok-btn" disabled={state.merged} onclick={() => structure('column', false)}>{t('Insert column right')}</button>
        <button class="ok-btn" disabled={state.merged || state.cells.length <= 1} onclick={() => structure('row', true)}>{t('Delete row')}</button>
        <button class="ok-btn" disabled={state.merged || state.widths.length <= 1} onclick={() => structure('column', true)}>{t('Delete column')}</button>
      </div>
      {#if state.merged}<small>{t('Split merged cells before changing rows or columns')}</small>{/if}
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

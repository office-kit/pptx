<script lang="ts">
  import { tick } from 'svelte';
  import { neighboringTableCell, tableSelectionBlock, tableCellsInRange } from '../core/table-selection.ts';
  import TextFormatBar from '../ui/TextFormatBar.svelte';
  import { getEditor } from '../core/context.ts';
  import { selectedShapeId } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import { getTableCellMargins, setTableCellMargins, getTableCellPosition, setTableCellBorders, type TableCellData, type TextFormat, getTableCellParagraphs, setTableCellTextFormat, mergeTableCells, splitTableCell, getTableCells, isTableShape, getTableCellText, getTableCellSpan, getTableCellFill, getTableColumnWidths, getTableRowHeights, getTableCellAlignment, getTableCellAnchor, setTableCellText, setTableCellFill, setTableCellAlignment, setTableCellAnchor, setTableRowHeight, setTableColumnWidth, insertTableRow, insertTableColumn, removeTableRow, removeTableColumn, inches } from '@office-kit/pptx';

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
  const block = $derived(tableState ? tableSelectionBlock(doc.selection.kind === 'cell' ? doc.selection : tableState) : null);
  const selectedCells = $derived(tableState && block ? tableCellsInRange(tableState.cells, block) : new Set<TableCellData>());
  const textRuns = $derived([...selectedCells].flatMap(cell => getTableCellParagraphs(cell).flatMap(paragraph => paragraph.elements.filter(element => element.kind !== 'br'))));
  function applyToCells(label: string, edit: (cell: TableCellData) => void) {
    const cells = [...selectedCells];
    if (cells.length) doc.transact(label, () => { for (const cell of cells) edit(cell); });
  }
  function formatCells(format: TextFormat) {
    applyToCells(t('Format selected cells'), cell => setTableCellTextFormat(cell, format));
  }
  const EMU_PER_POINT = 12700;
  // OOXML cell margins use signed 32-bit EMU coordinates.
  const MAX_MARGIN_POINTS = 2147483647 / EMU_PER_POINT;
  const marginSides = [
    { key: 'left', label: 'Left margin (points)', reset: 'Reset left margin' },
    { key: 'right', label: 'Right margin (points)', reset: 'Reset right margin' },
    { key: 'top', label: 'Top margin (points)', reset: 'Reset top margin' },
    { key: 'bottom', label: 'Bottom margin (points)', reset: 'Reset bottom margin' },
  ] as const;
  const cellMargins = $derived([...selectedCells].filter(cell => {
    const span = getTableCellSpan(cell);
    return !span.hMerge && !span.vMerge;
  }).map(getTableCellMargins));
  function marginValue(side: typeof marginSides[number]['key']): number | null | undefined {
    const first = cellMargins[0]?.[side];
    return cellMargins.every(margins => margins[side] === first) ? first : undefined;
  }
  function changeMargin(input: HTMLInputElement, side: typeof marginSides[number]['key']) {
    if (!input.reportValidity()) return;
    const value = input.value === '' ? null : Math.round(input.valueAsNumber * EMU_PER_POINT);
    applyToCells(t('Cell margins'), cell => setTableCellMargins(cell, { ...getTableCellMargins(cell), [side]: value }));
  }
  let borderColor = $state('#000000');
  let borderWidth = $state(1);
  let borderDash = $state('solid');
  let borderMode = $state('all');
  function applyBorders() {
    const current = tableState;
    const selected = selectedCells;
    if (!current) return;
    const line = { color: borderColor, widthEmu: Math.round(borderWidth * EMU_PER_POINT), dash: borderDash };
    applyToCells(t('Apply borders'), cell => {
      const { row, col } = getTableCellPosition(cell);
      const span = getTableCellSpan(cell);
      const outside = (r: number, c: number) => { const neighbor = current.cells[r]?.[c]; return !neighbor || !selected.has(neighbor); };
      setTableCellBorders(cell, {
        left: borderMode === 'all' || outside(row, col - 1) ? line : undefined,
        right: borderMode === 'all' || outside(row, col + span.gridSpan) ? line : undefined,
        top: borderMode === 'all' || outside(row - 1, col) ? line : undefined,
        bottom: borderMode === 'all' || outside(row + span.rowSpan, col) ? line : undefined,
      });
    });
  }
  const canMerge = $derived.by(() => {
    if (!tableState || !block || block.rowSpan * block.colSpan < 2 || block.row + block.rowSpan > tableState.cells.length || block.col + block.colSpan > tableState.widths.length) return false;
    return tableState.cells.slice(block.row, block.row + block.rowSpan).every(row => row.slice(block.col, block.col + block.colSpan).every(cell => {
      const span = getTableCellSpan(cell);
      return !span.hMerge && !span.vMerge && span.rowSpan === 1 && span.gridSpan === 1;
    }));
  });
  function select(row: number, col: number, extend: boolean) {
    if (!tableState) return;
    doc.selectCell(doc.selection.slideIndex, tableState.id, row, col, extend);
  }
  function onCellContext(event: MouseEvent, row: number, col: number) {
    event.preventDefault();
    if (!tableState) return;
    const cell = tableState.cells[row]?.[col];
    if (doc.selection.kind !== 'cell' || !cell || !selectedCells.has(cell)) select(row, col, false);
    editor.openContextMenu(event.clientX, event.clientY);
  }
  async function onCellKeydown(event: KeyboardEvent, row: number, col: number) {
    if (event.isComposing || event.ctrlKey || event.metaKey || event.altKey || !tableState) return;
    if (event.key === 'Delete' || event.key === 'Backspace') {
      if (doc.selection.kind !== 'cell') return;
      event.preventDefault();
      event.stopPropagation();
      editor.clearCellText();
      return;
    }
    if (!['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key)) return;
    event.preventDefault();
    event.stopPropagation();
    const next = neighboringTableCell(tableState.table, row, col,
      event.key === 'ArrowUp' ? -1 : event.key === 'ArrowDown' ? 1 : 0,
      event.key === 'ArrowLeft' ? -1 : event.key === 'ArrowRight' ? 1 : 0);
    if (!next) return;
    const grid = (event.currentTarget as HTMLElement).closest('.cell-grid');
    select(next.row, next.col, event.shiftKey);
    await tick();
    grid?.querySelector<HTMLButtonElement>(`[data-cell="${next.row},${next.col}"]`)?.focus();
  }
  function merge() {
    const s = tableState;
    const b = block;
    if (!s || !b || !canMerge) return;
    doc.transact(t('Merge cells'), () => {
      mergeTableCells(s.table, b, { coveredText: 'append' });
      doc.selectCell(doc.selection.slideIndex, s.id, b.row, b.col);
    });
  }
  function split() {
    const cell = tableState?.cell;
    if (cell) doc.transact(t('Split cell'), () => splitTableCell(cell));
  }
  function structure(axis: 'row' | 'column', remove: boolean) {
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
              <td rowspan={span.rowSpan} colspan={span.gridSpan}><button class="ok-btn" oncontextmenu={(e) => onCellContext(e, r, c)} onkeydown={(e) => onCellKeydown(e, r, c)} data-cell={`${r},${c}`} aria-label={`${t('Cell')} ${r + 1}, ${c + 1}`} aria-pressed={selectedCells.has(cell)} onclick={(e) => select(r, c, e.shiftKey)}>{getTableCellText(cell) || '—'}</button></td>
            {/if}
          {/each}
        </tr>{/each}
      </tbody></table>
    </div>
    <small>{t('Shift-click another cell to select a range')}</small>
    <small>{t('Arrow keys move between cells. Shift+Arrow extends the range. Delete clears text.')}</small>
    {#if tableState.cell}
      <div class="actions">
        <button class="ok-btn" disabled={!canMerge} onclick={merge}>{t('Merge cells')}</button>
        <button class="ok-btn" disabled={getTableCellSpan(tableState.cell).gridSpan === 1 && getTableCellSpan(tableState.cell).rowSpan === 1} onclick={split}>{t('Split cell')}</button>
      </div>
      <label>{t('Cell text')}<textarea class="ok-input" rows="3" value={getTableCellText(tableState.cell)} onchange={(e) => { const s = tableState; if (s?.cell) doc.transact(t('Edit cell text'), () => setTableCellText(s.cell!, e.currentTarget.value, { preserveFormatting: true })); }}></textarea></label>
      <TextFormatBar formats={textRuns.map(run => run.format ?? {})} selected={selectedCells.size > 0} onformat={formatCells} context="cells" />
      <label>{t('Cell fill')}<input type="color" value={getTableCellFill(tableState.cell) ?? '#ffffff'} onchange={(e) => { const value = e.currentTarget.value; applyToCells(t('Cell fill'), cell => setTableCellFill(cell, value)); }} /></label>
      <label>{t('Horizontal alignment')}<select aria-label={t('Horizontal alignment')} class="ok-input" value={getTableCellAlignment(tableState.cell) ?? 'l'} onchange={(e) => { const v = e.currentTarget.value; if (v === 'l' || v === 'ctr' || v === 'r') applyToCells(t('Horizontal alignment'), cell => setTableCellAlignment(cell, v)); }}>
        <option value="l">{t('Left')}</option><option value="ctr">{t('Center')}</option><option value="r">{t('Right')}</option>
      </select></label>
      <label>{t('Vertical alignment')}<select aria-label={t('Vertical alignment')} class="ok-input" value={getTableCellAnchor(tableState.cell) ?? 'top'} onchange={(e) => { const v = e.currentTarget.value; if (v === 'top' || v === 'center' || v === 'bottom') applyToCells(t('Vertical alignment'), cell => setTableCellAnchor(cell, v)); }}>
        <option value="top">{t('Top')}</option><option value="center">{t('Center')}</option><option value="bottom">{t('Bottom')}</option>
      </select></label>
      <details>
        <summary>{t('Cell margins')}</summary>
        <div class="margin-fields">
          {#each marginSides as side}
            {@const value = marginValue(side.key)}
            <div class="margin-side"><label>{t(side.label)}<input class="ok-input" type="number" min="0" max={MAX_MARGIN_POINTS} step="any" value={value == null ? '' : value / EMU_PER_POINT} placeholder={t(value === undefined ? 'Mixed' : 'Default')} onchange={(e) => changeMargin(e.currentTarget, side.key)} /></label>
            <button class="ok-btn" aria-label={t(side.reset)} disabled={cellMargins.every(margins => margins[side.key] === null)} onclick={() => applyToCells(t(side.reset), cell => setTableCellMargins(cell, { ...getTableCellMargins(cell), [side.key]: null }))}>{t('Default')}</button></div>
          {/each}
          <small>{t('Leave a margin blank to use the default')}</small>
          <button class="ok-btn" disabled={!cellMargins.some(margins => Object.values(margins).some(value => value !== null))} onclick={() => applyToCells(t('Reset cell margins'), cell => setTableCellMargins(cell, null))}>{t('Reset cell margins')}</button>
        </div>
      </details>
      <details>
        <summary>{t('Cell borders')}</summary>
        <form onsubmit={(e) => { e.preventDefault(); applyBorders(); }}>
          <label>{t('Border color')}<input type="color" bind:value={borderColor} /></label>
          <label>{t('Border width (points)')}<input class="ok-input" type="number" min="0.25" max="100" step="0.25" required bind:value={borderWidth} /></label>
          <label>{t('Border style')}<select aria-label={t('Border style')} class="ok-input" bind:value={borderDash}><option value="solid">{t('Solid line')}</option><option value="dash">{t('Dashed line')}</option><option value="dot">{t('Dotted line')}</option></select></label>
          <label>{t('Border placement')}<select aria-label={t('Border placement')} class="ok-input" bind:value={borderMode}><option value="all">{t('All cell borders')}</option><option value="outer">{t('Outside borders')}</option></select></label>
          <div class="actions"><button class="ok-btn" type="submit">{t('Apply borders')}</button><button class="ok-btn" type="button" onclick={() => applyToCells(t('Reset borders'), cell => setTableCellBorders(cell, null))}>{t('Reset borders')}</button></div>
        </form>
      </details>
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
  form, .margin-fields { display: grid; gap: 8px; margin-top: 8px; }
  summary { cursor: pointer; font-size: 12px; }
  strong { font-size: 12px; }
  label { display: grid; gap: 4px; font-size: 11px; }
  textarea, input, select { width: 100%; box-sizing: border-box; }
  .margin-side { display: grid; grid-template-columns: minmax(0, 1fr) auto; align-items: end; gap: 6px; }
  .cell-grid { overflow: auto; max-height: 220px; }
  table { border-collapse: collapse; width: 100%; }
  td { border: 1px solid var(--ok-border); padding: 1px; }
  td button { width: 100%; min-width: 50px; max-width: 130px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  button[aria-pressed='true'] { background: var(--ok-accent); color: white; }
  .actions { display: grid; grid-template-columns: 1fr 1fr; gap: 4px; }
  small { color: var(--ok-text-2); }
</style>

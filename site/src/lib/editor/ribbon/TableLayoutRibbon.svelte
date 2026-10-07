<script lang="ts">
  // Mac PowerPoint's Table Layout tab: Table (Select ▾, View Gridlines), Rows
  // & Columns (Delete ▾, Insert Row Above/Below, Insert Column Left/Right),
  // Merge (Merge Cells, Split Cells), Cell Size (Row, Column, Distribute Rows,
  // Distribute Columns), Alignment (six toggles, Text Direction ▾, Cell
  // Margins ▾), Table Size, Arrange and Format Pane. Below 1300 pt three of
  // the insert buttons become rows, Distribute loses its labels and Arrange
  // collapses into one button.
  import './contextual.css';
  import { emu, getTableCellAlignment, getTableCellAnchor, getTableCellSpan, getTableCellTextDirection, insertTableColumn, insertTableRow, mergeTableCells, removeTableColumn, removeTableRow, setShapeBounds, setTableCellAlignment, setTableCellAnchor, setTableCellMargins, setTableCellTextDirection, setTableColumnWidth, setTableRowHeight, splitTableCell, cm, type ShapeBounds, type SlideShapeData, type TableCellData } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { fitTableFrame, tableTarget } from '../core/table-target.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import ArrangeGroup from './ArrangeGroup.svelte';
  import MenuButton from './MenuButton.svelte';
  import SizeSpinners from './SizeSpinners.svelte';
  import { caption } from './caption.ts';
  import { RibbonCollapse } from './ribbon-collapse.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const collapse = new RibbonCollapse(1);
  const compact = $derived(collapse.step >= 1);
  const target = $derived.by(() => { doc.version; doc.selection; return tableTarget(editor); });
  const editable = $derived(target !== null && !editor.selectionLocked());
  const anchorState = $derived.by(() => {
    doc.version;
    const cell = target?.anchor;
    return cell ? { align: getTableCellAlignment(cell) ?? 'l', anchor: getTableCellAnchor(cell) ?? 'top', direction: getTableCellTextDirection(cell) ?? 'horz' } : null;
  });
  const canMerge = $derived.by(() => {
    const current = target;
    if (!current || current.block.rowSpan * current.block.colSpan < 2) return false;
    const { block, cells } = current;
    return cells.slice(block.row, block.row + block.rowSpan).every((row) => row.slice(block.col, block.col + block.colSpan).every((cell) => {
      const span = getTableCellSpan(cell);
      return !span.hMerge && !span.vMerge && span.rowSpan === 1 && span.gridSpan === 1;
    }));
  });
  const canSplit = $derived.by(() => {
    const cell = target?.anchor;
    if (!cell) return false;
    const span = getTableCellSpan(cell);
    return span.gridSpan > 1 || span.rowSpan > 1;
  });
  const rowHeight = $derived(target ? common(target.heights.slice(target.block.row, target.block.row + target.block.rowSpan)) : null);
  const columnWidth = $derived(target ? common(target.widths.slice(target.block.col, target.block.col + target.block.colSpan)) : null);
  const MERGED_REASON = 'Split merged cells before changing rows or columns';
  // PowerPoint's Cell Margins presets, in EMU (Normal is the default inset).
  const MARGINS = [
    ['Normal', null],
    ['None', { left: 0, right: 0, top: 0, bottom: 0 }],
    ['Narrow', { left: 45720, right: 45720, top: 45720, bottom: 45720 }],
    ['Wide', { left: 137160, right: 137160, top: 137160, bottom: 137160 }],
  ] as const;
  const DIRECTIONS = [
    ['horz', 'Horizontal'],
    ['vert', 'Rotate all text 90°'],
    ['vert270', 'Rotate all text 270°'],
    ['wordArtVert', 'Stacked'],
  ] as const;

  function common(values: readonly number[]): number | null {
    return values.length > 0 && values.every((value) => value === values[0]) ? Math.round((values[0]! / cm(1)) * 100) / 100 : null;
  }
  function edit(label: string, change: (table: SlideShapeData) => void) {
    const current = target;
    if (current && editable) doc.transact(t(label), () => change(current.table));
  }
  function cells(label: string, change: (cell: TableCellData) => void) {
    const current = target;
    if (current && editable) doc.transact(t(label), () => { for (const cell of current.selected) change(cell); });
  }
  function select(kind: 'cell' | 'row' | 'column' | 'table') {
    const current = target;
    if (!current) return;
    const { block, slideIndex, tableId, cells: grid, widths } = current;
    if (kind === 'table') { doc.selectShape(slideIndex, tableId); return; }
    const from = kind === 'column' ? { row: 0, col: block.col } : kind === 'row' ? { row: block.row, col: 0 } : block;
    const to = kind === 'column' ? { row: grid.length - 1, col: block.col + block.colSpan - 1 } : kind === 'row' ? { row: block.row + block.rowSpan - 1, col: widths.length - 1 } : block;
    doc.selectCell(slideIndex, tableId, from.row, from.col);
    if (kind !== 'cell') doc.selectCell(slideIndex, tableId, to.row, to.col, true);
  }
  function insertRow(below: boolean) {
    const current = target;
    if (!current) return;
    const at = below ? current.block.row + current.block.rowSpan : current.block.row;
    edit(below ? 'Insert Row Below' : 'Insert Row Above', (table) => {
      insertTableRow(table, at);
      fitTableFrame(doc.pres, table);
      doc.selectCell(current.slideIndex, current.tableId, at, current.block.col);
    });
  }
  function insertColumn(right: boolean) {
    const current = target;
    if (!current) return;
    const at = right ? current.block.col + current.block.colSpan : current.block.col;
    edit(right ? 'Insert Column Right' : 'Insert Column Left', (table) => {
      insertTableColumn(table, at);
      fitTableFrame(doc.pres, table);
      doc.selectCell(current.slideIndex, current.tableId, current.block.row, at);
    });
  }
  function remove(kind: 'rows' | 'columns' | 'table') {
    const current = target;
    if (!current) return;
    if (kind === 'table') {
      doc.selectShape(current.slideIndex, current.tableId);
      editor.deleteSelection();
      return;
    }
    const { block, cells: grid, widths } = current;
    edit(kind === 'rows' ? 'Delete Rows' : 'Delete Columns', (table) => {
      // Remove from the end so earlier indices stay valid.
      if (kind === 'rows') for (let row = block.row + block.rowSpan - 1; row >= block.row; row--) removeTableRow(table, row);
      else for (let col = block.col + block.colSpan - 1; col >= block.col; col--) removeTableColumn(table, col);
      fitTableFrame(doc.pres, table);
      const rows = grid.length - (kind === 'rows' ? block.rowSpan : 0);
      const cols = widths.length - (kind === 'columns' ? block.colSpan : 0);
      doc.selectCell(current.slideIndex, current.tableId, Math.min(block.row, rows - 1), Math.min(block.col, cols - 1));
    });
  }
  function merge() {
    const current = target;
    if (!current || !canMerge) return;
    edit('Merge Cells', (table) => {
      mergeTableCells(table, current.block, { coveredText: 'append' });
      doc.selectCell(current.slideIndex, current.tableId, current.block.row, current.block.col);
    });
  }
  function split() {
    const cell = target?.anchor;
    if (cell && canSplit) edit('Split Cells', () => splitTableCell(cell));
  }
  function resizeTracks(kind: 'row' | 'column', input: HTMLInputElement) {
    const current = target;
    const restore = () => { input.value = String((kind === 'row' ? rowHeight : columnWidth) ?? ''); };
    if (!current || !editable || !input.reportValidity() || !(input.valueAsNumber > 0)) { restore(); return; }
    const value = emu(cm(input.valueAsNumber));
    const { block } = current;
    edit(kind === 'row' ? 'Table Row Height' : 'Table Column Width', (table) => {
      if (kind === 'row') for (let row = block.row; row < block.row + block.rowSpan; row++) setTableRowHeight(table, row, value);
      else for (let col = block.col; col < block.col + block.colSpan; col++) setTableColumnWidth(table, col, value);
      fitTableFrame(doc.pres, table);
    });
  }
  // Distribute evens out the selected rows (columns), or all of them when the
  // whole table is selected, keeping their total.
  function distribute(kind: 'row' | 'column') {
    const current = target;
    if (!current) return;
    const { block } = current;
    const start = kind === 'row' ? block.row : block.col;
    const count = kind === 'row' ? block.rowSpan : block.colSpan;
    const tracks = (kind === 'row' ? current.heights : current.widths).slice(start, start + count);
    const each = emu(tracks.reduce((sum, value) => sum + value, 0) / count);
    edit(kind === 'row' ? 'Distribute Rows' : 'Distribute Columns', (table) => {
      for (let index = start; index < start + count; index++) {
        if (kind === 'row') setTableRowHeight(table, index, each);
        else setTableColumnWidth(table, index, each);
      }
      fitTableFrame(doc.pres, table);
    });
  }
  // Table Size scales the rows and columns with the frame, as PowerPoint does.
  function resizeTable(table: SlideShapeData, size: ShapeBounds) {
    const current = target;
    if (!current || current.table !== table) { setShapeBounds(table, size); return; }
    const width = current.widths.reduce((sum, value) => sum + value, 0);
    const height = current.heights.reduce((sum, value) => sum + value, 0);
    current.widths.forEach((value, col) => setTableColumnWidth(table, col, emu((value * size.w) / width)));
    current.heights.forEach((value, row) => setTableRowHeight(table, row, emu((value * size.h) / height)));
    fitTableFrame(doc.pres, table);
  }
</script>

{#snippet selectItems()}
  <button role="menuitem" onclick={() => select('cell')}>{t('Select Cell')}</button>
  <button role="menuitem" onclick={() => select('column')}>{t('Select Column')}</button>
  <button role="menuitem" onclick={() => select('row')}>{t('Select Row')}</button>
  <button role="menuitem" onclick={() => select('table')}>{t('Select Table')}</button>
{/snippet}

{#snippet deleteItems()}
  <button role="menuitem" disabled={!target || target.merged || target.block.colSpan >= target.widths.length} title={target?.merged ? t(MERGED_REASON) : undefined} onclick={() => remove('columns')}>{t('Delete Columns')}</button>
  <button role="menuitem" disabled={!target || target.merged || target.block.rowSpan >= target.cells.length} title={target?.merged ? t(MERGED_REASON) : undefined} onclick={() => remove('rows')}>{t('Delete Rows')}</button>
  <button role="menuitem" onclick={() => remove('table')}>{t('Delete Table')}</button>
{/snippet}

{#snippet directionItems()}
  {#each DIRECTIONS as [value, label] (value)}<button role="menuitemradio" aria-checked={anchorState?.direction === value} onclick={() => cells('Text Direction', (cell) => setTableCellTextDirection(cell, value === 'horz' ? null : value))}>{t(label)}<span>{anchorState?.direction === value ? '✓' : ''}</span></button>{/each}
{/snippet}

{#snippet marginItems()}
  {#each MARGINS as [label, margins] (label)}<button role="menuitem" onclick={() => cells('Cell Margins', (cell) => setTableCellMargins(cell, margins))}>{t(label)}</button>{/each}
  <hr />
  <button role="menuitem" onclick={() => editor.showShapeFormat()}>{t('Custom Margins...')}</button>
{/snippet}

{#snippet insertButton(name: string, icon: string, run: () => void, row: boolean)}
  {#if row}
    <button class="ctx-row" aria-label={t(name)} disabled={!editable || !!target?.merged} title={target?.merged ? t(MERGED_REASON) : t(name)} onclick={run}><Icon name={icon} size={16} /><span>{t(name)}</span></button>
  {:else}
    <button class="ctx-big" aria-label={t(name)} disabled={!editable || !!target?.merged} title={target?.merged ? t(MERGED_REASON) : t(name)} onclick={run}><span class="ctx-icon-row"><Icon name={icon} size={32} /></span><span class="ctx-caption">{caption(t(name))}</span></button>
  {/if}
{/snippet}

<div class="ctx-ribbon table-layout" bind:this={collapse.node} bind:clientWidth={collapse.width}>
  <section class="ctx-group" aria-label={t('Table')}>
    <MenuButton look="big" icon="select-table" label={t('Select')} disabled={!target}>{@render selectItems()}</MenuButton>
    <button class="ctx-big" aria-label={t('View Gridlines')} aria-pressed="true" title={t('The editor always shows table gridlines.')} disabled><span class="ctx-icon-row"><Icon name="gridlines" size={32} /></span><span class="ctx-caption">{caption(t('View Gridlines'))}</span></button>
  </section>

  <section class="ctx-group" aria-label={t('Rows & Columns')}>
    <MenuButton look="big" icon="trash" label={t('Delete')} disabled={!editable}>{@render deleteItems()}</MenuButton>
    <span class="ctx-sep" aria-hidden="true"></span>
    {@render insertButton('Insert Row Above', 'insert-row-above', () => insertRow(false), false)}
    {#if compact}
      <div class="ctx-rows">
        {@render insertButton('Insert Row Below', 'insert-row-below', () => insertRow(true), true)}
        {@render insertButton('Insert Column Left', 'insert-column-left', () => insertColumn(false), true)}
        {@render insertButton('Insert Column Right', 'insert-column-right', () => insertColumn(true), true)}
      </div>
    {:else}
      {@render insertButton('Insert Row Below', 'insert-row-below', () => insertRow(true), false)}
      {@render insertButton('Insert Column Left', 'insert-column-left', () => insertColumn(false), false)}
      {@render insertButton('Insert Column Right', 'insert-column-right', () => insertColumn(true), false)}
    {/if}
  </section>

  <section class="ctx-group" aria-label={t('Merge')}>
    <button class="ctx-big" aria-label={t('Merge Cells')} disabled={!editable || !canMerge} onclick={merge}><span class="ctx-icon-row"><Icon name="merge" size={32} /></span><span class="ctx-caption">{caption(t('Merge Cells'))}</span></button>
    <button class="ctx-big" aria-label={t('Split Cells')} title={canSplit ? t('Split Cells') : t('Only merged cells can be split in this editor.')} disabled={!editable || !canSplit} onclick={split}><span class="ctx-icon-row"><Icon name="split" size={32} /></span><span class="ctx-caption">{caption(t('Split Cells'))}</span></button>
  </section>

  <section class="ctx-group" aria-label={t('Cell Size')}>
    <div class="ctx-spins track-spins">
      <label><span class="ctx-spin-label"><Icon name="height" size={18} /></span><input class="ok-input" type="number" min="0" step="any" aria-label={t('Table Row Height')} disabled={!editable} value={rowHeight ?? ''} placeholder={rowHeight === null && target ? t('Mixed') : undefined} onchange={(event) => resizeTracks('row', event.currentTarget)} /></label>
      <label><span class="ctx-spin-label"><Icon name="width" size={18} /></span><input class="ok-input" type="number" min="0" step="any" aria-label={t('Table Column Width')} disabled={!editable} value={columnWidth ?? ''} placeholder={columnWidth === null && target ? t('Mixed') : undefined} onchange={(event) => resizeTracks('column', event.currentTarget)} /></label>
    </div>
    <div class="ctx-rows tools distribute">
      {#each [['row', 'Distribute Rows', 'distribute-rows'], ['column', 'Distribute Columns', 'distribute-columns']] as const as [kind, name, icon] (kind)}
        {#if compact}
          <button class="ctx-small" aria-label={t(name)} title={t(name)} disabled={!editable} onclick={() => distribute(kind)}><Icon name={icon} size={18} /></button>
        {:else}
          <button class="ctx-row tall" aria-label={t(name)} disabled={!editable} onclick={() => distribute(kind)}><Icon name={icon} size={18} /><span>{t(name)}</span></button>
        {/if}
      {/each}
    </div>
  </section>

  <section class="ctx-group" aria-label={t('Alignment')}>
    <div class="ctx-rows tools">
      <div class="ctx-tool-row">
        {#each [['l', 'Align Left', 'align'], ['ctr', 'Center Text', 'align-text'], ['r', 'Align Right', 'align']] as const as [value, name, icon] (value)}
          <button class="ctx-small" aria-label={t(name)} title={t(name)} aria-pressed={anchorState?.align === value} disabled={!editable} onclick={() => cells(name, (cell) => setTableCellAlignment(cell, value))}><Icon name={icon} size={18} /></button>
        {/each}
      </div>
      <div class="ctx-tool-row">
        {#each [['top', 'Align Top', 'align-top'], ['center', 'Center Vertically', 'align-middle'], ['bottom', 'Align Bottom', 'align-bottom']] as const as [value, name, icon] (value)}
          <button class="ctx-small" aria-label={t(name)} title={t(name)} aria-pressed={anchorState?.anchor === value} disabled={!editable} onclick={() => cells(name, (cell) => setTableCellAnchor(cell, value))}><Icon name={icon} size={18} /></button>
        {/each}
      </div>
    </div>
    <MenuButton look="big" icon="text-direction" label={t('Text Direction')} disabled={!editable}>{@render directionItems()}</MenuButton>
    <MenuButton look="big" icon="cell-margins" label={t('Cell Margins')} disabled={!editable}>{@render marginItems()}</MenuButton>
  </section>

  <section class="ctx-group" aria-label={t('Table Size')}>
    <SizeSpinners labels apply={resizeTable} />
  </section>

  <ArrangeGroup collapsed={compact} table />

  <section class="ctx-group" aria-label={t('Format Pane')}>
    <button class="ctx-big" aria-label={t('Format Pane')} title={t('Display the Format Pane')} disabled={!target} onclick={() => editor.showShapeFormat()}><span class="ctx-icon-row"><Icon name="format-pane" size={32} /></span><span class="ctx-caption">{caption(t('Format Pane'))}</span></button>
  </section>
</div>

<style>
  .ctx-sep { width: 1px; margin: 8px 7px 14px; background: var(--ok-border); }
  .track-spins { padding-top: 5px; }
  .distribute { padding-top: 4px; }
  .distribute .tall { height: 26px; }
</style>

<script lang="ts">
  // Mac PowerPoint's Table Design tab: Table Style Options (six check boxes),
  // Table Styles (style strip, Shading ▾, Borders ▾, Effects ▾), WordArt
  // Styles (Quick Styles, Text Fill, Text Outline, Text Effects) and Draw
  // Borders (Pen Style, Pen Weight, Pen Color, Draw Table, Eraser). Below
  // 1300 pt Shading, Borders and Effects lose their labels.
  import './contextual.css';
  import { getPresentationTheme, getTableCellPosition, getTableCellSpan, getTableStyleFlags, getTableStyleId, setTableCellBorders, setTableCellFill, setTableCellTextFormat, setTableStyleFlags, setTableStyleId, type Color, type TableCellBorder, type TableCellData, type TextFormat } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { tableTarget } from '../core/table-target.ts';
  import { resolveColor } from '../core/theme-color.ts';
  import { applyTableCellWordArtPreset, type WordArtPreset } from '../core/wordart-presets.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import ColorPicker from '../ui/ColorPicker.svelte';
  import Icon from '../ui/Icon.svelte';
  import WordArtGallery from '../ui/WordArtGallery.svelte';
  import MenuButton from './MenuButton.svelte';
  import { captionLines } from './caption.ts';
  import { RibbonCollapse } from './ribbon-collapse.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const collapse = new RibbonCollapse(1);
  const compact = $derived(collapse.step >= 1);
  const target = $derived.by(() => { doc.version; doc.selection; return tableTarget(editor); });
  const editable = $derived(target !== null && !editor.selectionLocked());
  const flags = $derived.by(() => { doc.version; return target ? getTableStyleFlags(target.table) : null; });
  const styleId = $derived.by(() => { doc.version; return target ? getTableStyleId(target.table) : null; });
  const theme = $derived.by(() => { doc.version; return getPresentationTheme(doc.pres); });
  let quickStylesOpen = $state(false);
  let quickStylesButton = $state<HTMLButtonElement>();

  const OPTIONS = [
    ['firstRow', 'Header Row'],
    ['lastRow', 'Total Row'],
    ['bandRow', 'Banded Rows'],
    ['firstCol', 'First Column'],
    ['lastCol', 'Last Column'],
    ['bandCol', 'Banded Columns'],
  ] as const;

  // The built-in styles the editor can draw. PowerPoint's other built-in
  // styles are stored by GUID alone, and the renderer has no definitions for
  // them yet, so offering them would show an unstyled table here.
  const STYLES = [
    { id: '{2D5ABB26-0587-4C30-8999-92F81FD0307C}', name: 'No Style, No Grid', kind: 'none' },
    { id: '{5940675A-B579-460E-94D1-54222C63F5DA}', name: 'No Style, Table Grid', kind: 'grid' },
    { id: '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}', name: 'Medium Style 2 - Accent 1', kind: 'medium2' },
  ] as const;
  const accent = $derived(resolveColor('accent1', [], theme) ?? '#156082');
  const text = $derived(resolveColor('tx1', [], theme) ?? '#000000');

  const PEN_STYLES = [['solid', 'Solid Line'], ['dash', 'Dashed Line'], ['sysDot', 'Round Dot Line'], ['dashDot', 'Dash Dot Line'], ['lgDash', 'Long Dash Line']] as const;
  const PEN_WEIGHTS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6] as const;
  const EMU_PER_POINT = 12700;
  let penStyle = $state<(typeof PEN_STYLES)[number][0]>('solid');
  let penWeight = $state<number>(1);
  let penColor = $state<Color>('#000000');

  type Edge = 'top' | 'bottom' | 'left' | 'right' | 'insideH' | 'insideV' | 'tlToBr' | 'blToTr';
  const BORDER_ITEMS: readonly [string, readonly Edge[] | null][] = [
    ['No Border', null],
    ['All Borders', ['top', 'bottom', 'left', 'right', 'insideH', 'insideV']],
    ['Outside Borders', ['top', 'bottom', 'left', 'right']],
    ['Inside Borders', ['insideH', 'insideV']],
    ['Top Border', ['top']],
    ['Bottom Border', ['bottom']],
    ['Left Border', ['left']],
    ['Right Border', ['right']],
    ['Inside Horizontal Border', ['insideH']],
    ['Inside Vertical Border', ['insideV']],
    ['Diagonal Down Border', ['tlToBr']],
    ['Diagonal Up Border', ['blToTr']],
  ];

  function cells(): TableCellData[] { return target ? [...target.selected] : []; }
  function edit(label: string, change: (cell: TableCellData) => void) {
    const selected = cells();
    if (selected.length && editable) doc.transact(t(label), () => { for (const cell of selected) change(cell); });
  }
  function toggleOption(key: (typeof OPTIONS)[number][0], on: boolean) {
    const table = target?.table;
    if (table && editable) doc.transact(t('Table Style Options'), () => setTableStyleFlags(table, { [key]: on }));
  }
  function applyStyle(id: string) {
    const table = target?.table;
    if (table && editable) doc.transact(t('Table Styles'), () => setTableStyleId(table, id));
  }
  // PowerPoint's Borders ▾ applies the current pen to the chosen edges of the
  // selection; inside edges are the shared sides between selected cells.
  function borders(edges: readonly Edge[] | null) {
    const current = target;
    if (!current) return;
    if (!edges) { edit('No Border', (cell) => setTableCellBorders(cell, null)); return; }
    const line: Partial<TableCellBorder> = { color: penColor, widthEmu: Math.round(penWeight * EMU_PER_POINT), dash: penStyle };
    const { block } = current;
    const want = new Set(edges);
    edit('Borders', (cell) => {
      const { row, col } = getTableCellPosition(cell);
      const span = getTableCellSpan(cell);
      const top = row === block.row, bottom = row + span.rowSpan === block.row + block.rowSpan;
      const left = col === block.col, right = col + span.gridSpan === block.col + block.colSpan;
      const side = (outer: boolean, outerEdge: Edge, innerEdge: Edge) => (want.has(outer ? outerEdge : innerEdge) ? line : undefined);
      setTableCellBorders(cell, {
        top: side(top, 'top', 'insideH'),
        bottom: side(bottom, 'bottom', 'insideH'),
        left: side(left, 'left', 'insideV'),
        right: side(right, 'right', 'insideV'),
        tlToBr: want.has('tlToBr') ? line : undefined,
        blToTr: want.has('blToTr') ? line : undefined,
      });
    });
  }
  function textFormat(label: string, format: TextFormat) {
    edit(label, (cell) => setTableCellTextFormat(cell, format));
  }
  function wordArt(preset: WordArtPreset) {
    quickStylesOpen = false;
    edit('WordArt Styles', (cell) => applyTableCellWordArtPreset(cell, preset));
  }
  // PowerPoint's outer shadow, reflection and glow presets for text.
  const TEXT_SHADOW = { color: '#000000', blurEmu: 38100, offsetEmu: 38100, angleDeg: 45, opacity: 0.4 } as const;
  const TEXT_GLOW = { color: 'accent1', radiusEmu: 63500, opacity: 0.4 } as const;
  const TEXT_REFLECTION = { blurEmu: 6350, offsetEmu: 0, angleDeg: 90, startOpacity: 0.5, opacity: 0.003, endPosition: 0.55 } as const;
  const THIN_OUTLINE_EMU = 9525;
</script>

{#snippet borderItems()}
  {#each BORDER_ITEMS as [label, edges] (label)}<button role="menuitem" onclick={() => borders(edges)}>{t(label)}</button>{/each}
{/snippet}

{#snippet effectItems()}
  <button role="menuitem" title={t('Cell bevels are not supported by the library yet.')} disabled>{t('Cell Bevel')}</button>
{/snippet}

{#snippet textEffectItems()}
  <button role="menuitem" onclick={() => textFormat('Text Effects', { shadow: TEXT_SHADOW })}>{t('Shadow')}</button>
  <button role="menuitem" onclick={() => textFormat('Text Effects', { reflection: TEXT_REFLECTION })}>{t('Reflection')}</button>
  <button role="menuitem" onclick={() => textFormat('Text Effects', { glow: TEXT_GLOW })}>{t('Glow')}</button>
  <hr />
  <button role="menuitem" onclick={() => textFormat('Text Effects', { shadow: null, reflection: null, glow: null })}>{t('No Effects')}</button>
{/snippet}

<div class="ctx-ribbon table-design" bind:this={collapse.node} bind:clientWidth={collapse.width}>
  <section class="ctx-group" aria-label={t('Table Style Options')}>
    <div class="options">
      {#each OPTIONS as [key, label] (key)}
        <label class="option"><input type="checkbox" checked={flags?.[key] ?? false} disabled={!editable} onchange={(event) => toggleOption(key, event.currentTarget.checked)} /><span>{t(label)}</span></label>
      {/each}
    </div>
  </section>

  <section class="ctx-group ctx-shrink" aria-label={t('Table Styles')}>
    <div class="ctx-gallery" role="group" aria-label={t('Table Styles')}>
      <span class="ctx-gallery-arrow hidden" aria-hidden="true"></span>
      <div class="ctx-gallery-items table-strip">
        {#each STYLES as style (style.id)}
          <button class="table-swatch" class:current={styleId?.toUpperCase() === style.id} aria-label={t(style.name)} title={t(style.name)} aria-pressed={styleId?.toUpperCase() === style.id} disabled={!editable} onclick={() => applyStyle(style.id)}>
            <span class="mini {style.kind}" style:--accent={accent} style:--text={text} aria-hidden="true">
              {#each [0, 1, 2, 3, 4] as row (row)}<span class="mini-row" class:header={row === 0 && flags?.firstRow !== false} class:band={row % 2 === 1}></span>{/each}
            </span>
          </button>
        {/each}
      </div>
      <button class="ctx-gallery-arrow" aria-label={t('Next Table Styles gallery')} title={t('The editor draws only these built-in table styles.')} disabled>›</button>
    </div>
    <div class="ctx-rows">
      <span class="ctx-paint {compact ? 'ctx-icon' : 'ctx-row'}" class:disabled={!editable}><Icon name="fill" size={16} />{#if !compact}<span>{t('Shading')}</span>{/if}<ColorPicker compact label={t('Shading')} disabled={!editable} choose={(color) => edit('Shading', (cell) => setTableCellFill(cell, color))} /></span>
      <MenuButton look={compact ? 'icon' : 'row'} icon="border" label={t('Borders')} disabled={!editable}>{@render borderItems()}</MenuButton>
      <MenuButton look={compact ? 'icon' : 'row'} icon="shadow" label={t('Effects')} disabled={!editable}>{@render effectItems()}</MenuButton>
    </div>
  </section>

  <section class="ctx-group" aria-label={t('WordArt Styles')}>
    <div class="anchor">
      <button class="ctx-big" bind:this={quickStylesButton} aria-label={t('WordArt Quick Styles')} aria-haspopup="menu" aria-expanded={quickStylesOpen} disabled={!editable} onclick={() => (quickStylesOpen = !quickStylesOpen)}><span class="ctx-icon-row"><span class="wordart" aria-hidden="true">A</span><span class="ctx-caret" aria-hidden="true">▾</span></span><span class="ctx-caption">{captionLines(t('Quick Styles'))}</span></button>
      {#if quickStylesOpen && quickStylesButton}
        <WordArtGallery anchor={quickStylesButton} label={t('WordArt Quick Styles')} choose={wordArt} close={() => (quickStylesOpen = false)} />
      {/if}
    </div>
    <div class="ctx-rows">
      <span class="ctx-paint ctx-icon" class:disabled={!editable}><Icon name="font-color" size={16} /><ColorPicker compact label={t('Text Fill')} disabled={!editable} choose={(color) => textFormat('Text Fill', { color })} /></span>
      <span class="ctx-paint ctx-icon" class:disabled={!editable}><Icon name="outline" size={16} /><ColorPicker compact label={t('Text Outline')} disabled={!editable} choose={(color) => textFormat('Text Outline', { outline: { color, widthEmu: THIN_OUTLINE_EMU } })} /></span>
      <MenuButton look="icon" icon="glow" label={t('Text Effects')} disabled={!editable}>{@render textEffectItems()}</MenuButton>
    </div>
  </section>

  <section class="ctx-group" aria-label={t('Draw Borders')}>
    <div class="ctx-rows pens">
      <select class="ok-input" aria-label={t('Pen Style')} title={t('Pen Style')} bind:value={penStyle}>{#each PEN_STYLES as [value, label] (value)}<option {value}>{t(label)}</option>{/each}</select>
      <select class="ok-input" aria-label={t('Pen Weight')} title={t('Pen Weight')} bind:value={penWeight}>{#each PEN_WEIGHTS as weight (weight)}<option value={weight}>{weight} pt</option>{/each}</select>
      <span class="ctx-paint ctx-row pen-color"><Icon name="pen" size={16} /><span>{t('Pen Color')}</span><ColorPicker compact label={t('Pen Color')} choose={(color) => (penColor = color)} /></span>
    </div>
    <button class="ctx-big" aria-label={t('Draw Table')} title={t('Drawing table borders with the pointer is not available in this editor yet; use Borders.')} disabled><span class="ctx-icon-row"><Icon name="draw-table" size={32} /></span><span class="ctx-caption">{captionLines(t('Draw Table'))}</span></button>
    <button class="ctx-big" aria-label={t('Eraser')} title={t('Erasing table borders with the pointer is not available in this editor yet; use Borders ▸ No Border.')} disabled><span class="ctx-icon-row"><Icon name="eraser" size={32} /></span><span class="ctx-caption">{t('Eraser')}</span></button>
  </section>
</div>

<style>
  .options { display: grid; grid-template-rows: repeat(3, 22px); grid-auto-flow: column; grid-auto-columns: max-content; column-gap: 0; }
  /* Native check boxes: columns 97 and 114 pt wide. */
  .option { display: flex; align-items: center; gap: 5px; height: 22px; padding: 0 4px 0 0; font-size: 12px; white-space: nowrap; }
  .option input { margin: 0; }
  .ctx-gallery-items.table-strip { --ctx-gallery-min: 300px; width: 518px; justify-content: flex-start; gap: 0; background: #fff; }
  .table-swatch { width: 76px; height: 56px; margin-right: -2px; padding: 4px 6px; }
  .table-swatch.current { border-color: var(--ok-selected-border); background: var(--ok-selected); }
  .mini { display: flex; flex-direction: column; width: 100%; height: 100%; }
  .mini-row { flex: 1; border-top: 1px dashed #c8c8c8; }
  .mini.grid .mini-row { border: 1px solid var(--text); border-bottom: none; }
  .mini.grid .mini-row:last-child { border-bottom: 1px solid var(--text); }
  .mini.medium2 .mini-row { border-top: 1px solid #fff; background: color-mix(in srgb, var(--accent) 20%, #fff); }
  .mini.medium2 .mini-row.band { background: color-mix(in srgb, var(--accent) 40%, #fff); }
  .mini.medium2 .mini-row.header { background: var(--accent); }
  .anchor { position: relative; display: flex; }
  .wordart { font: 700 28px/32px Georgia, serif; color: var(--ok-accent); }
  .pens { gap: 0; }
  .pens select { box-sizing: border-box; width: 120px; height: 22px; padding: 0 4px; font-size: 11px; }
  .pen-color { width: 92px; box-sizing: border-box; }
</style>

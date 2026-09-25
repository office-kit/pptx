<script lang="ts">
  import { tick } from 'svelte';
  import { getShapeKind, getShapeParagraphCount, getParagraphPropertiesEffective, setParagraphLineSpacing, getTableCells, getTableCellParagraphs, getTableCellSpan, type ParagraphProperties } from '@office-kit/pptx';
  import { shapeTextDefaults } from '../core/text-layout-defaults.ts';
  import { getEditor } from '../core/context.ts';
  import { tableCellsInRange, tableSelectionBlock } from '../core/table-selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import ParagraphDialog from '../ui/ParagraphDialog.svelte';
  const editor = getEditor();
  const doc = editor.doc;
  let open = $state(false);
  let options = $state(false);
  let trigger: HTMLButtonElement;
  let menu = $state<HTMLDivElement>();
  const targets = $derived.by(() => {
    doc.version;
    const selection = doc.selection;
    if (selection.kind === 'cell') {
      const table = doc.shapeById(selection.slideIndex, selection.shapeId);
      return table ? [...tableCellsInRange(getTableCells(table), tableSelectionBlock(selection))].filter(cell => {
        const span = getTableCellSpan(cell);
        return !span.hMerge && !span.vMerge;
      }).flatMap(shape => getTableCellParagraphs(shape).map((_, index) => ({ shape, index, defaultAlign: 'left' as const }))) : [];
    }
    const shapes = editor.selectedShapes();
    return shapes.every(shape => getShapeKind(shape) === 'shape') ? shapes.flatMap(shape => Array.from({ length: getShapeParagraphCount(shape) }, (_, index) => ({ shape, index, defaultAlign: shapeTextDefaults(shape).align }))) : [];
  });
  const properties = $derived<ParagraphProperties[]>(editor.inlineTextFormat?.paragraphs ?? targets.map(({ shape, index, defaultAlign }) => { const props = getParagraphPropertiesEffective(doc.pres, shape, index); return { ...props, align: props.align ?? defaultAlign }; }));
  const enabled = $derived(properties.length > 0 && !editor.selectionLocked());
  const spacing = $derived.by(() => {
    const values = properties.map(p => !p.lineSpacing ? 1 : p.lineSpacing.kind === 'pct' ? p.lineSpacing.value : null);
    return values.every(value => value === values[0]) ? values[0] : null;
  });
  function apply(edit: (shape: Parameters<typeof setParagraphLineSpacing>[0], index: number) => void) {
    if (!enabled) return;
    if (editor.inlineTextFormat) editor.inlineTextFormat.editParagraphs(edit);
    else doc.transact(t('Format paragraphs'), () => { for (const { shape, index } of targets) edit(shape, index); });
  }
  function close(restore = true) { open = false; if (restore) trigger.focus(); }
  async function show() { open = !open; if (open) { await tick(); (menu?.querySelector<HTMLButtonElement>('[aria-checked="true"]') ?? menu?.querySelector<HTMLButtonElement>('button'))?.focus(); } }
  function place(node: HTMLElement) {
    const bounds = trigger.getBoundingClientRect();
    node.style.left = `${Math.max(8, Math.min(bounds.left, innerWidth - node.offsetWidth - 8))}px`;
    node.style.top = `${Math.max(8, Math.min(bounds.bottom, innerHeight - node.offsetHeight - 8))}px`;
  }
  function keys(event: KeyboardEvent) {
    event.stopPropagation();
    if (event.key === 'Escape') { event.preventDefault(); close(); return; }
    if (event.key === 'Tab') { close(false); return; }
    if (!['ArrowUp', 'ArrowDown', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault();
    const items = [...menu!.querySelectorAll<HTMLButtonElement>('button')];
    const index = items.indexOf(event.target as HTMLButtonElement);
    items[event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : (index + (event.key === 'ArrowDown' ? 1 : -1) + items.length) % items.length]?.focus();
  }
  $effect(() => { if (!enabled) { open = false; options = false; } });
</script>
<svelte:window onpointerdown={event => { if (open && !menu?.contains(event.target as Node) && !trigger.contains(event.target as Node)) close(false); }} onblur={() => { if (open) close(false); }} onresize={() => { if (open) close(false); }} />
<button class="trigger ok-btn" bind:this={trigger} aria-label={t('Line spacing')} title={t('Line spacing')} aria-haspopup="menu" aria-expanded={open} disabled={!enabled} onclick={show}><Icon name="line-spacing" /><span>▾</span></button>
{#if open}
  <div class="menu" role="menu" aria-label={t('Line spacing')} tabindex="-1" bind:this={menu} use:place onkeydown={keys}>
    {#each [1, 1.5, 2, 2.5, 3] as value}<button role="menuitemradio" aria-label={value.toFixed(1)} aria-checked={spacing === value} onclick={() => { apply((shape, index) => setParagraphLineSpacing(shape, index, { kind: 'pct', value })); close(); }}><span class="check">{spacing === value ? '✓' : ''}</span>{value.toFixed(1)}</button>{/each}
    <button role="menuitem" onclick={() => { close(false); options = true; }}>{t('Line Spacing Options...')}</button>
  </div>
{/if}
{#if options}<ParagraphDialog {properties} {apply} onclose={() => { options = false; trigger.focus(); }} />{/if}
<style>
  .trigger { display: flex; align-items: center; gap: 2px; padding: 3px; }
  .menu { position: fixed; z-index: 400; min-width: 170px; padding: 4px; background: var(--ok-panel); color: var(--ok-text); border: 1px solid var(--ok-border); border-radius: 5px; box-shadow: var(--ok-shadow-lg); }
  .menu button { display: flex; gap: 6px; width: 100%; border: 0; padding: 5px 8px; background: transparent; color: inherit; text-align: left; font: inherit; }
  .check { width: 14px; }
  .menu button:hover, .menu button:focus-visible { background: var(--ok-accent); color: white; outline: none; }
</style>

<script lang="ts">
  import { eventTarget } from '../core/dom-root.ts';
  // Home ▸ Paragraph's Columns, Text Direction and Align Text menus and
  // Convert to SmartArt, acting on every selected text-bearing shape.
  import { placeBelowTrigger } from './place-menu.ts';
  import {
    getShapeKind,
    getShapeTextAnchor,
    getShapeTextDirection,
    setShapeTextAnchor,
    setShapeTextColumns,
    setShapeTextDirection,
  } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import TextColumnsDialog from '../ui/TextColumnsDialog.svelte';

  // The reference desktop app puts Columns on the paragraph group's first row and the rest on
  // the second, after the alignment buttons.
  let { row }: { row: 1 | 2 } = $props();
  const editor = getEditor();
  const doc = editor.doc;
  const shapes = $derived.by(() => { doc.version; return editor.selectedShapes().filter((shape) => getShapeKind(shape) === 'shape'); });
  const direction = $derived.by(() => { doc.version; return shapes[0] ? getShapeTextDirection(shapes[0]) ?? 'horz' : null; });
  const anchor = $derived.by(() => { doc.version; return shapes[0] ? getShapeTextAnchor(shapes[0]) : null; });
  let open = $state<'columns' | 'direction' | 'align' | null>(null);
  let moreColumns = $state(false);

  const DIRECTIONS = [
    ['horz', 'Horizontal'],
    ['vert', 'Rotate all text 90°'],
    ['vert270', 'Rotate all text 270°'],
    ['wordArtVert', 'Stacked'],
  ] as const;
  const ANCHORS = [
    ['top', 'Top'],
    ['center', 'Middle'],
    ['bottom', 'Bottom'],
  ] as const;

  function apply(label: string, edit: (shape: (typeof shapes)[number]) => void) {
    open = null;
    doc.transact(t(label), () => { for (const shape of shapes) edit(shape); });
  }
</script>

<svelte:window onpointerdown={(event) => { if (open && !(eventTarget(event) as Element).closest?.('.paragraph-layout')) open = null; }} onkeydown={(event) => { if (event.key === 'Escape') open = null; }} />

{#snippet trigger(name: 'columns' | 'direction' | 'align', icon: string, label: string)}
  <button class="tool menu-trigger" aria-label={t(label)} title={t(label)} aria-haspopup="menu" aria-expanded={open === name} disabled={shapes.length === 0} onclick={() => (open = open === name ? null : name)}><Icon name={icon} size={18} /><span aria-hidden="true">⌄</span></button>
{/snippet}

{#if row === 1}
<span class="paragraph-layout">
  {@render trigger('columns', 'columns', 'Columns')}
  {#if open === 'columns'}
    <div class="menu" role="menu" use:placeBelowTrigger aria-label={t('Columns')}>
      {#each [1, 2, 3] as count (count)}
        <button role="menuitem" onclick={() => apply('Columns', (shape) => setShapeTextColumns(shape, count === 1 ? null : { count }))}>{t(['One Column', 'Two Columns', 'Three Columns'][count - 1]!)}</button>
      {/each}
      <hr />
      <button role="menuitem" onclick={() => { open = null; moreColumns = true; }}>{t('More Columns...')}</button>
    </div>
  {/if}
</span>
{:else}
<span class="paragraph-layout">
  {@render trigger('direction', 'text-direction', 'Text Direction')}
  {#if open === 'direction'}
    <div class="menu" role="menu" use:placeBelowTrigger aria-label={t('Text Direction')}>
      {#each DIRECTIONS as [value, label] (value)}
        <button role="menuitemradio" aria-checked={direction === value} onclick={() => apply('Text Direction', (shape) => setShapeTextDirection(shape, value))}>{t(label)}</button>
      {/each}
    </div>
  {/if}
</span>
<span class="paragraph-layout">
  {@render trigger('align', 'align-text', 'Align Text')}
  {#if open === 'align'}
    <div class="menu" role="menu" use:placeBelowTrigger aria-label={t('Align Text')}>
      {#each ANCHORS as [value, label] (value)}
        <button role="menuitemradio" aria-checked={anchor === value} onclick={() => apply('Align Text', (shape) => setShapeTextAnchor(shape, value))}>{t(label)}</button>
      {/each}
    </div>
  {/if}
</span>
<button class="tool" aria-label={t('Convert to SmartArt')} title={t('SmartArt is not supported by the library yet.')} disabled><Icon name="smartart" size={18} /></button>
{/if}
{#if moreColumns}<TextColumnsDialog onclose={() => (moreColumns = false)} />{/if}

<style>
  .paragraph-layout { position: relative; display: inline-flex; }
  button { font: inherit; color: var(--ok-text); background: none; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  button:hover:not(:disabled) { background: var(--ok-hover); }
  button:disabled { opacity: 0.4; cursor: default; }
  .tool { display: flex; align-items: center; justify-content: center; gap: 1px; min-width: 38px; height: 26px; padding: 0 2px; font-size: 10px; }
  .tool:not(.menu-trigger) { min-width: 26px; }
  .menu { position: fixed; z-index: 400; display: flex; flex-direction: column; min-width: 190px; padding: 4px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .menu button { padding: 5px 8px; text-align: left; font-size: 12px; }
  .menu button[aria-checked='true'] { background: var(--ok-selected); }
  hr { width: 100%; border: none; border-top: 1px solid var(--ok-border); margin: 4px 0; }
</style>

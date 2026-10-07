<script lang="ts">
  // PowerPoint's WordArt gallery: twenty "A" swatches, five to a row, shared by
  // Shape Format ▸ WordArt Quick Styles and Insert ▸ WordArt.
  import { onMount } from 'svelte';
  import { getPresentationTheme } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { wordArtSwatchStyle } from '../core/wordart-swatch.ts';
  import { WORDART_PRESETS, type WordArtPreset } from '../core/wordart-presets.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let { label, choose, close, clear, anchor }: {
    label: string;
    choose: (preset: WordArtPreset) => void;
    close: () => void;
    /** Clear WordArt, under the swatches; Insert ▸ WordArt has none. */
    clear?: () => void;
    /** The button the gallery opens under. */
    anchor: HTMLElement;
  } = $props();

  const editor = getEditor();
  const theme = $derived.by(() => { editor.doc.version; return getPresentationTheme(editor.doc.pres); });
  const COLUMNS = 5;
  let menu = $state<HTMLDivElement>();

  function keys(event: KeyboardEvent) {
    event.stopPropagation();
    if (event.key === 'Escape') { event.preventDefault(); close(); return; }
    if (event.key === 'Tab') { close(); return; }
    const offsets: Record<string, number> = { ArrowLeft: -1, ArrowRight: 1, ArrowUp: -COLUMNS, ArrowDown: COLUMNS };
    if (!(event.key in offsets) && event.key !== 'Home' && event.key !== 'End') return;
    event.preventDefault();
    const items = [...menu!.querySelectorAll<HTMLButtonElement>('[role="menuitem"]')];
    if (event.key === 'Home' || event.key === 'End') { (event.key === 'Home' ? items[0] : items.at(-1))?.focus(); return; }
    // Down from the last row reaches Clear WordArt below the grid.
    const index = items.indexOf(event.target as HTMLButtonElement) + offsets[event.key]!;
    (items[index] ?? (index > 0 ? items.at(-1) : undefined))?.focus();
  }

  // Fixed, so the ribbon panel's overflow does not clip it.
  function place(node: HTMLElement) {
    const bounds = anchor.getBoundingClientRect();
    node.style.left = `${Math.max(8, Math.min(bounds.left, innerWidth - node.offsetWidth - 8))}px`;
    node.style.top = `${bounds.bottom + 2}px`;
  }
  onMount(() => menu?.querySelector<HTMLButtonElement>('[role="menuitem"]')?.focus());
</script>

<svelte:window onpointerdown={(event) => { if (!menu?.contains(event.target as Node) && !anchor.contains(event.target as Node)) close(); }} />

<div class="wordart-gallery" role="menu" aria-label={label} tabindex="-1" bind:this={menu} use:place onkeydown={keys}>
  <div class="grid">
    {#each WORDART_PRESETS as preset (preset.label)}
      <button role="menuitem" aria-label={t(preset.label)} title={t(preset.label)} onclick={() => choose(preset)}>
        <span class="letter" aria-hidden="true" style={wordArtSwatchStyle(preset, theme)}>A</span>
      </button>
    {/each}
  </div>
  {#if clear}
    <hr />
    <button role="menuitem" class="clear" onclick={clear}>{t('Clear WordArt')}</button>
  {/if}
</div>

<style>
  .wordart-gallery { position: fixed; z-index: 450; padding: 6px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .grid { display: grid; grid-template-columns: repeat(5, 48px); gap: 4px; }
  button { font: inherit; color: var(--ok-text); border: 1px solid transparent; border-radius: 3px; cursor: pointer; }
  .grid button { display: flex; align-items: center; justify-content: center; width: 48px; height: 48px; padding: 0; overflow: hidden; background: #fff; }
  .grid button:hover, .grid button:focus-visible { outline: 2px solid var(--ok-accent); outline-offset: -2px; }
  .letter { font: 34px/1 Calibri, Carlito, Arial, sans-serif; color: transparent; }
  hr { border: none; border-top: 1px solid var(--ok-border); margin: 6px 0 4px; }
  .clear { width: 100%; padding: 5px 8px; text-align: left; font-size: 12px; white-space: nowrap; background: none; }
  .clear:hover { background: var(--ok-hover); }
</style>

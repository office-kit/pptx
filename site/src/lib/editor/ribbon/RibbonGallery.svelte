<script lang="ts" generics="Item extends { key: string; unavailable?: string }">
  // Mac PowerPoint's in-ribbon gallery: one row of tiles in a framed box, with
  // an 18 pt column at each end for the previous / next page arrows (the
  // previous column stays empty on the first page). The ribbon decides how
  // many tiles fit; the arrows page through the rest.
  import { untrack, type Snippet } from 'svelte';
  import { t } from '../i18n/i18n.svelte.ts';

  interface Props {
    label: string;
    items: readonly Item[];
    checked: string | null;
    /** Tiles shown at once. */
    visible: number;
    /** Tile pitch in CSS px (PowerPoint: 92 for transitions, 64 for effects). */
    tileWidth: number;
    disabled?: boolean;
    /** The tile's caption in the active locale. */
    name: (item: Item) => string;
    choose: (item: Item) => void;
    tile: Snippet<[Item]>;
  }
  let { label, items, checked, visible, tileWidth, disabled = false, name, choose, tile }: Props = $props();

  let first = $state(0);
  const count = $derived(Math.max(1, visible));
  // Show the page holding the applied item when it changes, as PowerPoint does
  // when a slide with a later effect is selected.
  $effect.pre(() => {
    const index = items.findIndex((item) => item.key === checked);
    untrack(() => {
      if (index >= 0 && (index < first || index >= first + count)) first = Math.floor(index / count) * count;
    });
  });
  const page = $derived(items.slice(first, first + count));
</script>

<div class="gallery" style:--tile="{tileWidth}px" style:--count={count}>
  <span class="end">
    {#if first > 0}
      <button class="arrow" aria-label={t('Previous {name} gallery').replace('{name}', label)} onclick={() => (first = Math.max(0, first - count))}>‹</button>
    {/if}
  </span>
  <div class="tiles" role="radiogroup" aria-label={label}>
    {#each page as item (item.key)}
      <button
        class="tile"
        role="radio"
        aria-checked={item.key === checked}
        aria-label={name(item)}
        disabled={disabled || item.unavailable !== undefined}
        title={item.unavailable ? t(item.unavailable) : name(item)}
        onclick={() => choose(item)}
      >
        {@render tile(item)}
        <span class="name">{name(item)}</span>
      </button>
    {/each}
  </div>
  <span class="end">
    <button class="arrow" aria-label={t('Next {name} gallery').replace('{name}', label)} disabled={first + count >= items.length} onclick={() => (first += count)}>›</button>
  </span>
</div>

<style>
  .gallery { display: flex; align-self: flex-start; height: 60px; border-radius: 4px; background: var(--ok-panel); box-shadow: inset 0 0 0 1px var(--ok-border); }
  .end { display: flex; flex: none; width: 18px; }
  .tiles { display: flex; width: calc(var(--tile) * var(--count)); }
  button { font: inherit; color: var(--ok-text); background: none; border: 1px solid transparent; border-radius: 4px; cursor: pointer; }
  .arrow { width: 18px; height: 56px; margin-top: 2px; padding: 0; font-size: 16px; line-height: 1; }
  .arrow:hover:not(:disabled) { background: var(--ok-hover); }
  .arrow:disabled { opacity: 0.35; cursor: default; }
  .tile { display: flex; flex: none; flex-direction: column; align-items: center; justify-content: flex-end; gap: 1px; width: var(--tile); height: 56px; margin-top: 2px; padding: 2px 0 3px; font-size: 10px; line-height: 1.1; letter-spacing: -0.1px; }
  .tile:hover:not(:disabled) { background: var(--ok-hover); }
  .tile[aria-checked='true'] { background: var(--ok-selected); border-color: var(--ok-selected-border); }
  .tile:disabled { cursor: default; }
  .tile:disabled > :global(*) { opacity: 0.4; }
  .name { max-width: 100%; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
</style>

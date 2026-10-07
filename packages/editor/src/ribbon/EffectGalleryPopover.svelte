<script lang="ts">
  // Mac PowerPoint's Exit Effects popover — and the Emphasis Effects one the
  // ribbon collapses to in a narrow window: a 387 pt panel of 76 × 90 pt tiles
  // on a 74 pt pitch, five to a row, under the gallery's group headings
  // (Basic, Subtle, Moderate, Exciting), scrolling past 508 pt.
  import type { Snippet } from 'svelte';
  import { getLocale } from '../i18n/i18n.svelte.ts';
  import { placeBelowTrigger } from './place-menu.ts';
  import { groupedTiles, type EffectTile } from './animation-gallery.ts';

  let { label, tiles, checked, choose, tile }: {
    label: string;
    tiles: readonly EffectTile[];
    /** The key of the applied tile in this gallery, if it is one of these. */
    checked: string | null;
    choose: (item: EffectTile) => void;
    tile: Snippet<[EffectTile]>;
  } = $props();
  const name = (item: { en: string; ja: string }): string => (getLocale() === 'ja' ? item.ja : item.en);
  const sections = $derived(groupedTiles(tiles));
</script>

<div class="popover" role="dialog" aria-label={label} use:placeBelowTrigger lang={getLocale()}>
  {#each sections as section (section.heading.group)}
    <div class="section" role="radiogroup" aria-label={name(section.heading)}>
      <div class="heading" aria-hidden="true">{name(section.heading)}</div>
      <div class="tiles">
        {#each section.tiles as item (item.key)}
          <button class="tile" role="radio" aria-checked={item.key === checked} aria-label={name(item)} title={name(item)} onclick={() => choose(item)}>
            {@render tile(item)}
            <span class="name">{name(item)}</span>
          </button>
        {/each}
      </div>
    </div>
  {/each}
</div>

<style>
  .popover { position: fixed; z-index: 300; box-sizing: border-box; width: 387px; max-height: min(508px, 70vh); overflow-y: auto; padding: 4px 0; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .heading { position: sticky; top: -4px; z-index: 1; display: flex; align-items: center; height: 19px; padding: 0 4px; font-size: 13px; font-weight: 500; background: var(--ok-panel-2, var(--ok-panel)); color: var(--ok-text); }
  .tiles { display: grid; grid-template-columns: repeat(5, 74px); padding: 2px 0 4px; }
  .tile { display: flex; flex-direction: column; align-items: center; justify-content: flex-start; gap: 4px; width: 76px; height: 90px; margin: 0 -1px; padding: 8px 0 0; font: inherit; font-size: 11px; line-height: 1.15; color: var(--ok-text); background: none; border: 1px solid transparent; border-radius: 4px; cursor: pointer; }
  .tile:hover { background: var(--ok-hover); }
  .tile[aria-checked='true'] { background: var(--ok-selected); border-color: var(--ok-selected-border); }
  .tile :global(svg) { width: 44px; height: 44px; flex: none; }
  .name { max-width: 74px; overflow: hidden; text-align: center; overflow-wrap: anywhere; }
  .name:lang(ja) { font-size: 10px; }
</style>

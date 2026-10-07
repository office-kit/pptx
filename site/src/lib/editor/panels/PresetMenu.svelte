<script lang="ts" generics="T">
  import { tick, type Snippet } from 'svelte';
  import { getLocale, t } from '../i18n/i18n.svelte.ts';
  import type { Preset } from './effects-model.ts';

  // PowerPoint's gallery buttons — 39 × 26 pt (Presets, Depth ...) or
  // 59 × 60 pt (Top bevel, Material, Lighting, 3-D Rotation presets) — open a
  // popover of square tiles under headings. A gallery with an "off" choice
  // starts with its own heading ("No Shadow") over a single "None" tile.
  let { label, groups, selected, none, disabled = false, title, columns = 3, size = 'small', tileSize = 68, tile, choose }: {
    label: string;
    groups: readonly { readonly heading?: string; readonly ja?: string; readonly items: readonly Preset<T>[] }[];
    /** Label of the current preset, or undefined when none matches. */
    selected: string | undefined;
    /** The "No Shadow" heading and whether its None tile is the current choice. */
    none?: { label: string; choose: () => void; selected: boolean };
    disabled?: boolean;
    title?: string;
    columns?: number;
    size?: 'small' | 'large';
    /** Native tiles: 68 pt (shadow, reflection, glow), 52 pt (3-D), 44 pt (soft edges). */
    tileSize?: number;
    /** Picture tiles; without one, tiles show the preset's name. */
    tile?: Snippet<[T | null]>;
    choose: (value: T) => void;
  } = $props();
  let open = $state(false);
  let trigger = $state<HTMLButtonElement>();
  let menu = $state<HTMLDivElement>();
  const name = (preset: { label: string; ja?: string }) => (getLocale() === 'ja' && preset.ja ? preset.ja : t(preset.label));
  const current = $derived(groups.flatMap((group) => group.items).find((item) => item.label === selected));

  async function show() {
    if (open) { close(); return; }
    open = true;
    await tick();
    (menu?.querySelector<HTMLButtonElement>('[aria-checked="true"]') ?? menu?.querySelector<HTMLButtonElement>('button'))?.focus();
  }
  function close(restore = true) { open = false; if (restore) trigger?.focus(); }
  function place(node: HTMLElement) {
    const bounds = trigger!.getBoundingClientRect();
    const left = size === 'large' ? bounds.left : bounds.right - node.offsetWidth;
    node.style.left = `${Math.max(8, Math.min(left, innerWidth - node.offsetWidth - 8))}px`;
    node.style.top = `${Math.max(8, Math.min(bounds.bottom + 2, innerHeight - node.offsetHeight - 8))}px`;
  }
  function keys(event: KeyboardEvent) {
    event.stopPropagation();
    if (event.key === 'Escape') { event.preventDefault(); close(); return; }
    if (event.key === 'Tab') { close(false); return; }
    const offsets: Record<string, number> = { ArrowLeft: -1, ArrowRight: 1, ArrowUp: -columns, ArrowDown: columns };
    if (!(event.key in offsets) && event.key !== 'Home' && event.key !== 'End') return;
    event.preventDefault();
    const items = [...menu!.querySelectorAll<HTMLButtonElement>('button')];
    const index = items.indexOf(event.target as HTMLButtonElement);
    items[event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : (index + offsets[event.key]! + items.length) % items.length]?.focus();
  }
</script>

<svelte:window onpointerdown={(event) => { if (open && !menu?.contains(event.target as Node) && !trigger?.contains(event.target as Node)) close(false); }} onresize={() => { if (open) close(false); }} />

<button bind:this={trigger} class="ok-input preset-trigger {size}" aria-label={label} title={title ?? label} aria-haspopup="menu" aria-expanded={open} {disabled} onclick={show}>
  <span class="preview">{#if tile}{@render tile(current?.value ?? null)}{:else}<span class="glyph" aria-hidden="true"></span>{/if}</span><span class="arrow">▾</span>
</button>
{#if open}
  <div class="preset-gallery" role="menu" aria-label={label} tabindex="-1" bind:this={menu} use:place onkeydown={keys} style:--tile="{tileSize}px">
    {#if none}
      <div class="heading">{name(none)}</div>
      <div class="tiles" style:grid-template-columns="repeat({columns}, var(--tile))">
        <button role="menuitemradio" aria-label={t('None')} title={t('None')} aria-checked={none.selected} onclick={() => { none.choose(); close(); }}>{#if tile}{@render tile(null)}{:else}<span class="glyph" aria-hidden="true"></span>{/if}</button>
      </div>
    {/if}
    {#each groups as group}
      {#if group.heading}<div class="heading">{name({ label: group.heading, ja: group.ja })}</div>{/if}
      <div class="tiles" style:grid-template-columns="repeat({columns}, var(--tile))">
        {#each group.items as item (item.label)}
          <button role="menuitemradio" aria-label={name(item)} title={name(item)} aria-checked={item.label === selected} onclick={() => { choose(item.value); close(); }}>{#if tile}{@render tile(item.value)}{:else}<span class="name">{name(item)}</span>{/if}</button>
        {/each}
      </div>
    {/each}
  </div>
{/if}

<style>
  .preset-trigger { display: flex; align-items: center; justify-content: space-between; box-sizing: border-box; width: 39px; height: 26px; padding: 0 2px 0 4px; font: inherit; flex: none; }
  .preset-trigger.large { width: 59px; height: 60px; padding: 0 3px 0 6px; }
  .preview { display: grid; place-items: center; width: 22px; height: 18px; overflow: hidden; }
  .large .preview { width: 40px; height: 40px; }
  .small .preview :global(.tile) { transform: scale(0.5); }
  .glyph { width: 14px; height: 14px; border: 1px solid var(--ok-text-2); border-radius: 2px; background: linear-gradient(135deg, var(--ok-panel), var(--ok-border-strong)); }
  .large .glyph { width: 26px; height: 26px; }
  .arrow { font-size: 9px; color: var(--ok-text-2); }
  .preset-gallery { position: fixed; z-index: 400; display: flex; flex-direction: column; max-height: 80vh; overflow: auto; padding: 4px 0; background: var(--ok-panel); border: 1px solid var(--ok-border); border-radius: 6px; box-shadow: var(--ok-shadow-lg); font-size: 12px; }
  /* Native popovers: a 17 pt heading 4 pt in, then tiles on their own size as pitch, 2 pt overlapping. */
  .heading { padding: 4px 4px 2px; font-weight: 600; color: var(--ok-text-2); }
  .tiles { display: grid; gap: 0; }
  .preset-gallery button { display: grid; place-items: center; width: var(--tile); height: var(--tile); margin: 0 -1px; padding: 2px; border: 1px solid transparent; border-radius: 3px; background: transparent; color: var(--ok-text); font: inherit; font-size: 10px; line-height: 1.1; cursor: pointer; overflow: hidden; }
  .name { text-align: center; }
  .preset-gallery button:hover, .preset-gallery button:focus-visible, .preset-gallery button[aria-checked='true'] { background: var(--ok-hover); border-color: var(--ok-accent); }
</style>

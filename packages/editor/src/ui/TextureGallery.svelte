<script module lang="ts">
  import { TEXTURE_SIZE, texturePixels, type TextureId } from '../core/textures.ts';

  const swatches = new Map<TextureId, ImageData>();
  // Draws the full tile, scaled down by CSS, so a swatch shows the texture itself.
  function draw(canvas: HTMLCanvasElement, id: TextureId) {
    let image = swatches.get(id);
    if (!image) { image = new ImageData(texturePixels(id), TEXTURE_SIZE, TEXTURE_SIZE); swatches.set(id, image); }
    canvas.getContext('2d')?.putImageData(image, 0, 0);
  }
</script>

<script lang="ts">
  import { eventTarget } from '../core/dom-root.ts';
  // PowerPoint's Texture gallery, shared by the Format pane's Texture ▾ and Shape
  // Fill ▸ Texture: twenty-four swatches, five to a row, sized and spaced like Mac
  // PowerPoint's (47 pt pitch, 36 pt tiles). Like PowerPoint, it marks no swatch
  // as the current fill. Only the ribbon's submenu adds More Textures...; the
  // Format pane's gallery has none because Insert... sits beside it.
  import { onMount } from 'svelte';
  import { TEXTURES } from '../core/textures.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let { label, anchor, side = 'below', choose, more, close }: {
    label: string;
    /** The control the gallery opens from. */
    anchor: HTMLElement;
    /** Below a button, right edges aligned as in PowerPoint, or beside a menu item as a submenu. */
    side?: 'below' | 'right';
    choose: (id: TextureId) => void;
    more?: () => void;
    close: () => void;
  } = $props();

  const COLUMNS = 5;
  let menu = $state<HTMLDivElement>();

  function keys(event: KeyboardEvent) {
    event.stopPropagation();
    if (event.key === 'Escape' || (side === 'right' && event.key === 'ArrowLeft' && (event.target as HTMLElement).matches('.grid button:nth-child(5n+1)'))) {
      event.preventDefault(); close(); anchor.focus(); return;
    }
    if (event.key === 'Tab') { close(); return; }
    const offsets: Record<string, number> = { ArrowLeft: -1, ArrowRight: 1, ArrowUp: -COLUMNS, ArrowDown: COLUMNS };
    if (!(event.key in offsets) && event.key !== 'Home' && event.key !== 'End') return;
    event.preventDefault();
    const items = [...menu!.querySelectorAll<HTMLButtonElement>('button')];
    if (event.key === 'Home' || event.key === 'End') { (event.key === 'Home' ? items[0] : items.at(-1))?.focus(); return; }
    // Down from the last row reaches More Textures... below the grid.
    const index = items.indexOf(event.target as HTMLButtonElement) + offsets[event.key]!;
    (items[index] ?? (index > 0 ? items.at(-1) : undefined))?.focus();
  }

  // Fixed, so the pane's or ribbon's overflow does not clip it.
  function place(node: HTMLElement) {
    const bounds = anchor.getBoundingClientRect();
    // A submenu that does not fit on the right opens to the left, clear of its item.
    const left = side === 'below' ? bounds.right - node.offsetWidth : bounds.right + 2 + node.offsetWidth <= innerWidth - 8 ? bounds.right + 2 : bounds.left - 2 - node.offsetWidth;
    const top = side === 'right' ? bounds.top : bounds.bottom + 2;
    node.style.left = `${Math.max(8, Math.min(left, innerWidth - node.offsetWidth - 8))}px`;
    node.style.top = `${Math.max(8, Math.min(top, innerHeight - node.offsetHeight - 8))}px`;
  }
  onMount(() => menu?.querySelector<HTMLButtonElement>('button')?.focus());
</script>

<svelte:window onpointerdown={(event) => { if (!menu?.contains(eventTarget(event) as Node) && !anchor.contains(eventTarget(event) as Node)) close(); }} />

<div class="texture-gallery" role="menu" aria-label={label} tabindex="-1" bind:this={menu} use:place onkeydown={keys}>
  <div class="grid">
    {#each TEXTURES as texture (texture.id)}
      <button type="button" role="menuitem" aria-label={t(texture.name)} title={t(texture.name)} onclick={() => choose(texture.id)}>
        <canvas width={TEXTURE_SIZE} height={TEXTURE_SIZE} aria-hidden="true" use:draw={texture.id}></canvas>
      </button>
    {/each}
  </div>
  {#if more}
    <hr />
    <button type="button" role="menuitem" class="more" onclick={more}>{t('More Textures...')}</button>
  {/if}
</div>

<style>
  .texture-gallery { position: fixed; z-index: 450; padding: 8px 5px 10px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .grid { display: grid; grid-template-columns: repeat(5, 36px); gap: 11px; }
  button { font: inherit; color: var(--ok-text); border: 1px solid transparent; border-radius: 3px; cursor: pointer; }
  .grid button { width: 36px; height: 36px; padding: 0; border: none; border-radius: 0; background: none; }
  .grid canvas { display: block; width: 100%; height: 100%; border: 1px solid var(--ok-border); }
  .grid button:hover, .grid button:focus-visible { outline: 2px solid var(--ok-accent); outline-offset: 1px; }
  hr { border: none; border-top: 1px solid var(--ok-border); margin: 6px 0 4px; }
  .more { width: 100%; padding: 5px 8px; text-align: left; font-size: 12px; white-space: nowrap; background: none; }
  .more:hover, .more:focus-visible { background: var(--ok-hover); }
</style>

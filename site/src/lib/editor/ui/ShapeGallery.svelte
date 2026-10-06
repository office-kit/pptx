<script lang="ts">
  // PowerPoint's Shapes gallery. Choosing a shape arms drawing: drag on the
  // slide to place it, or click for a one-inch shape.
  import { getEditor } from '../core/context.ts';
  import { SHAPE_GALLERY, shapeSprite, type GalleryShape } from '../core/shape-gallery.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const anchor = $derived(editor.shapeGallery);
  const sprite = $derived(anchor ? shapeSprite() : null);
  const ICON_PX = 24;

  const NAMES: Partial<Record<GalleryShape, string>> = {
    line: 'Line', rect: 'Rectangle', roundRect: 'Rounded Rectangle', ellipse: 'Oval', triangle: 'Isosceles Triangle', rtTriangle: 'Right Triangle',
    star5: '5-Point Star', rightArrow: 'Right Arrow', leftArrow: 'Left Arrow', upArrow: 'Up Arrow', downArrow: 'Down Arrow', hexagon: 'Hexagon', diamond: 'Diamond',
  };
  // Other presets read as their schema names spelled out ("leftRightArrow" → "Left Right Arrow").
  const label = (preset: GalleryShape) => NAMES[preset] ?? preset.replace(/([a-z])([A-Z0-9])/g, '$1 $2').replace(/^./, (c) => c.toUpperCase());

  function choose(preset: GalleryShape) {
    editor.shapeGallery = null;
    editor.drawShape = preset;
  }
</script>

<svelte:window
  onkeydown={(event) => { if (event.key === 'Escape') { editor.shapeGallery = null; editor.drawShape = null; } }}
  onpointerdown={(event) => { if (anchor && !(event.target as Element).closest?.('.shape-gallery, [aria-label="Shapes"]')) editor.shapeGallery = null; }}
/>

{#if anchor && sprite}
  <div class="shape-gallery ok-scroll" role="menu" aria-label={t('Shapes')} style:left="{Math.max(8, Math.min(anchor.left, innerWidth - 340))}px" style:top="{anchor.bottom + 2}px">
    {#each SHAPE_GALLERY as group (group.title)}
      <div class="heading">{t(group.title)}</div>
      <div class="grid">
        {#each group.shapes as preset (preset)}
          {@const cell = sprite.cells.get(preset) ?? 0}
          <button
            role="menuitem"
            aria-label={t(label(preset))}
            title={t(label(preset))}
            style:background-image="url('{sprite.url}')"
            style:background-size="{sprite.columns * ICON_PX}px {sprite.rows * ICON_PX}px"
            style:background-position="-{(cell % sprite.columns) * ICON_PX}px -{Math.floor(cell / sprite.columns) * ICON_PX}px"
            onclick={() => choose(preset)}
          ></button>
        {/each}
      </div>
    {/each}
  </div>
{/if}

<style>
  .shape-gallery { position: fixed; z-index: 450; width: 330px; max-height: 70vh; overflow-y: auto; padding: 6px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .heading { padding: 4px 2px; font-size: 11px; font-weight: 600; color: var(--ok-text-2); }
  .grid { display: grid; grid-template-columns: repeat(10, 28px); gap: 3px; margin-bottom: 4px; }
  button { width: 28px; height: 28px; padding: 0; border: 1px solid transparent; border-radius: 3px; background-color: #fff; background-repeat: no-repeat; background-origin: content-box; box-sizing: border-box; padding: 1px; cursor: pointer; }
  button:hover, button:focus-visible { border-color: var(--ok-selected-border); }
</style>

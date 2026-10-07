<script lang="ts">
  // The colour section of a colour emphasis effect's Effect Options: the
  // theme colours with their tints and shades, then the standard colours, as
  // the colour pickers elsewhere in the editor offer them.
  import { asColor, getPresentationTheme, type AnimationColor } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { paletteSwatches, sameColorTransforms, type PaletteSwatch } from '../core/theme-palette.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let { value, choose }: { value: AnimationColor | null; choose: (color: AnimationColor) => void } = $props();
  const editor = getEditor();
  const theme = $derived.by(() => { editor.doc.version; return getPresentationTheme(editor.doc.pres); });
  const swatches = $derived(paletteSwatches(theme, true));
  const current = $derived(
    value === null
      ? null
      : typeof value === 'string'
        ? { color: value, transforms: undefined }
        : { color: value.color, transforms: value.colorTransforms },
  );
  const checked = (swatch: PaletteSwatch): boolean =>
    current !== null &&
    current.color.replace(/^scheme:/, '').toLowerCase() === swatch.color.toLowerCase() &&
    sameColorTransforms(current.transforms, swatch.transforms);
  const nameOf = (swatch: PaletteSwatch): string =>
    `${t(swatch.name)}${swatch.shade ? `, ${t(swatch.shade)} ${swatch.shadePercent}%` : ''}`;
  function pick(swatch: PaletteSwatch): void {
    const color = asColor(swatch.theme ? `scheme:${swatch.color}` : swatch.color);
    if (color === null) return;
    choose(swatch.transforms === undefined ? color : { color, colorTransforms: swatch.transforms });
  }
</script>

{#each [true, false] as isTheme (isTheme)}
  {@const list = swatches.filter((swatch) => swatch.theme === isTheme)}
  {#if list.length > 0}
    {@const heading = t(isTheme ? 'Theme Colors' : 'Standard Colors')}
    <div class="heading" role="presentation">{heading}</div>
    <div class="colors" role="group" aria-label={heading}>
      {#each list as swatch, at (at)}
        <button type="button" role="menuitemradio" aria-label={nameOf(swatch)} title={nameOf(swatch)} aria-checked={checked(swatch)} style:background={swatch.paint} onclick={() => pick(swatch)}></button>
      {/each}
    </div>
  {/if}
{/each}

<style>
  /* Mac PowerPoint's 16 pt swatches on an 18 pt pitch, ten to a row. */
  .heading { display: flex; align-items: center; height: 18px; padding: 0 9px; font-size: 12px; color: var(--ok-text-2); }
  .colors { display: grid; grid-template-columns: repeat(10, 16px); gap: 2px; padding: 4px 10px 8px; }
  .colors button { width: 16px; height: 16px; padding: 0; border: 1px solid var(--ok-border); border-radius: 0; cursor: pointer; }
  .colors button:hover, .colors button:focus-visible, .colors button[aria-checked='true'] { outline: 2px solid var(--ok-accent); outline-offset: 1px; }
</style>

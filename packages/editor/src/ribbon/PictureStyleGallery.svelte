<script lang="ts">
  // Mac PowerPoint's in-ribbon Picture Styles gallery: three of the 28
  // built-in styles at a time between the previous / next page arrows, each
  // tile the style applied to a sample picture, named by its tooltip.
  import { BUILTIN_PICTURE_STYLES, getShapeKind, getShapeMedia, getShapePictureStyle, setShapePictureStyle, type BuiltinPictureStyleName } from '@office-kit/pptx';
  import { untrack } from 'svelte';
  import { getEditor } from '../core/context.ts';
  import { pictureStyleSwatches } from '../core/picture-style-swatches.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const PAGE = 3;
  const pictures = $derived.by(() => {
    doc.version;
    const shapes = editor.selectedShapes();
    return shapes.length > 0 && shapes.every((shape) => getShapeKind(shape) === 'picture' && !getShapeMedia(shape)) ? shapes : [];
  });
  const disabled = $derived(!pictures.length || editor.selectionLocked());
  // The style every selected picture carries exactly, if they share one.
  const current = $derived.by(() => {
    const styles = new Set(pictures.map(getShapePictureStyle));
    return styles.size === 1 ? [...styles][0]! : null;
  });
  let images = $state<ReadonlyMap<BuiltinPictureStyleName, string>>(new Map());
  let error = $state('');
  let first = $state(0);
  $effect(() => {
    if (disabled || images.size) return;
    try {
      images = pictureStyleSwatches();
    } catch (cause) {
      error = cause instanceof Error ? cause.message : String(cause);
    }
  });
  // Page to the applied style when the selection changes, as PowerPoint does.
  $effect.pre(() => {
    const index = current ? BUILTIN_PICTURE_STYLES.indexOf(current) : -1;
    untrack(() => {
      if (index >= 0 && (index < first || index >= first + PAGE)) first = Math.floor(index / PAGE) * PAGE;
    });
  });
  const visible = $derived(BUILTIN_PICTURE_STYLES.slice(first, first + PAGE));

  function apply(style: BuiltinPictureStyleName) {
    if (disabled) return;
    const targets = pictures;
    doc.transact(t('Picture Styles'), () => {
      for (const shape of targets) setShapePictureStyle(shape, style);
    });
  }
</script>

<div class="ctx-gallery" role="group" aria-label={t('Picture Styles')}>
  <button class="ctx-gallery-arrow" class:hidden={first === 0} {disabled} aria-label={t('Previous Quick Styles gallery')} onclick={() => (first = Math.max(0, first - PAGE))}>‹</button>
  <div class="ctx-gallery-items picture-strip" role="radiogroup" aria-label={t('Quick Styles')}>
    {#each visible as style (style)}
      <button class="picture-style" role="radio" aria-checked={style === current} aria-label={t(style)} title={t(style)} {disabled} onclick={() => apply(style)}>
        {#if images.has(style)}<img src={images.get(style)} alt="" />{/if}
      </button>
    {/each}
  </div>
  <button class="ctx-gallery-arrow" aria-label={t('Next Quick Styles gallery')} disabled={disabled || first + PAGE >= BUILTIN_PICTURE_STYLES.length} onclick={() => (first += PAGE)}>›</button>
  {#if error}<p class="error" role="alert">{error}</p>{/if}
</div>

<style>
  .ctx-gallery-items.picture-strip { width: 212px; justify-content: space-around; }
  .picture-style { display: flex; align-items: center; justify-content: center; width: 62px; height: 54px; padding: 0; border: 1px solid transparent; border-radius: 3px; background: none; cursor: pointer; }
  .picture-style:hover:not(:disabled), .picture-style:focus-visible { outline: 2px solid var(--ok-accent); outline-offset: -2px; }
  .picture-style[aria-checked='true'] { background: var(--ok-selected); border-color: var(--ok-selected-border); }
  .picture-style:disabled { cursor: default; opacity: 0.4; }
  .picture-style img { width: 100%; height: 100%; display: block; object-fit: contain; }
  .error { max-width: 240px; color: var(--ok-danger); }
</style>

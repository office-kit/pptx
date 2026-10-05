<script lang="ts">
  import { getSlides, isSlideHidden, setSlideHidden } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { selectedSlideIndices } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const slides = $derived.by(() => {
    doc.version;
    const all = getSlides(doc.pres);
    return selectedSlideIndices(doc.selection).flatMap((index) => (all[index] ? [all[index]!] : []));
  });
  const hidden = $derived(slides.length > 0 && slides.every(isSlideHidden));

  function toggleHidden() {
    const value = !hidden;
    doc.transact(t('Hide Slide'), () => {
      for (const slide of slides) setSlideHidden(slide, value);
    });
  }
</script>

<div class="group">
  <button disabled={!editor.canPresent} onclick={() => editor.present('start')}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="18" height="13" rx="1" /><path d="M12 16v5M8 21h8M10 7l5 2.5-5 2.5z" /></svg>
    {t('Play from Start')}
  </button>
  <button disabled={!editor.canPresent} onclick={() => editor.present('current')}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="18" height="13" rx="1" /><path d="M12 16v5M8 21h8M7 7h5v5H7zM14 8h3M14 11h3" /></svg>
    {t('Play from Current Slide')}
  </button>
  <button disabled={!editor.canPresent} onclick={() => editor.present('presenter')}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="18" height="13" rx="1" /><path d="M12 16v5M8 21h8" /></svg>
    {t('Presenter View')}
  </button>
</div>
<div class="group">
  <button onclick={() => editor.openCustomShows()}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="13" height="10" rx="1" /><rect x="7" y="7" width="13" height="10" rx="1" /><path d="M13 17v4M9 21h8" /></svg>
    {t('Custom Show')}
  </button>
</div>
<div class="group">
  <button onclick={() => (editor.activeDialog = 'showProperties')}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="9" y="3" width="12" height="10" rx="1" /><path d="M3 4h4M3 8h4M3 12h4M15 13v5M11 21h8M15 18v3" /></svg>
    {t('Set Up Slide Show')}
  </button>
  <button aria-pressed={hidden} disabled={slides.length === 0} onclick={toggleHidden}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="18" height="15" rx="1" /><path d="M3 19 21 4" /></svg>
    {t('Hide Slide')}
  </button>
</div>

<style>
  .group { display: flex; align-items: center; flex-shrink: 0; gap: 2px; padding: 0 8px; border-right: 1px solid var(--ok-border); }
  button { display: flex; flex-direction: column; align-items: center; justify-content: flex-start; gap: 2px; min-width: 52px; max-width: 96px; min-height: 66px; padding: 3px 4px; font: inherit; font-size: 11px; line-height: 1.2; color: var(--ok-text); background: transparent; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  button:not(:disabled):hover, button[aria-pressed='true'] { background: var(--ok-hover); border-color: var(--ok-border); }
  button:disabled { opacity: 0.45; cursor: default; }
  svg { width: 32px; height: 32px; flex-shrink: 0; stroke: currentColor; fill: none; stroke-width: 1.1; }
</style>

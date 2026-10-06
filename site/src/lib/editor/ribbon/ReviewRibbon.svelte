<script lang="ts">
  import {
    clearAllSlideComments,
    clearSlideComments,
    getShapeKind,
    getShapeTextLanguage,
    getSlideComments,
    getSlides,
    getSlideShapes,
    isShapeHidden,
    setShapeHidden,
    setShapeTextLanguage,
  } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { isInk } from '../core/ink.svelte.ts';
  import { selectedShapeIds } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const COMMENTS = 'addSlideComment';
  // Proofing languages offered first by Mac PowerPoint's Language dialog.
  const LANGUAGES = ['en-US', 'en-GB', 'ja-JP', 'zh-CN', 'zh-TW', 'ko-KR', 'fr-FR', 'de-DE', 'es-ES', 'it-IT', 'pt-BR'];
  const languageName = (tag: string) => new Intl.DisplayNames([navigator.language], { type: 'language' }).of(tag) ?? tag;

  const slides = $derived.by(() => { doc.version; return getSlides(doc.pres); });
  const withComments = $derived.by(() => { doc.version; return slides.flatMap((slide, index) => (getSlideComments(slide).length ? [index] : [])); });
  const textShapes = $derived.by(() => {
    doc.version;
    return selectedShapeIds(doc.selection).flatMap((id) => {
      const shape = doc.shapeById(doc.selection.slideIndex, id);
      return shape && getShapeKind(shape) === 'shape' ? [shape] : [];
    });
  });
  const inks = $derived.by(() => { doc.version; return slides.flatMap((slide) => getSlideShapes(slide).filter(isInk)); });
  const inkHidden = $derived(inks.length > 0 && inks.every(isShapeHidden));
  let menu = $state<'delete' | 'language' | null>(null);

  function goToComment(direction: 1 | -1) {
    const here = doc.selection.slideIndex;
    const target = direction === 1 ? withComments.find((index) => index > here) : withComments.findLast((index) => index < here);
    if (target === undefined) return;
    doc.selectSlide(target);
    editor.activeDialog = COMMENTS;
  }
  function deleteComments(all: boolean) {
    menu = null;
    doc.transact(t('Delete'), () => {
      if (all) clearAllSlideComments(doc.pres);
      else if (doc.currentSlide) clearSlideComments(doc.currentSlide);
    });
  }
  function setLanguage(tag: string) {
    menu = null;
    doc.transact(t('Language'), () => { for (const shape of textShapes) setShapeTextLanguage(shape, tag); });
  }
  function toggleInk() {
    const hide = !inkHidden;
    doc.transact(t('Hide Ink'), () => setShapeHidden(inks, hide));
  }
</script>

<svelte:window onpointerdown={(event) => { if (menu && !(event.target as Element).closest?.('.menu-anchor')) menu = null; }} onkeydown={(event) => { if (event.key === 'Escape') menu = null; }} />

<div class="group">
  <button disabled title={t('Spelling is checked by the browser as you type.')}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 11 6 3l3 8M4 8h4M11 3v8h3a2 2 0 0 0 0-4h-3M21 4a3 3 0 1 0 0 6" /><path d="m6 17 4 4 9-9" stroke="#2e8b57" /></svg>
    {t('Spelling')}
  </button>
  <button disabled title={t('Select a word to look it up.')}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 5h7a2 2 0 0 1 2 2v13a2 2 0 0 0-2-2H3zM21 5h-7a2 2 0 0 0-2 2v13a2 2 0 0 1 2-2h7z" /></svg>
    {t('Thesaurus')}
  </button>
</div>
<div class="group">
  <button onclick={() => (editor.accessibilityOpen = true)}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 3h9l4 4v14H6z" /><circle cx="15" cy="12" r="1.2" stroke="#0078d4" /><path d="M12 14h6M15 14v3l-2 3M15 17l2 3" stroke="#0078d4" /></svg>
    {t('Check Accessibility')}
  </button>
</div>
<div class="group">
  <button disabled title={t('Translation needs an online service.')}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 17 6 9l3 8M4 14h4" stroke="#0078d4" /><path d="M13 4h8M17 3v2M14 11c3-1 5-3 6-6M15 7c1 2 3 4 6 4" stroke="#2e8b57" /></svg>
    {t('Translate')}
  </button>
  <div class="menu-anchor">
    <button aria-haspopup="menu" aria-expanded={menu === 'language'} disabled={textShapes.length === 0} onclick={() => (menu = menu === 'language' ? null : 'language')}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 20 9 5l5 15M6 15h6" /><path d="M16 4v6M14 6h6" stroke="#0078d4" /></svg>
      {t('Language')}
    </button>
    {#if menu === 'language'}
      <div class="menu" role="menu" aria-label={t('Language')}>
        {#each LANGUAGES as tag (tag)}
          <button role="menuitemradio" aria-checked={textShapes.length > 0 && getShapeTextLanguage(textShapes[0]!) === tag} onclick={() => setLanguage(tag)}>{languageName(tag)}</button>
        {/each}
      </div>
    {/if}
  </div>
</div>
<div class="group">
  <button disabled title={t('Co-authoring is not available here.')}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 3h9l4 4v14H6z" /><path d="m9 13 2 2 4-4" /></svg>
    {t('Mark All as Read')}
  </button>
  <button disabled title={t('Co-authoring is not available here.')}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="18" height="15" rx="1" /><rect x="11" y="8" width="7" height="5" /></svg>
    {t('Show Changes')}
  </button>
</div>
<div class="group">
  <button disabled={!doc.currentSlide} onclick={() => editor.runOrPrompt(COMMENTS)}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 5h16v11H9l-4 4v-4H3z" /><path d="M18 2v6M15 5h6" stroke="#2e8b57" /></svg>
    {t('New Comment')}
  </button>
  <div class="menu-anchor">
    <button aria-haspopup="menu" aria-expanded={menu === 'delete'} disabled={withComments.length === 0} onclick={() => (menu = menu === 'delete' ? null : 'delete')}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 7h16v11H9l-4 4v-4H3z" /><path d="m3 2 6 6M9 2 3 8" stroke="#d13438" /></svg>
      {t('Delete')}
    </button>
    {#if menu === 'delete'}
      <div class="menu" role="menu" aria-label={t('Delete')}>
        <button role="menuitem" disabled={!withComments.includes(doc.selection.slideIndex)} onclick={() => deleteComments(false)}>{t('Delete All Comments on This Slide')}</button>
        <button role="menuitem" onclick={() => deleteComments(true)}>{t('Delete All Comments in This Presentation')}</button>
      </div>
    {/if}
  </div>
  <button disabled={!withComments.some((index) => index < doc.selection.slideIndex)} onclick={() => goToComment(-1)}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 5h15v11h-9l-4 4v-4H6z" /><path d="M4 10H1m2-2-2 2 2 2" /></svg>
    {t('Previous')}
  </button>
  <button disabled={!withComments.some((index) => index > doc.selection.slideIndex)} onclick={() => goToComment(1)}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 5h15v11H9l-4 4v-4H3z" /><path d="M20 10h3m-2-2 2 2-2 2" /></svg>
    {t('Next')}
  </button>
</div>
<div class="group">
  <button aria-pressed={editor.activeDialog === COMMENTS} disabled={!doc.currentSlide} onclick={() => { if (editor.activeDialog === COMMENTS) editor.activeDialog = null; else editor.runOrPrompt(COMMENTS); }}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 5h18v12H10l-5 4v-4H3z" /></svg>
    {t('Show Comments')}
  </button>
</div>
<div class="group">
  <button disabled title={t('Document protection is not available here.')}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="m4 20 3-1 11-11-2-2L5 17z" /><circle cx="17" cy="17" r="4" stroke="#d13438" /><path d="m14 20 6-6" stroke="#d13438" /></svg>
    {t('Always Open Read-Only')}
  </button>
  <button disabled title={t('Document protection is not available here.')}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 3h9l4 4v14H5z" /><circle cx="15" cy="16" r="4" stroke="#d13438" /><path d="M13 16h4" stroke="#d13438" /></svg>
    {t('Restrict Permission')}
  </button>
</div>
<div class="group">
  <button aria-pressed={inkHidden} disabled={inks.length === 0} onclick={toggleInk}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 20c4-6 6-9 9-10s4 3 7 0" stroke="#d13438" /><path d="m3 3 18 18" /></svg>
    {t('Hide Ink')}
  </button>
</div>

<style>
  .group { display: flex; align-items: center; flex-shrink: 0; gap: 2px; padding: 0 8px; border-right: 1px solid var(--ok-border); }
  .group > button, .menu-anchor > button { display: flex; flex-direction: column; align-items: center; justify-content: flex-start; gap: 2px; min-width: 52px; max-width: 96px; min-height: 66px; padding: 3px 4px; font: inherit; font-size: 11px; line-height: 1.2; color: var(--ok-text); text-align: center; background: transparent; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  .group > button:not(:disabled):hover, .menu-anchor > button:not(:disabled):hover, button[aria-pressed='true'], button[aria-expanded='true'] { background: var(--ok-hover); border-color: var(--ok-border); }
  button:disabled { opacity: 0.45; cursor: default; }
  svg { width: 32px; height: 32px; flex-shrink: 0; stroke: currentColor; fill: none; stroke-width: 1.1; }
  .menu-anchor { position: relative; }
  .menu { position: absolute; top: 100%; left: 0; z-index: 300; display: flex; flex-direction: column; min-width: 240px; padding: 4px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .menu button { padding: 5px 10px; font: inherit; font-size: 12px; text-align: left; color: var(--ok-text); background: transparent; border: 0; border-radius: 4px; cursor: pointer; }
  .menu button:not(:disabled):hover, .menu button[aria-checked='true'] { background: var(--ok-hover); }
</style>

<script lang="ts">
  import { eventTarget } from '../core/dom-root.ts';
  import { placeBelowTrigger } from './place-menu.ts';
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
  import { getLocale, t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const COMMENTS = 'addSlideComment';
  // Proofing languages offered first by the reference desktop app's (Mac) Language dialog.
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

<svelte:window onpointerdown={(event) => { if (menu && !(eventTarget(event) as Element).closest?.('.menu-anchor')) menu = null; }} onkeydown={(event) => { if (event.key === 'Escape') menu = null; }} />

<!-- The reference desktop app's (Mac, 16) Review tab: Spelling, Thesaurus | Check Accessibility |
     Translate, Language | Mark All as Read, Show Changes | New Comment,
     Delete ▾, Previous, Next, a rule, Show Comments | Always Open Read-Only,
     Restrict Permission ▾ | Hide Ink. Unchanged from 1512 to 1200 pt. -->
<div class="review" lang={getLocale()}>
  <section class="cluster" role="group" aria-label={t('Proofing')}>
    <button class="big" style:--w="47px" disabled title={t('Spelling is checked by the browser as you type.')}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 11 6 3l3 8M4 8h4M11 3v8h3a2 2 0 0 0 0-4h-3M21 4a3 3 0 1 0 0 6" /><path d="m6 17 4 4 9-9" stroke="#2e8b57" /></svg>
      <span>{t('Spelling')}</span>
    </button>
    <button class="big" style:--w="59px" disabled title={t('Select a word to look it up.')}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 5h7a2 2 0 0 1 2 2v13a2 2 0 0 0-2-2H3zM21 5h-7a2 2 0 0 0-2 2v13a2 2 0 0 1 2-2h7z" /></svg>
      <span>{t('Thesaurus')}</span>
    </button>
  </section>
  <section class="cluster" role="group" aria-label={t('Accessibility')}>
    <button class="big" style:--w="70px" onclick={() => (editor.accessibilityOpen = true)}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 3h9l4 4v14H6z" /><circle cx="15" cy="12" r="1.2" style="stroke: var(--ok-accent)" /><path d="M12 14h6M15 14v3l-2 3M15 17l2 3" style="stroke: var(--ok-accent)" /></svg>
      <span>{t('Check Accessibility')}</span>
    </button>
  </section>
  <section class="cluster" role="group" aria-label={t('Language')}>
    <button class="big" style:--w="53px" disabled title={t('Translation needs an online service.')}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 17 6 9l3 8M4 14h4" style="stroke: var(--ok-accent)" /><path d="M13 4h8M17 3v2M14 11c3-1 5-3 6-6M15 7c1 2 3 4 6 4" stroke="#2e8b57" /></svg>
      <span>{t('Translate')}</span>
    </button>
    <div class="menu-anchor">
      <button class="big" style:--w="56px" aria-haspopup="menu" aria-expanded={menu === 'language'} disabled={textShapes.length === 0} onclick={() => (menu = menu === 'language' ? null : 'language')}>
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 20 9 5l5 15M6 15h6" /><path d="M16 4v6M14 6h6" style="stroke: var(--ok-accent)" /></svg>
        <span>{t('Language')}</span>
      </button>
      {#if menu === 'language'}
        <div class="menu" role="menu" use:placeBelowTrigger aria-label={t('Language')}>
          {#each LANGUAGES as tag (tag)}
            <button role="menuitemradio" aria-checked={textShapes.length > 0 && getShapeTextLanguage(textShapes[0]!) === tag} onclick={() => setLanguage(tag)}>{languageName(tag)}</button>
          {/each}
        </div>
      {/if}
    </div>
  </section>
  <section class="cluster" role="group" aria-label={t('Activity')}>
    <button class="big" style:--w="47px" disabled title={t('Co-authoring is not available here.')}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 3h9l4 4v14H6z" /><path d="m9 13 2 2 4-4" /></svg>
      <span>{t('Mark All as Read')}</span>
    </button>
    <button class="big" style:--w="51px" disabled title={t('Co-authoring is not available here.')}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="18" height="15" rx="1" /><rect x="11" y="8" width="7" height="5" /></svg>
      <span>{t('Show Changes')}</span>
    </button>
  </section>
  <section class="cluster" role="group" aria-label={t('Comments')}>
    <button class="big" style:--w="55px" disabled={!doc.currentSlide} onclick={() => editor.runOrPrompt(COMMENTS)}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 5h16v11H9l-4 4v-4H3z" /><path d="M18 2v6M15 5h6" stroke="#2e8b57" /></svg>
      <span>{t('New Comment')}</span>
    </button>
    <div class="menu-anchor">
      <button class="big" style:--w="50px" aria-haspopup="menu" aria-expanded={menu === 'delete'} disabled={withComments.length === 0} onclick={() => (menu = menu === 'delete' ? null : 'delete')}>
        <span class="icon-row"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 7h16v11H9l-4 4v-4H3z" /><path d="m3 2 6 6M9 2 3 8" stroke="#c92a2a" /></svg><span class="arrow" aria-hidden="true">⌄</span></span>
        <span>{t('Delete')}</span>
      </button>
      {#if menu === 'delete'}
        <div class="menu" role="menu" use:placeBelowTrigger aria-label={t('Delete')}>
          <button role="menuitem" disabled={!withComments.includes(doc.selection.slideIndex)} onclick={() => deleteComments(false)}>{t('Delete All Comments on This Slide')}</button>
          <button role="menuitem" onclick={() => deleteComments(true)}>{t('Delete All Comments in This Presentation')}</button>
        </div>
      {/if}
    </div>
    <button class="big" style:--w="50px" disabled={!withComments.some((index) => index < doc.selection.slideIndex)} onclick={() => goToComment(-1)}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 5h15v11h-9l-4 4v-4H6z" /><path d="M4 10H1m2-2-2 2 2 2" /></svg>
      <span>{t('Previous')}</span>
    </button>
    <button class="big" style:--w="38px" disabled={!withComments.some((index) => index > doc.selection.slideIndex)} onclick={() => goToComment(1)}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 5h15v11H9l-4 4v-4H3z" /><path d="M20 10h3m-2-2 2 2-2 2" /></svg>
      <span>{t('Next')}</span>
    </button>
    <span class="sep" aria-hidden="true"></span>
    <button class="big" style:--w="61px" aria-pressed={editor.activeDialog === COMMENTS} disabled={!doc.currentSlide} onclick={() => { if (editor.activeDialog === COMMENTS) editor.activeDialog = null; else editor.runOrPrompt(COMMENTS); }}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 5h18v12H10l-5 4v-4H3z" /></svg>
      <span>{t('Show Comments')}</span>
    </button>
  </section>
  <section class="cluster" role="group" aria-label={t('Protect')}>
    <button class="big" style:--w="71px" disabled title={t('Document protection is not available here.')}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="m4 20 3-1 11-11-2-2L5 17z" /><circle cx="17" cy="17" r="4" stroke="#c92a2a" /><path d="m14 20 6-6" stroke="#c92a2a" /></svg>
      <span>{t('Always Open Read-Only')}</span>
    </button>
    <button class="big" style:--w="62px" aria-haspopup="menu" disabled title={t('Document protection is not available here.')}>
      <span class="icon-row"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 3h9l4 4v14H5z" /><circle cx="15" cy="16" r="4" stroke="#c92a2a" /><path d="M13 16h4" stroke="#c92a2a" /></svg><span class="arrow" aria-hidden="true">⌄</span></span>
      <span>{t('Restrict Permission')}</span>
    </button>
  </section>
  <section class="cluster" role="group" aria-label={t('Ink')}>
    <button class="big" style:--w="50px" aria-pressed={inkHidden} disabled={inks.length === 0} onclick={toggleInk}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 20c4-6 6-9 9-10s4 3 7 0" stroke="#c92a2a" /><path d="m3 3 18 18" /></svg>
      <span>{t('Hide Ink')}</span>
    </button>
  </section>
</div>

<style>
  /* Geometry measured from the reference desktop app (Mac, 16) (NATIVE_PARITY.md, "Native
     geometry audit"). */
  .review { display: flex; align-items: stretch; width: 100%; min-width: 0; height: 72px; }
  .cluster { display: flex; flex: none; align-items: stretch; padding: 0 10px; border-right: 1px solid var(--ok-border); }
  .cluster:first-child { padding-left: 4px; }
  .cluster:last-child { border-right: none; }
  button { font: inherit; color: var(--ok-text); background: none; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  button:hover:not(:disabled) { background: var(--ok-hover); }
  button:disabled { opacity: 0.4; cursor: default; }
  button[aria-pressed='true'], button[aria-expanded='true'] { background: var(--ok-selected); border-color: var(--ok-selected-border); }
  .big { display: flex; flex-direction: column; align-items: center; justify-content: flex-start; gap: 2px; min-width: var(--w); padding: 4px 1px; font-size: 11px; line-height: 1.15; text-align: center; }
  .big > span:last-child { max-width: var(--w); margin: 0 -2px; }
  /* Japanese labels wrap per character, so they get at least six characters
     a line and a smaller size that fits three lines (the reference desktop app widens them). */
  .big > span:last-child:lang(ja) { max-width: max(calc(var(--w) - 4px), 6em); font-size: 10px; line-height: 1.1; }
  .icon-row { display: flex; align-items: center; gap: 1px; }
  .arrow { font-size: 11px; }
  svg { width: 32px; height: 32px; flex: none; fill: none; stroke: currentColor; stroke-width: 1.1; }
  .sep { align-self: center; width: 1px; height: 52px; margin: 0 7px; background: var(--ok-border); }
  .menu-anchor { position: relative; display: flex; }
  .menu { position: fixed; z-index: 300; display: flex; flex-direction: column; min-width: 240px; padding: 4px 0; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .menu button { height: 24px; padding: 0 16px; border: none; border-radius: 0; font-size: 13px; text-align: left; }
  .menu button[aria-checked='true'] { background: var(--ok-selected); }
</style>

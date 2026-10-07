<script lang="ts">
  import { eventTarget } from '../core/dom-root.ts';
  // Mac PowerPoint 16's Slide Show tab: Play from Start, Play from Current
  // Slide, Presenter View, Custom Show ▾ | Rehearse with Coach | Set Up Slide
  // Show, Hide Slide, Rehearse Timings, Record ▾ and four options | Always Use
  // Subtitles, Subtitle Settings ▾. It keeps this layout from 1512 down to
  // 1200 pt windows.
  import { getSlides, getSlideShowProperties, isSlideHidden, setSlideHidden, setSlideShowProperties, type SlideShowProperties } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { selectedSlideIndices } from '../core/selection.ts';
  import { getLocale, t } from '../i18n/i18n.svelte.ts';
  import { placeBelowTrigger } from './place-menu.ts';

  const SUBTITLES = 'Live subtitles need Microsoft 365 online services.';
  const editor = getEditor();
  const doc = editor.doc;
  const slides = $derived.by(() => {
    doc.version;
    const all = getSlides(doc.pres);
    return selectedSlideIndices(doc.selection).flatMap((index) => (all[index] ? [all[index]!] : []));
  });
  const hidden = $derived(slides.length > 0 && slides.every(isSlideHidden));

  const show = $derived.by(() => { doc.version; return getSlideShowProperties(doc.pres); });
  function setShow(label: string, change: Partial<SlideShowProperties>) {
    doc.transact(t(label), () => setSlideShowProperties(doc.pres, { ...show, ...change }));
  }

  function toggleHidden() {
    const value = !hidden;
    doc.transact(t('Hide Slide'), () => {
      for (const slide of slides) setSlideHidden(slide, value);
    });
  }
  let menu = $state<'custom' | 'record' | null>(null);
  const toggle = (name: 'custom' | 'record') => (menu = menu === name ? null : name);
</script>

<svelte:window onkeydown={(event) => { if (event.key === 'Escape') menu = null; }} onpointerdown={(event) => { if (menu && !(eventTarget(event) as Element).closest?.('.anchor')) menu = null; }} />

<div class="slide-show" lang={getLocale()}>
  <section class="cluster" role="group" aria-label={t('Start Slide Show')}>
    <button class="big" style:--w="54px" disabled={!editor.canPresent} onclick={() => editor.present('start')}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="18" height="13" rx="1" /><path d="M12 16v5M8 21h8M10 7l5 2.5-5 2.5z" /></svg>
      <span>{t('Play from Start')}</span>
    </button>
    <button class="big" style:--w="73px" disabled={!editor.canPresent} onclick={() => editor.present('current')}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="18" height="13" rx="1" /><path d="M12 16v5M8 21h8M7 7h5v5H7zM14 8h3M14 11h3" /></svg>
      <span>{t('Play from Current Slide')}</span>
    </button>
    <button class="big" style:--w="55px" disabled={!editor.canPresent} onclick={() => editor.present('presenter')}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="18" height="13" rx="1" /><path d="M12 16v5M8 21h8" /></svg>
      <span>{t('Presenter View')}</span>
    </button>
    <div class="anchor">
      <button class="big" style:--w="50px" aria-haspopup="menu" aria-expanded={menu === 'custom'} onclick={() => toggle('custom')}>
        <span class="icon-row"><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="13" height="10" rx="1" /><rect x="7" y="7" width="13" height="10" rx="1" /><path d="M13 17v4M9 21h8" /></svg><span class="arrow" aria-hidden="true">⌄</span></span>
        <span>{t('Custom Show')}</span>
      </button>
      {#if menu === 'custom'}
        <div class="menu" role="menu" aria-label={t('Custom Show')} use:placeBelowTrigger>
          <button role="menuitem" onclick={() => { menu = null; editor.openCustomShows(); }}>{t('Custom Slide Show...')}</button>
        </div>
      {/if}
    </div>
  </section>
  <section class="cluster" role="group" aria-label={t('Rehearse with Coach')}>
    <button class="big" style:--w="63px" disabled title={t('Rehearse with Coach needs Microsoft 365 online services.')}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="16" height="12" rx="1" /><path d="M11 16v4M7 20h8" /><path d="M19 9a2 2 0 0 1 2 2v3a2 2 0 0 1-4 0v-3a2 2 0 0 1 2-2z" style="stroke: var(--ok-accent)" /></svg>
      <span>{t('Rehearse with Coach')}</span>
    </button>
  </section>
  <section class="cluster" role="group" aria-label={t('Set Up')}>
    <button class="big" style:--w="62px" onclick={() => (editor.activeDialog = 'showProperties')}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="9" y="3" width="12" height="10" rx="1" /><path d="M3 4h4M3 8h4M3 12h4M15 13v5M11 21h8M15 18v3" /></svg>
      <span>{t('Set Up Slide Show')}</span>
    </button>
    <button class="big" style:--w="38px" aria-pressed={hidden} disabled={slides.length === 0} onclick={toggleHidden}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="18" height="15" rx="1" /><path d="M3 19 21 4" /></svg>
      <span>{t('Hide Slide')}</span>
    </button>
    <button class="big" style:--w="53px" disabled={!editor.canPresent} onclick={() => editor.present('rehearse')}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="16" height="12" rx="1" /><circle cx="17" cy="16" r="5" style="stroke: var(--ok-accent)" /><path d="M17 13v3l2 1" style="stroke: var(--ok-accent)" /></svg>
      <span>{t('Rehearse Timings')}</span>
    </button>
    <div class="anchor">
      <button class="big" style:--w="50px" aria-haspopup="menu" aria-expanded={menu === 'record'} disabled={!editor.canPresent} onclick={() => toggle('record')}>
        <span class="icon-row"><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="14" height="11" rx="1" /><path d="M10 14v5M6 19h8" /><circle cx="19" cy="14" r="3" fill="#c92a2a" stroke="none" /></svg><span class="arrow" aria-hidden="true">⌄</span></span>
        <span>{t('Record')}</span>
      </button>
      {#if menu === 'record'}
        <div class="menu" role="menu" aria-label={t('Record')} use:placeBelowTrigger>
          <button role="menuitem" onclick={() => { menu = null; editor.present('start'); }}>{t('From Beginning')}</button>
          <button role="menuitem" onclick={() => { menu = null; editor.present('current'); }}>{t('From Current Slide')}</button>
        </div>
      {/if}
    </div>
    <div class="checks">
      <label title={t('Keep Slides Updated needs a presentation shared from OneDrive or SharePoint.')}><input type="checkbox" disabled checked={false} />{t('Keep Slides Updated')}</label>
      <label><input type="checkbox" checked={show.showNarration} onchange={(event) => setShow('Play Narrations', { showNarration: event.currentTarget.checked })} />{t('Play Narrations')}</label>
      <label><input type="checkbox" checked={show.useTimings} onchange={(event) => setShow('Use Timings', { useTimings: event.currentTarget.checked })} />{t('Use Timings')}</label>
      <label><input type="checkbox" checked={show.showMediaControls ?? true} onchange={(event) => setShow('Show Media Controls', { showMediaControls: event.currentTarget.checked })} />{t('Show Media Controls')}</label>
    </div>
  </section>
  <section class="cluster" role="group" aria-label={t('Captions & Subtitles')}>
    <div class="checks single">
      <label title={t(SUBTITLES)}><input type="checkbox" disabled />{t('Always Use Subtitles')}</label>
      <button class="row" aria-haspopup="menu" disabled title={t(SUBTITLES)}>
        <svg viewBox="0 0 16 16" aria-hidden="true"><rect x="1.5" y="3" width="13" height="10" rx="1" /><path d="M4 9h3M9 9h3M4 11h8" /></svg>
        {t('Subtitle Settings')}<span class="arrow" aria-hidden="true">⌄</span>
      </button>
    </div>
  </section>
</div>

<style>
  /* Geometry measured from Mac PowerPoint 16 (POWERPOINT_PARITY.md, "Native
     geometry audit"). */
  .slide-show { display: flex; align-items: stretch; width: 100%; min-width: 0; height: 72px; }
  .cluster { display: flex; flex: none; align-items: stretch; padding: 0 10px; border-right: 1px solid var(--ok-border); }
  .cluster:first-child { padding-left: 4px; }
  .cluster:last-child { border-right: none; }
  button { font: inherit; color: var(--ok-text); background: none; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  button:hover:not(:disabled) { background: var(--ok-hover); }
  button:disabled, label:has(input:disabled) { opacity: 0.4; cursor: default; }
  button[aria-pressed='true'], button[aria-expanded='true'] { background: var(--ok-selected); border-color: var(--ok-selected-border); }
  .big { display: flex; flex-direction: column; align-items: center; justify-content: flex-start; gap: 2px; min-width: var(--w); padding: 4px 1px; font-size: 11px; line-height: 1.15; text-align: center; }
  .big > span:last-child { max-width: var(--w); margin: 0 -2px; }
  /* Japanese labels wrap per character, so they get at least six characters
     a line and a smaller size that fits three lines (PowerPoint widens them). */
  .big > span:last-child:lang(ja) { max-width: max(calc(var(--w) - 4px), 6em); font-size: 10px; line-height: 1.1; }
  .icon-row { display: flex; align-items: center; gap: 1px; }
  .arrow { font-size: 11px; }
  svg { width: 32px; height: 32px; flex: none; fill: none; stroke: currentColor; stroke-width: 1.1; }
  .anchor { position: relative; display: flex; }
  /* Two columns of 26 pt rows on a 32 pt pitch, filled top to bottom. */
  .checks { display: grid; grid-auto-flow: column; grid-template-rows: 26px 26px; gap: 6px 0; align-self: flex-start; margin-top: 4px; font-size: 12px; }
  .checks label { display: flex; align-items: center; gap: 6px; padding: 0 8px 0 4px; white-space: nowrap; }
  .checks input { margin: 0; accent-color: var(--ok-accent); }
  .row { display: flex; align-items: center; gap: 4px; height: 26px; padding: 0 3px; font-size: 12px; white-space: nowrap; }
  .row svg { width: 16px; height: 16px; }
  .menu { position: fixed; z-index: 300; display: flex; flex-direction: column; min-width: 200px; padding: 4px 0; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .menu button { height: 24px; padding: 0 16px; border: none; border-radius: 0; font-size: 13px; text-align: left; }
</style>

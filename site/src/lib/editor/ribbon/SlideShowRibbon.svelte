<script lang="ts">
  import { getSlides, getSlideShowProperties, isSlideHidden, setSlideHidden, setSlideShowProperties, type SlideShowProperties } from '@office-kit/pptx';
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
  <button disabled title={t('Rehearse with Coach needs Microsoft 365 online services.')}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="16" height="12" rx="1" /><path d="M11 16v4M7 20h8" /><path d="M19 9a2 2 0 0 1 2 2v3a2 2 0 0 1-4 0v-3a2 2 0 0 1 2-2z" stroke="#0078d4" /></svg>
    {t('Rehearse with Coach')}
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
<div class="group">
  <button disabled={!editor.canPresent} onclick={() => editor.present('rehearse')}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="16" height="12" rx="1" /><circle cx="17" cy="16" r="5" stroke="#0078d4" /><path d="M17 13v3l2 1" stroke="#0078d4" /></svg>
    {t('Rehearse Timings')}
  </button>
  <button disabled={!editor.canPresent} onclick={() => editor.present('start')}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="14" height="11" rx="1" /><path d="M10 14v5M6 19h8" /><circle cx="19" cy="14" r="3" fill="#d13438" stroke="none" /></svg>
    {t('Record')}
  </button>
</div>
<div class="group checks">
  <label><input type="checkbox" disabled checked={false} />{t('Keep Slides Updated')}</label>
  <label><input type="checkbox" checked={show.useTimings} onchange={(event) => setShow('Use Timings', { useTimings: event.currentTarget.checked })} />{t('Use Timings')}</label>
  <label><input type="checkbox" checked={show.showNarration} onchange={(event) => setShow('Play Narrations', { showNarration: event.currentTarget.checked })} />{t('Play Narrations')}</label>
  <label><input type="checkbox" checked={show.showMediaControls ?? true} onchange={(event) => setShow('Show Media Controls', { showMediaControls: event.currentTarget.checked })} />{t('Show Media Controls')}</label>
</div>
<div class="group checks">
  <label title={t('Live subtitles need Microsoft 365 online services.')}><input type="checkbox" disabled />{t('Always Use Subtitles')}</label>
  <button class="subtitle-settings" disabled title={t('Live subtitles need Microsoft 365 online services.')}>{t('Subtitle Settings')} ⌄</button>
</div>

<style>
  .group { display: flex; align-items: center; flex-shrink: 0; gap: 2px; padding: 0 8px; border-right: 1px solid var(--ok-border); }
  button { display: flex; flex-direction: column; align-items: center; justify-content: flex-start; gap: 2px; min-width: 52px; max-width: 96px; min-height: 66px; padding: 3px 4px; font: inherit; font-size: 11px; line-height: 1.2; color: var(--ok-text); background: transparent; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  button:not(:disabled):hover, button[aria-pressed='true'] { background: var(--ok-hover); border-color: var(--ok-border); }
  button:disabled { opacity: 0.45; cursor: default; }
  .checks { display: grid; grid-template-columns: auto auto; gap: 6px 14px; font-size: 12px; }
  .checks label { display: flex; align-items: center; gap: 6px; white-space: nowrap; }
  .checks label:has(input:disabled) { opacity: 0.45; }
  .checks input { margin: 0; accent-color: var(--ok-accent); }
  .checks:last-child { grid-template-columns: auto; }
  .checks .subtitle-settings { flex-direction: row; min-height: 0; max-width: none; padding: 0; font-size: 12px; }
  svg { width: 32px; height: 32px; flex-shrink: 0; stroke: currentColor; fill: none; stroke-width: 1.1; }
</style>

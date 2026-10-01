<script lang="ts">
  import { getCustomShows, getSlideShowProperties, getSlides, setSlideShowProperties, type SlideShowProperties } from '@office-kit/pptx';
  import { onMount, untrack } from 'svelte';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const presentation = untrack(() => doc.pres);
  const version = untrack(() => doc.version);
  const initial = untrack(() => getSlideShowProperties(presentation));
  const slideCount = untrack(() => getSlides(presentation).length);
  const customShows = untrack(() => getCustomShows(presentation));

  let mode = $state(initial.mode.kind);
  let scrollbar = $state(initial.mode.kind === 'browse' ? initial.mode.showScrollbar : true);
  let loop = $state(initial.loop);
  let withoutNarration = $state(!initial.showNarration);
  let withoutAnimation = $state(!initial.showAnimation);
  let advance = $state(initial.useTimings ? 'timings' : 'manual');
  let slides = $state(initial.slides.kind);
  let start = $state(initial.slides.kind === 'range' ? initial.slides.start : 1);
  let end = $state(initial.slides.kind === 'range' ? initial.slides.end : Math.max(1, slideCount));
  let customId = $state(initial.slides.kind === 'customShow' ? initial.slides.id : customShows[0]?.id ?? 0);
  let error = $state('');
  let dialog: HTMLDialogElement;

  const hasCustomShows = customShows.length > 0;
  const valid = $derived(
    (slides !== 'range' || (Number.isInteger(start) && Number.isInteger(end) && start >= 1 && start <= end && end <= slideCount)) &&
      (slides !== 'customShow' || customShows.some(show => show.id === customId)),
  );

  onMount(() => dialog.showModal());

  function submit(event: SubmitEvent) {
    event.preventDefault();
    if (!valid) return;
    if (doc.pres !== presentation || doc.version !== version) {
      error = t('The presentation changed. Reopen Set Up Show.');
      return;
    }
    const settings: SlideShowProperties = {
      mode: mode === 'browse' ? { kind: 'browse', showScrollbar: scrollbar } : mode === 'kiosk' ? { kind: 'kiosk', restart: initial.mode.kind === 'kiosk' ? initial.mode.restart : 300000 } : { kind: 'present' },
      slides: slides === 'range' ? { kind: 'range', start, end } : slides === 'customShow' ? { kind: 'customShow', id: customId } : { kind: 'all' },
      loop: mode === 'kiosk' ? true : loop,
      showNarration: !withoutNarration,
      showAnimation: !withoutAnimation,
      useTimings: advance === 'timings',
    };
    try {
      doc.transact(t('Set Up Show'), () => setSlideShowProperties(presentation, settings));
      editor.closeDialog();
    } catch (cause) {
      error = `${t('Show settings update failed')}: ${cause instanceof Error ? cause.message : String(cause)}`;
    }
  }
</script>

<dialog bind:this={dialog} aria-label={t('Set Up Show')} onclose={() => editor.closeDialog()}>
  <form onsubmit={submit}>
    <header><strong>{t('Set Up Show')}</strong><button type="button" class="close" aria-label={t('Close')} onclick={() => editor.closeDialog()}>✕</button></header>
    <fieldset><legend>{t('Show type')}</legend>
      <label><input type="radio" name="mode" value="present" bind:group={mode} />{t('Presented by a speaker (full screen)')}</label>
      <label><input type="radio" name="mode" value="browse" bind:group={mode} />{t('Browsed by an individual (window)')}</label>
      <label class="indent"><input type="checkbox" bind:checked={scrollbar} disabled={mode !== 'browse'} />{t('Show scrollbar')}</label>
      <label><input type="radio" name="mode" value="kiosk" bind:group={mode} />{t('Browsed at a kiosk (full screen)')}</label>
    </fieldset>
    <div class="columns">
      <fieldset><legend>{t('Show options')}</legend>
        <label><input type="checkbox" checked={mode === 'kiosk' || loop} disabled={mode === 'kiosk'} onchange={(event) => { if (mode !== 'kiosk') loop = (event.currentTarget as HTMLInputElement).checked; }} />{t("Loop continuously until 'Esc'")}</label>
        <label><input type="checkbox" bind:checked={withoutNarration} />{t('Show without narration')}</label>
        <label><input type="checkbox" bind:checked={withoutAnimation} />{t('Show without animation')}</label>
      </fieldset>
      <fieldset><legend>{t('Show slides')}</legend>
        <label><input type="radio" name="slides" value="all" bind:group={slides} />{t('All')}</label>
        <div class="range"><label><input type="radio" name="slides" value="range" bind:group={slides} />{t('From:')}</label><input aria-label={t('From slide')} type="number" min="1" max={slideCount} step="1" bind:value={start} disabled={slides !== 'range'} /><label>{t('To:')} <input aria-label={t('To slide')} type="number" min="1" max={slideCount} step="1" bind:value={end} disabled={slides !== 'range'} /></label></div>
        <label><input type="radio" name="slides" value="customShow" bind:group={slides} disabled={!hasCustomShows} />{t('Custom show:')}</label>
        <select aria-label={t('Custom show')} bind:value={customId} disabled={slides !== 'customShow' || !hasCustomShows}>{#each customShows as show}<option value={show.id}>{show.name}</option>{/each}</select>
      </fieldset>
    </div>
    <fieldset><legend>{t('Advance slides')}</legend>
      <label><input type="radio" name="advance" value="manual" bind:group={advance} />{t('Manually')}</label>
      <label><input type="radio" name="advance" value="timings" bind:group={advance} />{t('Using timings, if present')}</label>
    </fieldset>
    {#if !valid}<p role="alert">{t('Enter a valid slide range and choose an available custom show.')}</p>{/if}
    {#if error}<p role="alert">{error}</p>{/if}
    <footer><button type="button" onclick={() => editor.closeDialog()}>{t('Cancel')}</button><button type="submit" class="primary" disabled={!valid}>{t('OK')}</button></footer>
  </form>
</dialog>

<style>
  dialog { width: min(610px, 92vw); max-height: 90vh; padding: 0; border: 1px solid var(--ok-border); border-radius: var(--ok-radius-lg); background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  dialog::backdrop { background: #0006; }
  form { display: grid; gap: 14px; padding: 18px; }
  header, footer { display: flex; align-items: center; justify-content: space-between; gap: 10px; }
  header { margin: -18px -18px 0; padding: 10px 14px; border-bottom: 1px solid var(--ok-border); }
  .close { border: 0; background: transparent; color: inherit; font-size: 15px; }
  fieldset { display: grid; gap: 7px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius); padding: 10px 12px; }
  legend { padding: 0 4px; font-weight: 600; }
  label { display: flex; align-items: center; gap: 6px; font-size: 12px; }
  .indent { margin-left: 22px; }
  .columns { display: grid; grid-template-columns: 1fr 1.2fr; gap: 12px; }
  .range { display: flex; align-items: center; gap: 5px; flex-wrap: wrap; }
  input[type='number'], select { min-width: 65px; background: var(--ok-input); color: inherit; border: 1px solid var(--ok-border); border-radius: var(--ok-radius); padding: 3px 5px; }
  footer { justify-content: flex-end; }
  footer button { border: 1px solid var(--ok-border); border-radius: var(--ok-radius); padding: 5px 14px; background: var(--ok-panel); color: inherit; }
  footer .primary { background: var(--ok-accent); color: white; }
  p { margin: 0; font-size: 12px; color: #bf3131; }
  @media (max-width: 560px) { .columns { grid-template-columns: 1fr; } }
</style>

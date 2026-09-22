<script lang="ts">
  import { onMount, untrack } from 'svelte';
  import { getSlideLayouts, getSlideLayout, getSlideLayoutName, getSlideLayoutPartName } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  const editor = getEditor();
  const doc = editor.doc;
  const version = untrack(() => doc.version);
  const slide = untrack(() => doc.currentSlide);
  const layouts = untrack(() => getSlideLayouts(doc.pres));
  const current = slide ? getSlideLayout(slide) : null;
  const initial = layouts.find(layout => current && getSlideLayoutPartName(layout) === getSlideLayoutPartName(current)) ?? layouts[0];
  let layoutId = $state(initial ? getSlideLayoutPartName(initial) : '');
  let error = $state('');
  let dialog: HTMLDialogElement;
  onMount(() => dialog.showModal());
  function submit(event: SubmitEvent) {
    event.preventDefault();
    if (doc.version !== version || doc.currentSlide !== slide) { error = t('The slide changed. Reopen this dialog.'); return; }
    const layout = layouts.find(item => getSlideLayoutPartName(item) === layoutId);
    if (!layout) return;
    if (editor.invoke('addSlide', { options: { layout } })) editor.closeDialog();
  }
</script>

<dialog bind:this={dialog} aria-label={t('New slide from layout')} onclose={() => editor.closeDialog()}>
  <form onsubmit={submit}>
    <header><strong>{t('New slide from layout')}</strong><button type="button" class="ok-btn" aria-label={t('Close')} onclick={() => editor.closeDialog()}>✕</button></header>
    <p>{t('Choose a layout to insert a slide after the current slide.')}</p>
    <label>{t('Slide layout')}<select class="ok-input" aria-label={t('Slide layout')} bind:value={layoutId}>
      {#each layouts as layout}<option value={getSlideLayoutPartName(layout)}>{t(getSlideLayoutName(layout))}</option>{/each}
    </select></label>
    {#if !layouts.length}<p>{t('No slide layouts are available.')}</p>{/if}
    {#if error}<p role="alert">{error}</p>{/if}
    <footer><button type="button" class="ok-btn" onclick={() => editor.closeDialog()}>{t('Cancel')}</button><button type="submit" class="ok-btn primary" disabled={!layoutId}>{t('Insert slide')}</button></footer>
  </form>
</dialog>

<style>
  dialog { width: min(520px, 90vw); max-height: 85vh; padding: 18px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius-lg); background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  dialog::backdrop { background: #0006; }
  form, label { display: grid; gap: 14px; }
  header, footer { display: flex; justify-content: space-between; align-items: center; gap: 12px; }
  footer { justify-content: flex-end; }
  p { font-size: 12px; color: var(--ok-text-2); margin: 0; }
  [role='alert'] { color: #bf3131; }
</style>

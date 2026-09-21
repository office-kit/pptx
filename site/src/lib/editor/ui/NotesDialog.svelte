<script lang="ts">
  import { onMount, untrack } from 'svelte';
  import { getSlideNotes, setSlideNotes } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  const editor = getEditor();
  const doc = editor.doc;
  const slide = untrack(() => doc.currentSlide);
  const version = untrack(() => doc.version);
  const original = slide ? getSlideNotes(slide) ?? '' : '';
  let value = $state(original);
  let error = $state('');
  let dialog: HTMLDialogElement;
  onMount(() => dialog.showModal());
  function submit(event: SubmitEvent) {
    event.preventDefault();
    if (!slide) return;
    if (doc.version !== version || doc.currentSlide !== slide) { error = t('The slide changed. Reopen this dialog.'); return; }
    try {
      if (value !== original) doc.transact(t('Speaker notes'), () => setSlideNotes(slide, value));
      editor.closeDialog();
    } catch (cause) { error = `${t('Slide update failed')}: ${cause instanceof Error ? cause.message : String(cause)}`; }
  }
</script>

<dialog bind:this={dialog} aria-label={t('Speaker notes')} onclose={() => editor.closeDialog()}>
  <form onsubmit={submit}>
    <header><strong>{t('Speaker notes')}</strong><button type="button" class="ok-btn" aria-label={t('Close')} onclick={() => editor.closeDialog()}>✕</button></header>
    <label>{t('Notes content')}<textarea class="ok-input" rows="12" bind:value aria-label={t('Notes content')}></textarea></label>
    <p>{t('These notes belong to the selected slide and are included in the exported presentation.')}</p>
    {#if error}<p role="alert">{error}</p>{/if}
    <footer><button type="button" class="ok-btn" onclick={() => editor.closeDialog()}>{t('Cancel')}</button><button type="submit" class="ok-btn primary">{t('Apply')}</button></footer>
  </form>
</dialog>

<style>
  dialog { width: min(620px, 90vw); max-height: 85vh; padding: 18px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius-lg); background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  dialog::backdrop { background: #0006; }
  form, label { display: grid; gap: 14px; }
  header, footer { display: flex; justify-content: space-between; align-items: center; gap: 12px; }
  footer { justify-content: flex-end; }
  textarea { resize: vertical; max-height: 50vh; line-height: 1.5; }
  p { font-size: 12px; color: var(--ok-text-2); margin: 0; }
  [role='alert'] { color: #bf3131; }
</style>

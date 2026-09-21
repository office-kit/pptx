<script lang="ts">
  import { onMount, untrack } from 'svelte';
  import { getShapeHyperlink, getShapeHyperlinkTooltip, getShapeKind, getShapeText, setShapeHyperlink } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { selectedShapeIds } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  const editor = getEditor();
  const doc = editor.doc;
  const selection = untrack(() => doc.selection);
  const version = untrack(() => doc.version);
  const shapes = selectedShapeIds(selection).map(id => doc.shapeById(selection.slideIndex, id));
  const supported = shapes.length > 0 && shapes.every(shape => shape && getShapeKind(shape) === 'shape' && getShapeText(shape).length > 0);
  const urls = shapes.map(shape => shape ? getShapeHyperlink(shape) : null);
  const tips = shapes.map(shape => shape ? getShapeHyperlinkTooltip(shape) : null);
  const mixed = urls.some(url => url !== urls[0]);
  let url = $state(mixed ? '' : urls[0] ?? '');
  let tooltip = $state(tips.every(tip => tip === tips[0]) ? tips[0] ?? '' : '');
  let error = $state('');
  let dialog: HTMLDialogElement;
  onMount(() => dialog.showModal());
  function apply(remove = false) {
    if (!supported || (!remove && !url.trim())) return;
    if (doc.version !== version || doc.selection !== selection) { error = t('The selection changed. Reopen this dialog.'); return; }
    try {
      doc.transact(t(remove ? 'Remove link' : 'Edit link'), () => {
        for (const shape of shapes) if (shape) setShapeHyperlink(shape, remove ? null : url.trim(), remove ? undefined : tooltip.trim() || undefined);
      });
      editor.closeDialog();
    } catch (cause) { error = `${t('Link update failed')}: ${cause instanceof Error ? cause.message : String(cause)}`; }
  }
</script>

<dialog bind:this={dialog} aria-label={t('Edit link')} onclose={() => editor.closeDialog()}>
  <form onsubmit={(event) => { event.preventDefault(); apply(); }}>
    <header><strong>{t('Edit link')}</strong><button type="button" class="ok-btn" aria-label={t('Close')} onclick={() => editor.closeDialog()}>✕</button></header>
    {#if supported}
      <p>{t('Applies to all text in the selected shapes.')}</p>
      <label>{t('Link address')}<input class="ok-input" type="url" required bind:value={url} placeholder="https://example.com" aria-label={t('Link address')} /></label>
      {#if mixed}<p>{t('The selected shapes have different links.')}</p>{/if}
      <label>{t('Link description')}<input class="ok-input" bind:value={tooltip} aria-label={t('Link description')} /></label>
    {:else}<p role="alert">{t('Select text shapes to edit their links.')}</p>{/if}
    {#if error}<p role="alert">{error}</p>{/if}
    <footer><button type="button" class="ok-btn" disabled={!supported || !urls.some(Boolean)} onclick={() => apply(true)}>{t('Remove link')}</button><span></span><button type="button" class="ok-btn" onclick={() => editor.closeDialog()}>{t('Cancel')}</button><button type="submit" class="ok-btn primary" disabled={!supported || !url.trim()}>{t('Apply')}</button></footer>
  </form>
</dialog>

<style>
  dialog { width: min(520px, 90vw); max-height: 85vh; padding: 18px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius-lg); background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  dialog::backdrop { background: #0006; }
  form, label { display: grid; gap: 12px; }
  header, footer { display: flex; align-items: center; gap: 10px; }
  header { justify-content: space-between; }
  footer span { flex: 1; }
  label, p { font-size: 12px; }
  p { color: var(--ok-text-2); margin: 0; }
  [role='alert'] { color: #bf3131; }
</style>

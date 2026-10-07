<script lang="ts">
  // Mac PowerPoint's Compress Pictures sheet (OfficeArt CompressPictDlg):
  // Picture Quality, Delete cropped areas of pictures, Apply to.
  import { onMount, untrack } from 'svelte';
  import { getShapeId, setShapeImage, setShapeImageCrop } from '@office-kit/pptx';
  import { collectPictures, compressPictures, PICTURE_QUALITIES, type PictureQuality } from '../core/compress-pictures.ts';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  const editor = getEditor();
  const doc = editor.doc;
  const presentation = untrack(() => doc.pres);
  const selectedIds = new Set(untrack(() => editor.selectedShapes()).map(getShapeId));
  const selectedSlide = untrack(() => doc.slideAt(doc.selection.slideIndex));
  const all = collectPictures(presentation);
  const selected = all.filter((item) => item.slide === selectedSlide && item.ids.some((id) => selectedIds.has(id)));
  // A new Mac deck stores p14:defaultImageDpi 32767 — High Fidelity.
  let quality = $state<PictureQuality>('highFidelity');
  let deleteCropped = $state(true);
  let scope = $state<'all' | 'selected'>(selected.length ? 'selected' : 'all');
  let busy = $state(false);
  let error = $state('');
  let dialog: HTMLDialogElement;
  onMount(() => dialog.showModal());

  async function submit(event: SubmitEvent) {
    event.preventDefault();
    if (busy) return;
    busy = true;
    try {
      const ppi = PICTURE_QUALITIES.find((item) => item.id === quality)!.ppi;
      const results = await compressPictures(scope === 'selected' ? selected : all, ppi, deleteCropped);
      if (doc.pres !== presentation) { error = t('The presentation changed. Reopen Compress Pictures.'); return; }
      if (results.length) {
        doc.transact(t('Compress Pictures'), () => {
          for (const result of results) {
            setShapeImage(result.target.shape, result.bytes);
            setShapeImageCrop(result.target.shape, result.crop);
          }
        });
      }
      editor.closeDialog();
    } catch (cause) {
      error = String(cause);
    } finally {
      busy = false;
    }
  }
</script>
<dialog bind:this={dialog} aria-label={t('Compress Pictures')} onclose={() => editor.closeDialog()}>
  <form onsubmit={submit}>
    <h2>{t('Compress Pictures')}</h2>
    <p>{t('Compress pictures to reduce the size of this file.')}</p>
    <label class="quality">{t('Picture Quality:')} <select bind:value={quality} disabled={busy}>{#each PICTURE_QUALITIES as item (item.id)}<option value={item.id}>{t(item.label)}</option>{/each}</select></label>
    <label><input type="checkbox" bind:checked={deleteCropped} disabled={busy} />{t('Delete cropped areas of pictures')}</label>
    <fieldset aria-label={t('Apply to:')}>
      <legend>{t('Apply to:')}</legend>
      <label><input type="radio" name="compress-scope" value="all" bind:group={scope} disabled={busy} />{t('All pictures in this file')}</label>
      <label><input type="radio" name="compress-scope" value="selected" bind:group={scope} disabled={busy || !selected.length} />{t('Selected pictures only')}</label>
    </fieldset>
    {#if busy}<p role="status">{t('Compressing pictures...')}</p>{/if}
    {#if error}<p role="alert">{error}</p>{/if}
    <footer><button type="button" onclick={() => editor.closeDialog()}>{t('Cancel')}</button><button type="submit" disabled={busy}>{t('OK')}</button></footer>
  </form>
</dialog>
<style>
  dialog { width: 340px; border: 1px solid #666; border-radius: 7px; background: #303030; color: #eee; padding: 0; box-shadow: 0 15px 60px #0008; font-size: 13px; color-scheme: dark; }
  dialog::backdrop { background: #0003; }
  form { display: grid; gap: 10px; padding-bottom: 16px; }
  h2 { font-size: 13px; font-weight: 600; text-align: center; margin: 0; padding: 6px; background: #3a3a3a; border-bottom: 1px solid #242424; }
  p { margin: 0 19px; }
  label { display: flex; align-items: center; gap: 6px; margin: 0 19px; }
  .quality { justify-content: space-between; }
  select { border: 0; border-radius: 5px; padding: 3px 8px; background: #4a4a4a; color: inherit; font: inherit; }
  fieldset { display: grid; gap: 6px; border: 0; margin: 0; padding: 0; }
  legend { margin: 0 19px 6px; padding: 0; }
  fieldset label { margin-left: 34px; }
  input { margin: 0; accent-color: #1685f8; }
  footer { display: flex; justify-content: flex-end; gap: 10px; padding: 0 17px; }
  footer button { min-width: 70px; padding: 3px 8px; border: 0; border-radius: 5px; background: #626262; color: inherit; font-size: 12px; }
  footer button[type='submit'] { background: #087bfa; color: white; }
</style>

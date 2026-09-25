<script lang="ts">
  import { onMount, untrack } from 'svelte';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  const editor = getEditor();
  const initial = untrack(() => Math.round((editor.viewMode === 'sorter' ? editor.sorterZoom : editor.zoom) * 100));
  const presets = [400, 200, 100, 66, 50, 33];
  let value = $state(initial);
  let selected = $state(untrack(() => editor.viewMode === 'normal' && editor.autoFitZoom) ? 'fit' : presets.includes(initial) ? String(initial) : '');
  let dialog: HTMLDialogElement;
  onMount(() => dialog.showModal());
  function submit(event: SubmitEvent) {
    event.preventDefault();
    if (selected === 'fit') editor.zoomFit();
    else {
      if (!Number.isFinite(value) || value < editor.minZoomPercent || value > editor.maxZoomPercent) return;
      editor.setZoom(value / 100);
    }
    editor.closeDialog();
  }
</script>
<dialog bind:this={dialog} aria-label={t('Zoom')} onclose={() => editor.closeDialog()}>
  <form onsubmit={submit}>
    <h2>{t('Zoom')}</h2>
    <fieldset aria-label={t('Zoom to')}>
      <label><input type="radio" name="zoom-preset" value="fit" disabled={editor.viewMode === 'sorter'} bind:group={selected} />{t('Fit')}</label>
      {#each presets as preset}<label><input type="radio" name="zoom-preset" value={String(preset)} disabled={preset > editor.maxZoomPercent} bind:group={selected} onchange={() => value = preset} />{preset}%</label>{/each}
    </fieldset>
    <label class="percent">{t('Percent:')}<input type="number" min={selected === 'fit' ? undefined : editor.minZoomPercent} max={selected === 'fit' ? undefined : editor.maxZoomPercent} required={selected !== 'fit'} bind:value oninput={() => selected = ''} /></label>
    <footer><button type="button" onclick={() => editor.closeDialog()}>{t('Cancel')}</button><button type="submit">{t('OK')}</button></footer>
  </form>
</dialog>
<style>
  dialog { width: 230px; border: 1px solid #666; border-radius: 7px; background: #303030; color: #eee; padding: 0; box-shadow: 0 15px 60px #0008; font-size: 13px; color-scheme: dark; }
  dialog::backdrop { background: #0003; }
  h2 { font-size: 13px; font-weight: 600; text-align: center; margin: 0 0 12px; padding: 6px; background: #3a3a3a; border-bottom: 1px solid #242424; }
  fieldset { display: flex; flex-direction: column; gap: 7px; border: 0; margin: 0; padding: 0 19px; }
  fieldset label { display: flex; align-items: center; gap: 6px; height: 16px; }
  fieldset input { margin: 0; width: 14px; height: 14px; accent-color: #1685f8; }
  .percent { display: flex; align-items: center; justify-content: space-between; margin: 14px 19px; }
  .percent input { width: 70px; padding: 2px 4px; border: 1px solid #777; border-radius: 4px; background: #252525; color: inherit; font: inherit; }
  footer { display: flex; justify-content: flex-end; gap: 10px; padding: 0 17px 16px; }
  footer button { min-width: 70px; padding: 3px 8px; border: 0; border-radius: 5px; background: #626262; color: inherit; font-size: 12px; }
  footer button[type='submit'] { background: #087bfa; color: white; }
</style>

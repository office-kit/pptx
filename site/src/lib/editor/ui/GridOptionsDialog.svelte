<script lang="ts">
  import { onMount, untrack } from 'svelte';
  import { getGridSpacing, getSnapToGrid, getDrawingGuidesVisible, setGridSpacing, setSnapToGrid } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  const editor = getEditor();
  const doc = editor.doc;
  const presentation = untrack(() => doc.pres);
  const initial = getGridSpacing(presentation) ?? { x: 72000, y: 72000 };
  let dialog: HTMLDialogElement;
  let snap = $state(getSnapToGrid(presentation) ?? false);
  let grid = $state(untrack(() => editor.view.grid));
  let smart = $state(untrack(() => editor.view.smart));
  let drawing = $state(untrack(() => editor.view.drawing) ?? getDrawingGuidesVisible(presentation) ?? false);
  const presets = [8, 6, 5, 4, 3, 2].map(count => ({ label: `${count} grids per cm`, value: 360000 / count }))
    .concat([1, 2, 3, 4, 5].map(count => ({ label: `${count} cm`, value: count * 360000 })));
  let spacing = $state(String(presets.find(p => Math.abs(p.value - initial.x) < 20 && Math.abs(p.value - initial.y) < 20)?.value ?? 'custom'));
  let spacingChanged = $state(false);
  let customCm = $state(initial.x / 360000);
  let customChanged = $state(false);
  let makeDefault = $state(false);
  let error = $state('');
  onMount(() => dialog.showModal());
  function submit(event: SubmitEvent) {
    event.preventDefault();
    if (doc.pres !== presentation) { error = t('The presentation changed. Reopen Grid Options.'); return; }
    const customEmu = Math.round(Number(customCm) * 360000);
    if (spacing === 'custom' && customChanged && (!Number.isSafeInteger(customEmu) || customEmu <= 0)) { error = t('Enter a positive grid spacing.'); return; }
    const next = spacing === 'custom'
      ? (customChanged ? { x: customEmu, y: customEmu } : initial)
      : (!spacingChanged ? initial : { x: Number(spacing), y: Number(spacing) });
    try {
      const previous = getGridSpacing(presentation) ?? initial;
      if (snap !== (getSnapToGrid(presentation) ?? false) || next.x !== previous.x || next.y !== previous.y) {
        doc.transact(t('Grid Options'), () => {
          setGridSpacing(presentation, next);
          setSnapToGrid(presentation, snap);
        });
      }
      editor.view.save({ grid, smart, drawing });
      if (makeDefault) {
        try { localStorage.setItem('office-grid-defaults', JSON.stringify({ ...next, snap })); } catch { /* Session remains usable. */ }
      }
      editor.closeDialog();
    } catch (cause) { error = String(cause); }
  }
</script>
<dialog bind:this={dialog} aria-label={t('Grid and Guides')} onclose={() => editor.closeDialog()}>
  <form onsubmit={submit}>
    <h2>{t('Grid and Guides')}</h2>
    <section><h3>{t('Snap to')}</h3><div class="settings"><label><input type="checkbox" bind:checked={snap} />{t('Snap objects to grid')}</label></div></section>
    <section><h3>{t('Grid Settings')}</h3><div class="settings">
      <label>{t('Spacing:')} <select bind:value={spacing} onchange={() => spacingChanged = true}>{#each presets as preset}<option value={String(preset.value)}>{t(preset.label)}</option>{/each}<option value="custom">{t('Custom')}</option></select></label>
      {#if spacing === 'custom'}<label>{t('Custom spacing:')} <input class="custom-spacing" type="number" step="any" min={1 / 360000} required bind:value={customCm} oninput={() => customChanged = true} />cm</label>{/if}
      <label><input type="checkbox" bind:checked={grid} />{t('Display grid on screen')}</label>
    </div></section>
    <section><h3>{t('Guide Settings')}</h3><div class="settings">
      <label><input type="checkbox" bind:checked={drawing} />{t('Display drawing guides on screen')}</label>
      <label><input type="checkbox" bind:checked={smart} />{t('Display smart guides when shapes are aligned')}</label>
    </div></section>
    {#if error}<p role="alert">{error}</p>{/if}
    <footer><button type="button" class="ok-btn" aria-pressed={makeDefault} onclick={() => makeDefault = true}>{t('Set as Default')}</button><span></span><button type="button" class="ok-btn" onclick={() => editor.closeDialog()}>{t('Cancel')}</button><button type="submit" class="ok-btn primary">{t('OK')}</button></footer>
  </form>
</dialog>
<style>
  dialog { width: min(390px, 90vw); padding: 0; border: 1px solid #515151; border-radius: 26px; background: #1c1c1c; color: #dedede; box-shadow: 0 15px 35px #0007; color-scheme: dark; font-family: -apple-system, BlinkMacSystemFont, sans-serif; }
  dialog::backdrop { background: #0006; }
  form { display: grid; gap: 18px; padding-bottom: 18px; }
  h2 { margin: 0; padding: 9px; font-size: 13px; text-align: center; border-bottom: 1px solid #333; }
  section { margin: 0 18px; }
  h3 { margin: 0 10px 4px; font-size: 13px; font-weight: 600; }
  .settings { border-radius: 12px; padding: 20px; display: grid; gap: 5px; background: #242424; }
  label { display: flex; align-items: center; gap: 5px; font-size: 12px; white-space: nowrap; }
  input { margin: 0; width: 16px; height: 16px; accent-color: #707070; }
  .custom-spacing { width: 90px; height: auto; border: 1px solid #555; border-radius: 4px; padding: 3px 6px; }
  select { border: 0; border-radius: 5px; padding: 3px 8px; background: #383838; color: inherit; font: inherit; }
  footer { display: flex; align-items: center; gap: 8px; margin: 0 18px; }
  footer span { flex: 1; }
  footer button { padding: 3px 12px; border: 0; border-radius: 5px; background: #333; color: inherit; font-size: 12px; }
  footer button:hover { background: #464646; }
  [role='alert'] { margin: 0 18px; }
</style>

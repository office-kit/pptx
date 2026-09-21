<script lang="ts">
  import { onMount, untrack } from 'svelte';
  import { emu, getSlideSize, setSlideSize, SLIDE_SIZE_4_3, SLIDE_SIZE_16_9, SLIDE_SIZE_16_10 } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  const editor = getEditor();
  const doc = editor.doc;
  const presentation = untrack(() => doc.pres);
  const version = untrack(() => doc.version);
  const initial = getSlideSize(presentation) ?? SLIDE_SIZE_16_9;
  const presets = [
    { id: 'wide', label: 'Widescreen (16:9)', size: SLIDE_SIZE_16_9 },
    { id: 'standard', label: 'Standard (4:3)', size: SLIDE_SIZE_4_3 },
    { id: 'wide10', label: 'Widescreen (16:10)', size: SLIDE_SIZE_16_10 },
  ];
  let dialog: HTMLDialogElement;
  let preset = $state(presets.find(item => item.size.width === initial.width && item.size.height === initial.height)?.id ?? 'custom');
  let unit = $state('in');
  let width = $state<number | undefined>(initial.width / 914400);
  let height = $state<number | undefined>(initial.height / 914400);
  let error = $state('');
  const factor = $derived(unit === 'cm' ? 360000 : 914400);
  const maximum = $derived(56 * 914400 / factor);
  const valid = $derived(width !== undefined && height !== undefined && Number.isFinite(width) && Number.isFinite(height) && width * factor >= 914400 && height * factor >= 914400 && width <= maximum && height <= maximum);
  onMount(() => dialog.showModal());
  function choosePreset(id: string) {
    preset = id;
    const chosen = presets.find(item => item.id === id);
    if (chosen) { width = chosen.size.width / factor; height = chosen.size.height / factor; }
  }
  function changeUnit(next: string) {
    const nextFactor = next === 'cm' ? 360000 : 914400;
    if (width !== undefined) width = width * factor / nextFactor;
    if (height !== undefined) height = height * factor / nextFactor;
    unit = next;
  }
  function submit(event: SubmitEvent) {
    event.preventDefault();
    if (!valid || width === undefined || height === undefined) return;
    if (presentation !== doc.pres || version !== doc.version) { error = t('The presentation changed. Reopen page setup.'); return; }
    const size = presets.find(item => item.id === preset)?.size ?? { width: emu(width * factor), height: emu(height * factor) };
    try {
      doc.transact(t('Page setup'), () => setSlideSize(presentation, size));
      editor.closeDialog();
    } catch (cause) { error = `${t('Slide update failed')}: ${cause instanceof Error ? cause.message : String(cause)}`; }
  }
</script>

<dialog bind:this={dialog} aria-label={t('Page setup')} onclose={() => editor.closeDialog()}>
  <form onsubmit={submit}>
    <header><strong>{t('Page setup')}</strong><button type="button" class="ok-btn" aria-label={t('Close')} onclick={() => editor.closeDialog()}>✕</button></header>
    <label>{t('Slide size')}<select class="ok-input" aria-label={t('Slide size')} value={preset} onchange={event => choosePreset(event.currentTarget.value)}>
      {#each presets as item}<option value={item.id}>{t(item.label)}</option>{/each}<option value="custom">{t('Custom size')}</option>
    </select></label>
    <label>{t('Units')}<select class="ok-input" aria-label={t('Units')} value={unit} onchange={event => changeUnit(event.currentTarget.value)}><option value="in">{t('Inches')}</option><option value="cm">{t('Centimeters')}</option></select></label>
    <div class="dimensions">
      <label>{t('Page width')}<input class="ok-input" type="number" min={914400 / factor} max={maximum} step="any" required bind:value={width} oninput={() => preset = 'custom'} /></label>
      <label>{t('Page height')}<input class="ok-input" type="number" min={914400 / factor} max={maximum} step="any" required bind:value={height} oninput={() => preset = 'custom'} /></label>
    </div>
    <p>{t('Applies to every slide. Objects keep their positions and sizes.')}</p>
    {#if !valid}<p role="alert">{t(unit === 'cm' ? 'Enter a width and height between 2.54 and 142.24 cm.' : 'Enter a width and height between 1 and 56 inches.')}</p>{/if}
    {#if error}<p role="alert">{error}</p>{/if}
    <footer><button type="button" class="ok-btn" onclick={() => editor.closeDialog()}>{t('Cancel')}</button><button type="submit" class="ok-btn primary" disabled={!valid}>{t('Apply')}</button></footer>
  </form>
</dialog>

<style>
  dialog { width: min(420px, 90vw); max-height: 85vh; padding: 18px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius-lg); background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  dialog::backdrop { background: #0006; }
  form { display: grid; gap: 14px; }
  header, footer { display: flex; align-items: center; justify-content: space-between; gap: 12px; }
  footer { justify-content: flex-end; }
  label { display: grid; gap: 6px; }
  .dimensions { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
  input { width: 100%; }
  p { font-size: 12px; color: var(--ok-text-2); margin: 0; }
  [role='alert'] { color: #bf3131; }
</style>

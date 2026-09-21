<script lang="ts">
  import { getSlideBackground, getSlideLayout, getSlideLayouts, getSlideLayoutName, getSlideLayoutPartName, setSlideBackground, setSlideBackgroundImage, clearSlideBackground, setSlideLayout } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  const editor = getEditor();
  const doc = editor.doc;
  const slide = $derived.by(() => { doc.version; return doc.currentSlide; });
  const background = $derived.by(() => { doc.version; return slide ? getSlideBackground(slide) : null; });
  const layouts = $derived.by(() => { doc.version; return getSlideLayouts(doc.pres); });
  const layout = $derived.by(() => { doc.version; return slide ? getSlideLayout(slide) : null; });
  const layoutId = $derived(layout ? getSlideLayoutPartName(layout) : '');
  let fileInput = $state<HTMLInputElement>();
  let error = $state('');
  let loading = $state(false);
  function apply(label: string, operation: () => void) {
    try { doc.transact(t(label), operation); error = ''; }
    catch (cause) { error = cause instanceof Error ? cause.message : String(cause); }
  }
  function changeLayout(id: string) {
    const target = layouts.find(item => getSlideLayoutPartName(item) === id);
    if (slide && target) apply('Slide layout', () => setSlideLayout(slide!, target));
  }
  async function upload(event: Event) {
    const input = event.currentTarget;
    if (!(input instanceof HTMLInputElement)) return;
    const file = input.files?.[0];
    const target = slide;
    if (!file || !target) return;
    const presentation = doc.pres;
    const version = doc.version;
    loading = true;
    error = '';
    try {
      const bytes = new Uint8Array(await file.arrayBuffer());
      if (doc.pres !== presentation || doc.version !== version || doc.currentSlide !== target) {
        error = t('The slide changed. Choose the background image again.');
        return;
      }
      apply('Background image', () => setSlideBackgroundImage(target, bytes));
    } catch (cause) { error = cause instanceof Error ? cause.message : String(cause); }
    finally { loading = false; input.value = ''; }
  }
</script>

{#if slide && (doc.selection.kind === 'none' || doc.selection.kind === 'slide')}
  <section aria-label={t('Slide options')}>
    <strong>{t('Slide options')}</strong>
    <label>{t('Slide layout')}<select class="ok-input" aria-label={t('Slide layout')} value={layoutId} onchange={event => changeLayout(event.currentTarget.value)}>
      {#if !layout}<option value="">{t('None')}</option>{/if}
      {#each layouts as item}<option value={getSlideLayoutPartName(item)}>{t(getSlideLayoutName(item))}</option>{/each}
    </select></label>
    <label>{t('Background color')}<input type="color" aria-label={t('Background color')} value={background?.kind === 'solid' && /^#[0-9a-f]{6}$/i.test(background.color) ? background.color : '#ffffff'} onchange={event => { const color = event.currentTarget.value; apply('Background color', () => setSlideBackground(slide!, color)); }} /></label>
    <input bind:this={fileInput} aria-label={t('Background image')} type="file" accept="image/*" hidden disabled={loading} onchange={upload} />
    <button class="ok-btn" disabled={loading} onclick={() => fileInput?.click()}>{t('Choose background image')}</button>
    <button class="ok-btn" disabled={background?.kind === 'inherit'} onclick={() => apply('Reset background', () => clearSlideBackground(slide!))}>{t('Reset background')}</button>
    {#if error}<p role="alert">{t('Slide update failed')}: {error}</p>{/if}
  </section>
{/if}

<style>
  section { display: grid; gap: 10px; padding: 12px; border-bottom: 1px solid var(--ok-border); }
  strong { font-size: 12px; }
  label { display: grid; gap: 6px; font-size: 11px; }
  input[type='file'] { width: 100%; font-size: 11px; }
  input[type='color'] { width: 100%; height: 26px; }
  [role='alert'] { color: #bf3131; font-size: 11px; }
</style>

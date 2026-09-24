<script lang="ts">
  import { getSlides, isSlideBackgroundGraphicsHidden, setSlideBackgroundGraphicsHidden, asColor, type Color, type SlideData, getSlideBackground, setSlideBackground, setSlideBackgroundImage, setSlideBackgroundGradientFill, setSlideBackgroundPatternFill, clearSlideBackground } from '@office-kit/pptx';
  import { readSlideBackground } from '../core/slide-background.ts';
  import BackgroundPictureLayout from './BackgroundPictureLayout.svelte';
  import PatternFillSection from './PatternFillSection.svelte';
  import GradientFillSection from './GradientFillSection.svelte';
  import ColorPicker from '../ui/ColorPicker.svelte';
  import { selectedSlideIndices } from '../core/selection.ts';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  const editor = getEditor();
  const doc = editor.doc;
  const slide = $derived.by(() => { doc.version; return doc.currentSlide; });
  const slides = $derived.by(() => { doc.version; const all = getSlides(doc.pres); return selectedSlideIndices(doc.selection).flatMap(index => all[index] ? [all[index]!] : []); });
  const background = $derived.by(() => { doc.version; return slide ? readSlideBackground(doc.pres, slide).fill : null; });
  const mixedBackground = $derived(slides.some(item => JSON.stringify(readSlideBackground(doc.pres, item).fill) !== JSON.stringify(background)));
  const gradientBackground = $derived(slides.length > 0 && slides.every(item => readSlideBackground(doc.pres, item).fill.kind === 'gradient'));
  const imageBackground = $derived(slides.length > 0 && slides.every(item => readSlideBackground(doc.pres, item).fill.kind === 'image'));
  const patternBackground = $derived(slides.length > 0 && slides.every(item => readSlideBackground(doc.pres, item).fill.kind === 'pattern'));
  const solidBackground = $derived(slides.length > 0 && slides.every(item => ['solid', 'inherit'].includes(readSlideBackground(doc.pres, item).fill.kind)));
  const graphicsHidden = $derived(slides.length > 0 && slides.every(isSlideBackgroundGraphicsHidden));
  const mixedGraphics = $derived(!graphicsHidden && slides.some(isSlideBackgroundGraphicsHidden));
  const opacity = $derived(background?.kind === 'solid' ? background.opacity ?? 1 : 1);
  const mixedOpacity = $derived(slides.some(item => { const value = readSlideBackground(doc.pres, item).fill; return (value.kind === 'solid' ? value.opacity ?? 1 : 1) !== opacity; }));
  function changeSolid(color?: Color, opacity?: number) {
    apply('Background color', target => {
      const current = readSlideBackground(doc.pres, target).fill;
      setSlideBackground(target, color ?? (current.kind === 'solid' ? asColor(current.color) : null) ?? '#FFFFFF', opacity ?? (current.kind === 'solid' ? current.opacity : undefined));
    });
  }
  const canReset = $derived(slides.some(item => getSlideBackground(item)?.kind !== 'inherit'));
  let fileInput = $state<HTMLInputElement>();
  let error = $state('');
  let loading = $state(false);
  function apply(label: string, operation: (target: SlideData) => void, targets = slides) {
    try { doc.transact(t(label), () => { for (const target of targets) operation(target); }); error = ''; }
    catch (cause) { error = cause instanceof Error ? cause.message : String(cause); }
  }
  async function upload(event: Event) {
    const input = event.currentTarget;
    if (!(input instanceof HTMLInputElement)) return;
    const file = input.files?.[0];
    if (!file) return;
    try { await insertImage(() => file.arrayBuffer()); }
    finally { input.value = ''; }
  }
  async function pasteImage() {
    await insertImage(async () => {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        const type = item.types.find(type => type.startsWith('image/'));
        if (type) return (await item.getType(type)).arrayBuffer();
      }
      throw new Error(t('The clipboard does not contain a picture.'));
    });
  }
  async function insertImage(read: () => Promise<ArrayBuffer>) {
    if (loading || !slides.length) return;
    const targets = slides;
    const selection = doc.selection;
    const presentation = doc.pres;
    const version = doc.version;
    loading = true;
    error = '';
    try {
      const bytes = new Uint8Array(await read());
      if (doc.pres !== presentation || doc.version !== version || doc.selection !== selection) {
        error = t('The slide changed. Choose the background image again.');
        return;
      }
      apply('Background image', target => setSlideBackgroundImage(target, bytes), targets);
    } catch (cause) { error = cause instanceof DOMException && cause.name === 'NotAllowedError' ? t('Clipboard access was denied') : cause instanceof Error ? cause.message : String(cause); }
    finally { loading = false; }
  }
</script>

{#if slide}
  <section aria-label={t('Format Background')}>
    <details open>
      <summary>{t('Fill')}</summary>
      <div class="fill-controls">
        <div class="background-types" role="radiogroup" aria-label={t('Background fill')}>
          <label class="check"><input type="radio" name="background-fill" checked={solidBackground} onchange={() => apply('Background color', target => setSlideBackground(target, '#FFFFFF'))} />{t('Solid fill')}</label>
          <label class="check"><input type="radio" name="background-fill" checked={gradientBackground} onchange={() => apply('Gradient fill', target => setSlideBackgroundGradientFill(target, { stops: [{ offset: 0, color: 'accent1', brightness: 0.95 }, { offset: 1, color: 'accent1', brightness: 0.7 }], angleDeg: 0, scaled: false }))} />{t('Gradient fill')}</label>
          <label class="check"><input type="radio" name="background-fill" checked={imageBackground} disabled={loading} onclick={event => { if (!imageBackground) { event.preventDefault(); fileInput?.click(); } }} />{t('Picture or texture fill')}</label>
          <label class="check"><input type="radio" name="background-fill" checked={patternBackground} onchange={() => apply('Pattern fill', target => setSlideBackgroundPatternFill(target, {}))} />{t('Pattern fill')}</label>
        </div>
        <label class="check"><input type="checkbox" checked={graphicsHidden} indeterminate={mixedGraphics} onchange={event => { const hidden = event.currentTarget.checked; apply('Hide Background Graphics', target => setSlideBackgroundGraphicsHidden(target, hidden)); }} />{t('Hide Background Graphics')}</label>
        {#if gradientBackground}
          <GradientFillSection background />
        {:else if imageBackground}
          <span class="selection">{t('Picture source')}</span>
          <button class="ok-btn" disabled={loading} onclick={() => fileInput?.click()}>{t('Insert...')}</button>
          <button class="ok-btn" disabled={loading || !navigator.clipboard?.read} onclick={pasteImage}>{t('Clipboard')}</button>
          <BackgroundPictureLayout />
        {:else if patternBackground}
          <PatternFillSection background />
        {:else}
          <div class="color-field">{t('Background color')}<ColorPicker label={t('Background color')} value={!mixedBackground && background?.kind === 'solid' ? background.color : undefined} choose={color => changeSolid(color)} /></div>
        {/if}
        {#if solidBackground}
          <label>{t('Transparency')}<div class="transparency">
            <input type="range" aria-label={t('Background transparency')} min="0" max="100" value={Math.round((1 - opacity) * 100)} onchange={event => changeSolid(undefined, 1 - Number(event.currentTarget.value) / 100)} />
            <input class="ok-input" type="number" aria-label={t('Background transparency')} min="0" max="100" placeholder={mixedOpacity ? t('Mixed') : undefined} value={mixedOpacity ? '' : Math.round((1 - opacity) * 100)} onchange={event => { if (event.currentTarget.value !== '' && event.currentTarget.validity.valid) changeSolid(undefined, 1 - Number(event.currentTarget.value) / 100); }} /><span>%</span>
          </div></label>
        {/if}
        {#if mixedBackground}<span class="selection">{t('Background color')}: {t('Mixed')}</span>{/if}
        <input bind:this={fileInput} aria-label={t('Background image')} type="file" accept="image/*" hidden disabled={loading} onchange={upload} />
        {#if !imageBackground}<button class="ok-btn" disabled={loading} onclick={() => fileInput?.click()}>{t('Choose background image')}</button>{/if}
      </div>
    </details>
    <div class="actions">
      <button class="ok-btn" onclick={() => editor.invoke('applySlideBackgroundToAll')}>{t('Apply to All')}</button>
      <button class="ok-btn" disabled={!canReset} onclick={() => apply('Reset background', target => clearSlideBackground(target))}>{t('Reset background')}</button>
    </div>
    {#if error}<p role="alert">{t('Slide update failed')}: {error}</p>{/if}
  </section>
{/if}
<style>
  section { flex: 1; min-height: 0; display: flex; flex-direction: column; }
  details { flex: 1; min-height: 0; overflow-y: auto; padding: 12px; }
  .transparency { display: grid; grid-template-columns: minmax(0, 1fr) 55px auto; gap: 6px; align-items: center; }
  .transparency input { min-width: 0; width: 100%; }
  .background-types { display: grid; gap: 6px; }
  .selection { font-size: 11px; color: var(--ok-muted); }
  summary { font-size: 12px; cursor: pointer; }
  .fill-controls { display: grid; gap: 10px; padding-top: 12px; }
  .actions { flex-shrink: 0; display: flex; justify-content: center; flex-wrap: wrap; gap: 8px; padding: 10px 12px; }
  label, .color-field { display: grid; gap: 6px; font-size: 11px; }
  .check { display: flex; align-items: center; }
  input[type='file'] { width: 100%; font-size: 11px; }
  [role='alert'] { flex-shrink: 0; margin: 0; padding: 0 12px 10px; color: #bf3131; font-size: 11px; }
</style>

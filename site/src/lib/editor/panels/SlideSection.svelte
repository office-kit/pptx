<script lang="ts">
  import { getSlides, asColor, type Color, type SlideData, isSlideHidden, setSlideHidden, getSlideBackground, getSlideLayout, getSlideLayouts, getSlideLayoutName, getSlideLayoutPartName, setSlideBackground, setSlideBackgroundImage, setSlideBackgroundGradientFill, setSlideBackgroundPatternFill, clearSlideBackground, setSlideLayout } from '@office-kit/pptx';
  import { readSlideBackground } from '../core/slide-background.ts';
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
  const skipped = $derived(slides.length > 0 && slides.every(isSlideHidden));
  // Deck-wide, unlike everything else here, so it does not read from `slides`.
  const slideNumbers = $derived.by(() => { doc.version; return editor.slideNumbersOn(); });
  const mixedSkipped = $derived(!skipped && slides.some(isSlideHidden));
  const background = $derived.by(() => { doc.version; return slide ? readSlideBackground(doc.pres, slide).fill : null; });
  const layouts = $derived.by(() => { doc.version; return getSlideLayouts(doc.pres); });
  const layout = $derived.by(() => { doc.version; return slide ? getSlideLayout(slide) : null; });
  const layoutId = $derived(layout ? getSlideLayoutPartName(layout) : '');
  const mixedLayout = $derived(slides.some(item => { const value = getSlideLayout(item); return (value ? getSlideLayoutPartName(value) : '') !== layoutId; }));
  const mixedBackground = $derived(slides.some(item => JSON.stringify(readSlideBackground(doc.pres, item).fill) !== JSON.stringify(background)));
  const gradientBackground = $derived(slides.length > 0 && slides.every(item => readSlideBackground(doc.pres, item).fill.kind === 'gradient'));
  const patternBackground = $derived(slides.length > 0 && slides.every(item => readSlideBackground(doc.pres, item).fill.kind === 'pattern'));
  const solidBackground = $derived(slides.length > 0 && slides.every(item => ['solid', 'inherit'].includes(readSlideBackground(doc.pres, item).fill.kind)));
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
  function changeLayout(id: string) {
    const target = layouts.find(item => getSlideLayoutPartName(item) === id);
    if (slide && target) apply('Slide layout', item => setSlideLayout(item, target));
  }
  async function upload(event: Event) {
    const input = event.currentTarget;
    if (!(input instanceof HTMLInputElement)) return;
    const file = input.files?.[0];
    const targets = slides;
    const selection = doc.selection;
    if (!file || !targets.length) return;
    const presentation = doc.pres;
    const version = doc.version;
    loading = true;
    error = '';
    try {
      const bytes = new Uint8Array(await file.arrayBuffer());
      if (doc.pres !== presentation || doc.version !== version || doc.selection !== selection) {
        error = t('The slide changed. Choose the background image again.');
        return;
      }
      apply('Background image', target => setSlideBackgroundImage(target, bytes), targets);
    } catch (cause) { error = cause instanceof Error ? cause.message : String(cause); }
    finally { loading = false; input.value = ''; }
  }
</script>

{#if slide && (doc.selection.kind === 'none' || doc.selection.kind === 'slide')}
  <section aria-label={t('Slide options')}>
    <strong>{t('Slide options')}</strong>
    {#if slides.length > 1}<span class="selection">{t('Selected slides')}: {slides.length}</span>{/if}
    <label class="check"><input type="checkbox" checked={skipped} indeterminate={mixedSkipped} onchange={event => { const hidden = event.currentTarget.checked; apply('Skip during presentation', target => setSlideHidden(target, hidden)); }} />{t('Skip during presentation')}</label>
    <label class="check"><input type="checkbox" checked={slideNumbers} onchange={event => editor.setSlideNumbers(event.currentTarget.checked)} />{t('Slide numbers')}</label>
    <button class="ok-btn" onclick={() => editor.runOrPrompt('setSlideNotes')}>{t(slides.length > 1 ? 'Speaker notes (current slide)' : 'Speaker notes')}</button>
    <button class="ok-btn" onclick={() => editor.runOrPrompt('setSlideTransition')}>{t('Slide transition')}</button>
    <button class="ok-btn" onclick={() => editor.runOrPrompt('setSlideSize')}>{t('Page setup')}</button>
    <label>{t('Slide layout')}<select class="ok-input" aria-label={t('Slide layout')} value={mixedLayout ? '__mixed__' : layoutId} onchange={event => changeLayout(event.currentTarget.value)}>
      {#if mixedLayout}<option value="__mixed__" disabled>{t('Mixed')}</option>{/if}
      {#if !layout}<option value="">{t('None')}</option>{/if}
      {#each layouts as item}<option value={getSlideLayoutPartName(item)}>{t(getSlideLayoutName(item))}</option>{/each}
    </select></label>
    <button class="ok-btn" onclick={() => editor.invoke('resetSlideLayout')}>{t('Reset layout')}</button>
    <button class="ok-btn" onclick={() => editor.invoke('addMissingSlidePlaceholders')}>{t('Restore deleted placeholders')}</button>
    <button class="ok-btn" onclick={() => editor.invoke('resetSlidePlaceholderTextFormatting')}>{t('Reset placeholder text formatting')}</button>
    <button class="ok-btn" onclick={() => editor.invoke('resetSlidePlaceholderGeometry')}>{t('Reset placeholder positions')}</button>
    <div class="background-types" role="radiogroup" aria-label={t('Background fill')}>
      <label class="check"><input type="radio" name="background-fill" checked={solidBackground} onchange={() => apply('Background color', target => setSlideBackground(target, '#FFFFFF'))} />{t('Solid fill')}</label>
      <label class="check"><input type="radio" name="background-fill" checked={gradientBackground} onchange={() => apply('Gradient fill', target => setSlideBackgroundGradientFill(target, { stops: [{ offset: 0, color: 'accent1', brightness: 0.95 }, { offset: 1, color: 'accent1', brightness: 0.7 }], angleDeg: 0, scaled: false }))} />{t('Gradient fill')}</label>
      <label class="check"><input type="radio" name="background-fill" checked={patternBackground} onchange={() => apply('Pattern fill', target => setSlideBackgroundPatternFill(target, {}))} />{t('Pattern fill')}</label>
    </div>
    {#if gradientBackground}
      <GradientFillSection background />
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
    <button class="ok-btn" disabled={loading} onclick={() => fileInput?.click()}>{t('Choose background image')}</button>
    <button class="ok-btn" disabled={!canReset} onclick={() => apply('Reset background', target => clearSlideBackground(target))}>{t('Reset background')}</button>
    {#if error}<p role="alert">{t('Slide update failed')}: {error}</p>{/if}
  </section>
{/if}

<style>
  section { display: grid; gap: 10px; padding: 12px; border-bottom: 1px solid var(--ok-border); }
  .transparency { display: grid; grid-template-columns: minmax(0, 1fr) 55px auto; gap: 6px; align-items: center; }
  .transparency input { min-width: 0; width: 100%; }
  .background-types { display: grid; gap: 6px; }
  .selection { font-size: 11px; color: var(--ok-muted); }
  strong { font-size: 12px; }
  label, .color-field { display: grid; gap: 6px; font-size: 11px; }
  .check { display: flex; align-items: center; }
  input[type='file'] { width: 100%; font-size: 11px; }
  [role='alert'] { color: #bf3131; font-size: 11px; }
</style>

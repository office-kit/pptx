<script lang="ts">
  import { getSlides, type SlideData, isSlideHidden, setSlideHidden, getSlideLayout, getSlideLayouts, getSlideLayoutName, getSlideLayoutPartName, setSlideLayout } from '@office-kit/pptx';
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
  const layouts = $derived.by(() => { doc.version; return getSlideLayouts(doc.pres); });
  const layout = $derived.by(() => { doc.version; return slide ? getSlideLayout(slide) : null; });
  const layoutId = $derived(layout ? getSlideLayoutPartName(layout) : '');
  const mixedLayout = $derived(slides.some(item => { const value = getSlideLayout(item); return (value ? getSlideLayoutPartName(value) : '') !== layoutId; }));
  function apply(label: string, operation: (target: SlideData) => void, targets = slides) {
    try { doc.transact(t(label), () => { for (const target of targets) operation(target); }); error = ''; }
    catch (cause) { error = cause instanceof Error ? cause.message : String(cause); }
  }
  function changeLayout(id: string) {
    const target = layouts.find(item => getSlideLayoutPartName(item) === id);
    if (slide && target) apply('Slide layout', item => setSlideLayout(item, target));
  }
  let error = $state('');
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
    <button class="ok-btn" onclick={() => editor.showBackgroundFormat()}>{t('Format Background')}</button>
    {#if error}<p role="alert">{t('Slide update failed')}: {error}</p>{/if}
  </section>
{/if}

<style>
  section { display: grid; gap: 10px; padding: 12px; border-bottom: 1px solid var(--ok-border); }
  .selection { font-size: 11px; color: var(--ok-muted); }
  strong { font-size: 12px; }
  label { display: grid; gap: 6px; font-size: 11px; }
  .check { display: flex; align-items: center; }
  [role='alert'] { color: #bf3131; font-size: 11px; }
</style>

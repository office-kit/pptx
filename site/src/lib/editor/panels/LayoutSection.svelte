<script lang="ts">
  // Editing the layout behind the current slide — PowerPoint's slide master
  // view, Google Slides' theme builder. Everything here is deck-wide by
  // nature: the layout is shared, so the section says so rather than letting
  // the edit look slide-local.
  import {
    getSlideLayout,
    getSlideLayoutBackground,
    getSlideLayoutName,
    getSlideLayoutPartName,
    getSlides,
  } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  const editor = getEditor();
  const doc = editor.doc;
  const slide = $derived.by(() => { doc.version; return doc.currentSlide; });
  const layout = $derived.by(() => { doc.version; return slide ? getSlideLayout(slide) : null; });
  const name = $derived(layout ? getSlideLayoutName(layout) : '');
  const background = $derived.by(() => { doc.version; return layout ? getSlideLayoutBackground(layout) : null; });
  const sharedWith = $derived.by(() => {
    doc.version;
    if (!layout) return 0;
    const partName = getSlideLayoutPartName(layout);
    return getSlides(doc.pres).filter(item => {
      const value = getSlideLayout(item);
      return value !== null && getSlideLayoutPartName(value) === partName;
    }).length;
  });
</script>

{#if layout && (doc.selection.kind === 'none' || doc.selection.kind === 'slide')}
  <section aria-label={t('Layout')}>
    <strong>{t('Layout')}</strong>
    <span class="shared">{t('Slides using this layout')}: {sharedWith}</span>
    <label>{t('Layout name')}
      <input class="ok-input" aria-label={t('Layout name')} value={name} onchange={event => editor.invoke('setSlideLayoutName', { name: event.currentTarget.value })} />
    </label>
    <label>{t('Layout background color')}
      <input type="color" aria-label={t('Layout background color')} value={background?.kind === 'solid' && /^#[0-9a-f]{6}$/i.test(background.color) ? background.color : '#ffffff'} onchange={event => editor.invoke('setSlideLayoutBackground', { color: event.currentTarget.value })} />
    </label>
    <button class="ok-btn" disabled={background?.kind === 'inherit'} onclick={() => editor.invoke('clearSlideLayoutBackground')}>{t('Reset layout background')}</button>
    <button class="ok-btn" onclick={() => editor.runOrPrompt('setSlideLayoutPlaceholderBounds')}>{t('Move a layout placeholder')}</button>
  </section>
{/if}

<style>
  section { display: grid; gap: 10px; padding: 12px; border-bottom: 1px solid var(--ok-border); }
  strong { font-size: 12px; }
  .shared { font-size: 11px; color: var(--ok-muted); }
  label { display: grid; gap: 6px; font-size: 11px; }
  input[type='color'] { width: 100%; height: 26px; }
</style>

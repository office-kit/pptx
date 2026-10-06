<script lang="ts">
  // Mac PowerPoint's Design tab: the Themes gallery, Variants, Colors, Fonts,
  // Background Styles, Layout, Slide Size and Design Suggestions.
  import { placeBelowTrigger } from './place-menu.ts';
  import {
    getPresentationFonts,
    getPresentationTheme,
    getSlideLayout,
    getSlideLayoutName,
    getSlideLayoutPartName,
    getSlideLayouts,
    setPresentationFonts,
    setPresentationTheme,
    setSlideLayout,
    setSlideSize,
    SLIDE_SIZE_16_9,
    SLIDE_SIZE_4_3,
    type SlideSize,
  } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { accentsOf, COLOR_SETS, FONT_PAIRS, THEMES, type ColorSet, type FontPair } from '../core/design-presets.ts';
  import { selectedSlideIndices } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import LayoutThumbnail from '../ui/LayoutThumbnail.svelte';
  import BackgroundStyles from './BackgroundStyles.svelte';

  const editor = getEditor();
  const doc = editor.doc;
  const theme = $derived.by(() => { doc.version; return getPresentationTheme(doc.pres); });
  const fonts = $derived.by(() => { doc.version; return getPresentationFonts(doc.pres); });
  const layouts = $derived.by(() => { doc.version; return getSlideLayouts(doc.pres); });
  const currentLayout = $derived.by(() => {
    doc.version;
    const layout = doc.currentSlide ? getSlideLayout(doc.currentSlide) : null;
    return layout ? getSlideLayoutPartName(layout) : '';
  });
  let open = $state<'colors' | 'fonts' | 'layout' | 'size' | null>(null);
  let page = $state(0);
  const THEMES_PER_PAGE = 6;
  const visibleThemes = $derived(THEMES.slice(page * THEMES_PER_PAGE, (page + 1) * THEMES_PER_PAGE));

  const hex = (value: string | undefined) => (value ?? '').replace('#', '').toUpperCase();
  const sameColors = (colors: ColorSet) => !!theme && accentsOf(colors).every((value, i) => hex(value) === hex(accentsOf(theme)[i]));
  const sameFonts = (pair: FontPair) => !!fonts && fonts.majorLatin === pair.majorLatin && fonts.minorLatin === pair.minorLatin;

  function colors(set: ColorSet) {
    open = null;
    doc.transact(t('Colors'), () => setPresentationTheme(doc.pres, set));
  }
  function fontPair(pair: FontPair) {
    open = null;
    const { name: _name, ...input } = pair;
    doc.transact(t('Fonts'), () => setPresentationFonts(doc.pres, input));
  }
  function applyTheme(index: number) {
    const preset = THEMES[index]!;
    const { name: _name, ...input } = preset.fonts;
    doc.transact(t('Themes'), () => {
      setPresentationTheme(doc.pres, preset.colors);
      setPresentationFonts(doc.pres, input);
    });
  }
  function layout(target: (typeof layouts)[number]) {
    open = null;
    const slides = selectedSlideIndices(doc.selection).map((index) => doc.slideAt(index)).filter((slide) => slide !== null);
    if (slides.length) doc.transact(t('Slide layout'), () => { for (const slide of slides) setSlideLayout(slide, target); });
  }
  function size(value: SlideSize) {
    open = null;
    // PowerPoint asks whether to maximize or ensure fit; Ensure Fit is the
    // choice that never pushes content off the slide.
    doc.transact(t('Slide Size'), () => setSlideSize(doc.pres, value, { content: 'fit' }));
  }
</script>

<svelte:window onpointerdown={(event) => { if (open && !(event.target as Element).closest?.('.design .anchor')) open = null; }} onkeydown={(event) => { if (open && event.key === 'Escape') open = null; }} />

{#snippet trigger(name: NonNullable<typeof open>, icon: string, label: string, enabled = true)}
  <button class="big" aria-label={t(label)} aria-haspopup="menu" aria-expanded={open === name} disabled={!enabled} onclick={() => (open = open === name ? null : name)}><Icon name={icon} size={32} /><span>{t(label)} ⌄</span></button>
{/snippet}

<div class="design">
  <section class="group themes" aria-label={t('Themes')}>
    <div class="gallery" role="listbox" aria-label={t('Themes')}>
      {#each visibleThemes as preset, i (preset.name)}
        {@const selected = sameColors(preset.colors) && sameFonts(preset.fonts)}
        <button class="theme" role="option" aria-selected={selected} aria-label={t(preset.name)} title={t(preset.name)} disabled={!theme} onclick={() => applyTheme(page * THEMES_PER_PAGE + i)}>
          <span class="aa" style:font-family={`'${preset.fonts.majorLatin}', sans-serif`} style:color={preset.colors.dark2}>Aa</span>
          <span class="bar">{#each accentsOf(preset.colors) as color, j (j)}<span style:background={color}></span>{/each}</span>
        </button>
      {/each}
    </div>
    <div class="pager">
      <button aria-label={t('Previous Themes')} disabled={page === 0} onclick={() => (page -= 1)}>⌃</button>
      <button aria-label={t('More Themes')} disabled={(page + 1) * THEMES_PER_PAGE >= THEMES.length} onclick={() => (page += 1)}>⌄</button>
    </div>
  </section>

  <section class="group" aria-label={t('Variants')}>
    <button class="big" aria-label={t('Variants')} title={t('This theme has no variants.')} disabled><Icon name="theme" size={32} /><span>{t('Variants')} ⌄</span></button>
    <div class="anchor">
      {@render trigger('colors', 'theme', 'Colors', !!theme)}
      {#if open === 'colors'}
        <div class="menu" role="menu" use:placeBelowTrigger aria-label={t('Colors')}>
          {#each COLOR_SETS as set (set.name)}
            <button role="menuitemradio" aria-checked={sameColors(set)} onclick={() => colors(set)}><span class="chips">{#each [set.dark2, set.light2, ...accentsOf(set)] as color, j (j)}<span style:background={color}></span>{/each}</span>{t(set.name)}</button>
          {/each}
          <hr />
          <button role="menuitem" disabled={!editor.canRun('setPresentationTheme')} onclick={() => { open = null; editor.runOrPrompt('setPresentationTheme'); }}>{t('Customize Colors...')}</button>
        </div>
      {/if}
    </div>
    <div class="anchor">
      {@render trigger('fonts', 'font', 'Fonts', !!fonts)}
      {#if open === 'fonts'}
        <div class="menu" role="menu" use:placeBelowTrigger aria-label={t('Fonts')}>
          {#each FONT_PAIRS as pair (pair.name)}
            <button role="menuitemradio" aria-checked={sameFonts(pair)} onclick={() => fontPair(pair)}><span class="pair"><strong>{t(pair.name)}</strong><span style:font-family={pair.majorLatin}>{pair.majorLatin}</span><span style:font-family={pair.minorLatin}>{pair.minorLatin}</span></span></button>
          {/each}
          <hr />
          <button role="menuitem" disabled={!editor.canRun('setPresentationFonts')} onclick={() => { open = null; editor.runOrPrompt('setPresentationFonts'); }}>{t('Customize Fonts...')}</button>
        </div>
      {/if}
    </div>
    <BackgroundStyles />
  </section>

  <section class="group" aria-label={t('Customize')}>
    <div class="anchor">
      {@render trigger('layout', 'layout', 'Layout', !!doc.currentSlide)}
      {#if open === 'layout'}
        <div class="menu" role="menu" use:placeBelowTrigger aria-label={t('Layout')}>
          <div class="layout-grid">
            {#each layouts as item (getSlideLayoutPartName(item))}
              <button class="layout-item" role="menuitemradio" aria-checked={getSlideLayoutPartName(item) === currentLayout} onclick={() => layout(item)}><LayoutThumbnail pres={doc.pres} layout={item} /><span>{t(getSlideLayoutName(item))}</span></button>
            {/each}
          </div>
        </div>
      {/if}
    </div>
    <div class="anchor">
      {@render trigger('size', 'resize', 'Slide Size')}
      {#if open === 'size'}
        <div class="menu" role="menu" use:placeBelowTrigger aria-label={t('Slide Size')}>
          <button role="menuitem" onclick={() => size(SLIDE_SIZE_4_3)}>{t('Standard (4:3)')}</button>
          <button role="menuitem" onclick={() => size(SLIDE_SIZE_16_9)}>{t('Widescreen (16:9)')}</button>
          <hr />
          <button role="menuitem" disabled={!editor.canRun('setSlideSize')} onclick={() => { open = null; editor.runOrPrompt('setSlideSize'); }}>{t('Page Setup...')}</button>
        </div>
      {/if}
    </div>
  </section>

  <section class="group" aria-label={t('Designer')}>
    <button class="big" aria-label={t('Design Suggestions')} title={t('Designer needs the Microsoft 365 design service.')} disabled><Icon name="designer" size={32} /><span>{t('Design Suggestions')}</span></button>
  </section>
</div>

<style>
  .design { display: flex; align-items: stretch; min-width: 0; }
  .group { position: relative; display: flex; align-items: center; gap: 2px; padding: 0 6px; border-right: 1px solid var(--ok-border); flex: none; }
  .group:last-child { border-right: none; }
  button { font: inherit; color: var(--ok-text); background: none; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  button:hover:not(:disabled) { background: var(--ok-hover); }
  button:disabled { opacity: 0.4; cursor: default; }
  .big { display: flex; flex-direction: column; align-items: center; gap: 2px; min-width: 48px; padding: 3px; font-size: 11px; line-height: 1.15; }
  .big > span { max-width: 72px; text-align: center; }
  .anchor { position: relative; }
  .gallery { display: flex; gap: 6px; padding: 2px; border: 1px solid var(--ok-border); border-radius: 4px; }
  .theme { display: flex; flex-direction: column; justify-content: space-between; width: 86px; height: 52px; padding: 4px 6px; background: #fff; border: 1px solid var(--ok-border); border-radius: 2px; }
  .theme[aria-selected='true'] { outline: 2px solid var(--ok-accent); outline-offset: 1px; }
  .aa { font-size: 20px; line-height: 1; text-align: left; }
  .bar { display: flex; gap: 2px; }
  .bar span { width: 10px; height: 5px; }
  .pager { display: flex; flex-direction: column; }
  .pager button { padding: 0 3px; font-size: 11px; line-height: 1.2; }
  .menu { position: fixed; z-index: 400; display: flex; flex-direction: column; min-width: 240px; max-height: 70vh; overflow-y: auto; padding: 4px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .menu button { display: flex; align-items: center; gap: 8px; padding: 5px 8px; text-align: left; font-size: 12px; white-space: nowrap; }
  .menu button[aria-checked='true'] { background: var(--ok-selected); }
  .layout-grid { display: grid; grid-template-columns: repeat(3, 120px); gap: 6px; padding: 2px 4px; }
  .menu .layout-item { flex-direction: column; align-items: stretch; gap: 3px; padding: 4px; text-align: center; font-size: 11px; white-space: normal; }
  .layout-item span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .menu hr { width: 100%; border: none; border-top: 1px solid var(--ok-border); margin: 4px 0; }
  .chips { display: inline-flex; gap: 1px; }
  .chips span { width: 10px; height: 14px; border: 1px solid #0002; }
  .pair { display: grid; gap: 1px; }
  .pair span { font-size: 13px; }
</style>

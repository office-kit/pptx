<script lang="ts">
  // Mac PowerPoint 16's Design tab: the Themes gallery, then Variants, Colors,
  // Fonts and Background Styles; Layout and Slide Size; Design Suggestions.
  // The gallery takes the width the other groups leave, in whole 95 pt theme
  // slots, which is all that changes between 1512 and 1200 pt.
  import { tick } from 'svelte';
  import { placeBelowTrigger } from './place-menu.ts';
  import {
    getPresentationFonts,
    getPresentationTheme,
    getSlideLayout,
    getSlideLayoutName,
    getSlideLayoutPartName,
    getSlideLayouts,
    getSlideSize,
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
  import { captionLines } from './caption.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const theme = $derived.by(() => { doc.version; return getPresentationTheme(doc.pres); });
  const fonts = $derived.by(() => { doc.version; return getPresentationFonts(doc.pres); });
  const layouts = $derived.by(() => { doc.version; return getSlideLayouts(doc.pres); });
  const slideSize = $derived.by(() => { doc.version; return getSlideSize(doc.pres); });
  const currentLayout = $derived.by(() => {
    doc.version;
    const layout = doc.currentSlide ? getSlideLayout(doc.currentSlide) : null;
    return layout ? getSlideLayoutPartName(layout) : '';
  });
  type Menu = 'colors' | 'fonts' | 'layout' | 'size';
  let open = $state<Menu | null>(null);
  let root = $state<HTMLDivElement>();

  // Mac PowerPoint's gallery: an 18 pt Previous and Next button either side
  // of whole 95 pt slots (85 pt thumbnails, 10 pt apart).
  const SLOT = 95;
  const PAGER = 18;
  let width = $state(0);
  let restWidth = $state(0);
  // The Themes group's own padding (3 + 9) and its rule.
  const GROUP_CHROME = 13;
  const slots = $derived(Math.max(1, Math.floor((width - restWidth - GROUP_CHROME - 2 * PAGER) / SLOT)));
  let first = $state(0);
  const visibleThemes = $derived(THEMES.slice(first, first + slots));
  $effect(() => {
    if (first > 0 && first >= THEMES.length) first = Math.max(0, THEMES.length - slots);
  });

  const hex = (value: string | undefined) => (value ?? '').replace('#', '').toUpperCase();
  const sameColors = (colors: ColorSet) => !!theme && accentsOf(colors).every((value, i) => hex(value) === hex(accentsOf(theme)[i]));
  const sameFonts = (pair: FontPair) => !!fonts && fonts.majorLatin === pair.majorLatin && fonts.minorLatin === pair.minorLatin;
  const currentFontName = $derived(FONT_PAIRS.find(sameFonts)?.name ?? theme?.name ?? '');
  const sameSize = (size: SlideSize) => !!slideSize && slideSize.width === size.width && slideSize.height === size.height;

  async function toggle(menu: Menu) {
    open = open === menu ? null : menu;
    if (open) {
      await tick();
      root?.querySelector<HTMLElement>('.menu [aria-checked="true"], .menu button:not(:disabled)')?.focus();
    }
  }
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

<svelte:window onpointerdown={(event) => { if (open && !(event.target as Element).closest?.('.design .anchor')) open = null; }} onkeydown={(event) => { if (open && event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); open = null; } }} />

<!-- PowerPoint's ▾ buttons here put the arrow beside a 32 pt icon. -->
{#snippet trigger(name: Menu, icon: string, label: string, enabled = true)}
  <button class="big" aria-label={t(label)} aria-haspopup="menu" aria-expanded={open === name} disabled={!enabled} onclick={() => toggle(name)}><span class="icon-row"><Icon name={icon} size={32} /><span class="arrow" aria-hidden="true">⌄</span></span><span class="caption">{captionLines(t(label))}</span></button>
{/snippet}

{#snippet chips(set: Pick<ColorSet, 'dark2' | 'light2' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6'>)}
  <span class="chips" aria-hidden="true">{#each [set.dark2, set.light2, ...accentsOf(set)] as color, j (j)}<span style:background={color}></span>{/each}</span>
{/snippet}

{#snippet fontSample(name: string, major: string, minor: string)}
  <span class="aa-box" aria-hidden="true" style:font-family={`'${major}', sans-serif`}>Aa</span>
  <span class="pair" aria-hidden="true"><span class="name">{t(name)}</span><span style:font-family={`'${major}', sans-serif`}>{major}</span><span class="minor" style:font-family={`'${minor}', sans-serif`}>{minor}</span></span>
{/snippet}

<div class="design" bind:this={root} bind:clientWidth={width}>
  <section class="cluster themes" aria-label={t('Themes')}>
    <div class="gallery">
      <button class="pager" aria-label={t('Previous Themes')} disabled={first === 0} onclick={() => (first = Math.max(0, first - slots))}>‹</button>
      <div class="slots" role="listbox" aria-label={t('Themes')} style:width="{slots * SLOT}px">
        {#each visibleThemes as preset, i (preset.name)}
          {@const selected = sameColors(preset.colors) && sameFonts(preset.fonts)}
          <button class="theme" role="option" aria-selected={selected} aria-label={t(preset.name)} title={t(preset.name)} disabled={!theme} onclick={() => applyTheme(first + i)}>
            <span class="aa" style:font-family={`'${preset.fonts.majorLatin}', sans-serif`} style:color={preset.colors.dark2}>Aa</span>
            <span class="bar">{#each accentsOf(preset.colors) as color, j (j)}<span style:background={color}></span>{/each}</span>
          </button>
        {/each}
      </div>
      <button class="pager" aria-label={t('More Themes')} disabled={first + slots >= THEMES.length} onclick={() => (first += slots)}>›</button>
    </div>
  </section>

  <div class="rest" bind:offsetWidth={restWidth}>
    <section class="cluster" aria-label={t('Variants')}>
      <button class="big" aria-label={t('Variants')} aria-haspopup="menu" title={t('This theme has no variants.')} disabled><span class="icon-row"><Icon name="theme" size={32} /><span class="arrow" aria-hidden="true">⌄</span></span><span class="caption">{t('Variants')}</span></button>
      <div class="anchor">
        {@render trigger('colors', 'theme', 'Colors', !!theme)}
        {#if open === 'colors'}
          <div class="menu colors" role="menu" use:placeBelowTrigger aria-label={t('Colors')}>
            <div class="list">
              {#if theme}
                <div role="group" aria-label={t('Theme Colors')}>
                  <div class="heading">{t('Theme Colors')}</div>
                  <button class="color-item" role="menuitemradio" aria-label={t(theme.name)} aria-checked="true" onclick={() => (open = null)}>{@render chips(theme)}<span>{t(theme.name)}</span></button>
                </div>
              {/if}
              <div role="group" aria-label={t('All Colors')}>
                <div class="heading">{t('All Colors')}</div>
                {#each COLOR_SETS as set (set.name)}
                  <button class="color-item" role="menuitemradio" aria-label={t(set.name)} aria-checked={sameColors(set)} onclick={() => colors(set)}>{@render chips(set)}<span>{t(set.name)}</span></button>
                {/each}
              </div>
            </div>
            <hr />
            <button class="action" role="menuitem" disabled={!editor.canRun('setPresentationTheme')} onclick={() => { open = null; editor.runOrPrompt('setPresentationTheme'); }}>{t('Customize Colors...')}</button>
            <button class="action" role="menuitem" disabled title={t('The library does not write per-slide theme colors yet.')}>{t('Reset Slide Theme Colors')}</button>
          </div>
        {/if}
      </div>
      <div class="anchor">
        {@render trigger('fonts', 'font', 'Fonts', !!fonts)}
        {#if open === 'fonts'}
          <div class="menu fonts" role="menu" use:placeBelowTrigger aria-label={t('Fonts')}>
            <div class="list">
              {#if fonts?.majorLatin && fonts.minorLatin}
                <div role="group" aria-label={t('Theme Fonts')}>
                  <div class="heading">{t('Theme Fonts')}</div>
                  <button class="font-item" role="menuitemradio" aria-label={t(currentFontName)} aria-checked="true" onclick={() => (open = null)}>{@render fontSample(currentFontName, fonts.majorLatin, fonts.minorLatin)}</button>
                </div>
              {/if}
              <div role="group" aria-label={t('All Fonts')}>
                <div class="heading">{t('All Fonts')}</div>
                {#each FONT_PAIRS as pair (pair.name)}
                  <button class="font-item" role="menuitemradio" aria-label={t(pair.name)} aria-checked={sameFonts(pair)} onclick={() => fontPair(pair)}>{@render fontSample(pair.name, pair.majorLatin, pair.minorLatin)}</button>
                {/each}
              </div>
            </div>
            <!-- Mac PowerPoint's Fonts menu ends with the list; the editor keeps
                 its theme-font dialog here. -->
            <hr />
            <button class="action" role="menuitem" disabled={!editor.canRun('setPresentationFonts')} onclick={() => { open = null; editor.runOrPrompt('setPresentationFonts'); }}>{t('Customize Fonts...')}</button>
          </div>
        {/if}
      </div>
      <BackgroundStyles />
    </section>

    <section class="cluster" aria-label={t('Customize')}>
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
            <button class="action" role="menuitemradio" aria-checked={sameSize(SLIDE_SIZE_4_3)} onclick={() => size(SLIDE_SIZE_4_3)}>{t('Standard (4:3)')}</button>
            <button class="action" role="menuitemradio" aria-checked={sameSize(SLIDE_SIZE_16_9)} onclick={() => size(SLIDE_SIZE_16_9)}>{t('Widescreen (16:9)')}</button>
            <hr />
            <button class="action" role="menuitem" disabled={!editor.canRun('setSlideSize')} onclick={() => { open = null; editor.runOrPrompt('setSlideSize'); }}>{t('Page Setup...')}</button>
          </div>
        {/if}
      </div>
    </section>

    <section class="cluster" aria-label={t('Designer')}>
      <button class="big suggestions" aria-label={t('Design Suggestions')} aria-pressed="false" title={t('Designer needs the Microsoft 365 design service.')} disabled><Icon name="designer" size={32} /><span class="caption">{captionLines(t('Design Suggestions'))}</span></button>
    </section>
  </div>
</div>

<style>
  /* The Home tab's metrics (see HomeRibbon.svelte). */
  .design { display: flex; align-items: stretch; min-width: 0; width: 100%; height: 72px; }
  .rest { display: flex; align-items: stretch; flex: none; }
  .cluster { position: relative; display: flex; align-items: center; padding: 0 9px; border-right: 1px solid var(--ok-border); flex: none; }
  .design > .cluster:first-child { padding-left: 3px; }
  .rest .cluster:last-child { border-right: none; }
  button { font: inherit; color: var(--ok-text); background: none; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  button:hover:not(:disabled) { background: var(--ok-hover); }
  button:disabled { opacity: 0.4; cursor: default; }
  .big { display: flex; flex-direction: column; align-items: center; justify-content: flex-start; align-self: stretch; gap: 2px; min-width: 50px; padding: 4px 1px; font-size: 11px; line-height: 1.15; }
  .caption { white-space: pre-line; text-align: center; }
  .icon-row { display: flex; align-items: center; gap: 1px; }
  .arrow { font-size: 10px; }
  .anchor { position: relative; display: flex; align-self: stretch; }
  .gallery { display: flex; align-items: center; align-self: flex-start; height: 60px; border: 1px solid var(--ok-border); border-radius: 5px; background: var(--ok-panel); }
  .pager { align-self: stretch; width: 18px; padding: 0; font-size: 14px; border-radius: 4px; }
  .pager:disabled { visibility: hidden; }
  .slots { display: flex; align-items: center; gap: 10px; padding-left: 5px; box-sizing: border-box; overflow: hidden; }
  .theme { display: flex; flex-direction: column; justify-content: space-between; flex: none; width: 85px; height: 48px; padding: 4px 6px; background: #fff; border: 1px solid var(--ok-border); border-radius: 2px; }
  .theme[aria-selected='true'] { outline: 2px solid var(--ok-accent); outline-offset: 1px; }
  .aa { font-size: 20px; line-height: 1; text-align: left; }
  .bar { display: flex; gap: 2px; }
  .bar span { width: 9px; height: 5px; }
  .suggestions { min-width: 69px; }
  .menu { position: fixed; z-index: 400; display: flex; flex-direction: column; min-width: 150px; max-height: 70vh; padding: 4px 0; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .menu .list { overflow-y: auto; min-height: 0; }
  .menu.colors { width: 259px; }
  .menu.fonts { width: 237px; }
  .menu .heading { padding: 3px 4px; font-size: 12px; color: var(--ok-text-2); border-bottom: 1px solid var(--ok-border); background: var(--ok-hover); }
  .action { display: flex; align-items: center; width: 100%; height: 24px; padding: 0 20px; text-align: left; font-size: 12px; white-space: nowrap; }
  .action[aria-checked='true']::before { content: '✓'; margin-left: -14px; margin-right: 4px; }
  .color-item { display: flex; align-items: center; gap: 6px; width: calc(100% - 4px); height: 27px; margin: 0 2px; padding: 0 4px; font-size: 13px; text-align: left; white-space: nowrap; }
  .font-item { display: flex; align-items: center; gap: 8px; width: calc(100% - 4px); height: 62px; margin: 0 2px; padding: 0 6px; text-align: left; }
  .color-item[aria-checked='true'], .font-item[aria-checked='true'] { background: var(--ok-selected); border-color: var(--ok-selected-border); }
  .chips { display: inline-flex; flex: none; }
  .chips span { width: 13px; height: 14px; border: 1px solid #0002; margin-right: -1px; }
  .aa-box { display: flex; align-items: flex-end; flex: none; width: 42px; height: 42px; padding: 0 3px; box-sizing: border-box; font-size: 24px; line-height: 1.2; border: 1px solid var(--ok-border-strong); background: #fff; color: #444; }
  .pair { display: grid; gap: 1px; min-width: 0; font-size: 13px; }
  .pair span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .pair .name { font-size: 11px; color: var(--ok-text-2); }
  .pair .minor { font-size: 11px; }
  .layout-grid { display: grid; grid-template-columns: repeat(5, 94px); gap: 4px; padding: 2px 6px; }
  .layout-item { display: flex; flex-direction: column; align-items: stretch; gap: 3px; padding: 3px; text-align: center; font-size: 11px; }
  .layout-item[aria-checked='true'] { background: var(--ok-selected); border-color: var(--ok-selected-border); }
  .layout-item span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .menu hr { width: 100%; border: none; border-top: 1px solid var(--ok-border); margin: 5px 0; }
</style>

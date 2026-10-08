<script lang="ts">
  // The Home tab laid out like the reference desktop app (Mac, 16): clipboard, slides, font,
  // paragraph, insert and drawing clusters without group captions. As the
  // ribbon narrows, groups collapse into single buttons in the reference desktop app's order
  // (Drawing first; then Slides, Paragraph and Insert; Font last). The
  // clipboard cluster never collapses.
  import { placeBelowTrigger } from './place-menu.ts';
  import { tick, type Snippet } from 'svelte';
  import { eventTarget } from '../core/dom-root.ts';
  import {
    getShapeKind,
    getSlideLayout,
    getSlideLayoutName,
    getSlideLayoutPartName,
    getSlideLayouts,
    setParagraphBullet,
    setParagraphLevel,
    setSlideLayout,
    type BulletStyle,
    type Color,
    type SlideLayoutData,
  } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { selectedShapeIds, selectedSlideIndices } from '../core/selection.ts';
  import { editTargetParagraphs, targetParagraphProperties } from '../core/paragraph-targets.ts';
  import { getLocale, t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import ParagraphLayoutMenus from './ParagraphLayoutMenus.svelte';
  import SectionMenu from './SectionMenu.svelte';
  import LayoutThumbnail from '../ui/LayoutThumbnail.svelte';
  import ColorPicker from '../ui/ColorPicker.svelte';
  import ShapeFillPicker from './ShapeFillPicker.svelte';
  import FontRibbon from './FontRibbon.svelte';
  import ParagraphAlignment from './ParagraphAlignment.svelte';
  import LineSpacingMenu from './LineSpacingMenu.svelte';
  import ParagraphDialog from '../ui/ParagraphDialog.svelte';
  import ArrangeMenu from './ArrangeMenu.svelte';
  import ShapeQuickStyles from './ShapeQuickStyles.svelte';
  import { PRESET } from './config.ts';

  const editor = getEditor();
  const doc = editor.doc;

  // Ribbon widths (CSS px) below which each collapse step applies. Measured
  // from the reference desktop app's (Mac) window at 1512 (all expanded), 1200 (Drawing
  // collapsed, Slides/Insert in small icons), 1000 (Slides, Paragraph, Insert
  // collapsed) and 800 (Font collapsed as well).
  const SMALL_ICONS_BELOW = 1300;
  const GROUPS_COLLAPSE_BELOW = 1080;
  const FONT_COLLAPSES_BELOW = 840;
  type Group = 'Slides' | 'Font' | 'Paragraph' | 'Insert' | 'Drawing';

  // The ribbon spans the window, so start from its width: measuring only after
  // mount would paint the expanded layout first and shift the canvas below.
  let width = $state(typeof window === 'undefined' ? Number.POSITIVE_INFINITY : window.innerWidth);
  // The steps above, in order; 0 is fully expanded.
  const stepAt = (value: number) =>
    value < FONT_COLLAPSES_BELOW ? 3 : value < GROUPS_COLLAPSE_BELOW ? 2 : value < SMALL_ICONS_BELOW ? 1 : 0;
  // The thresholds fit the reference desktop app's English labels. Longer labels (Japanese,
  // or a platform with wider fonts) can still overflow near a threshold, so
  // the ribbon takes the next step while its content does not fit.
  let extraSteps = $state(0);
  let home = $state<HTMLDivElement>();
  const step = $derived(Math.min(3, stepAt(width) + extraSteps));
  const small = $derived(step >= 1);
  const collapsed = $derived<ReadonlySet<Group>>(
    new Set<Group>([
      ...(step >= 1 ? (['Drawing'] as const) : []),
      ...(step >= 2 ? (['Slides', 'Paragraph', 'Insert'] as const) : []),
      ...(step >= 3 ? (['Font'] as const) : []),
    ]),
  );
  $effect.pre(() => {
    width;
    getLocale();
    extraSteps = 0;
  });
  $effect(() => {
    step;
    width;
    getLocale();
    if (home && step < 3 && home.scrollWidth > home.clientWidth + 1) extraSteps += 1;
  });
  let openGroup = $state<Group | null>(null);
  let paragraphOptions = $state(false);
  let openMenu = $state<'newSlide' | 'layout' | 'paste' | null>(null);
  let popup = $state<HTMLDivElement>();

  const shapeIds = $derived(selectedShapeIds(doc.selection));
  const canCopy = $derived(doc.selection.kind === 'slide' || shapeIds.length > 0);
  const hasSlide = $derived(doc.currentSlide !== null);
  const layouts = $derived.by(() => { doc.version; return getSlideLayouts(doc.pres); });
  const currentLayout = $derived.by(() => {
    doc.version;
    const slide = doc.currentSlide;
    const layout = slide ? getSlideLayout(slide) : null;
    return layout ? getSlideLayoutPartName(layout) : '';
  });

  const paragraphs = $derived.by(() => { doc.version; return targetParagraphProperties(editor); });
  const paragraphEnabled = $derived(paragraphs.length > 0 && !editor.selectionLocked());
  $effect(() => { if (!paragraphEnabled) paragraphOptions = false; });
  const isBulleted = (bullet: BulletStyle | null) => bullet === 'bullet' || (typeof bullet === 'object' && bullet !== null && 'char' in bullet);
  const isNumbered = (bullet: BulletStyle | null) => bullet === 'number' || (typeof bullet === 'object' && bullet !== null && 'autoNum' in bullet);
  const bulleted = $derived(paragraphEnabled && paragraphs.every(p => isBulleted(p.bullet)));
  const numbered = $derived(paragraphEnabled && paragraphs.every(p => isNumbered(p.bullet)));
  // OOXML outline levels run 0..8 (ST_TextIndentLevelType).
  const MAX_LEVEL = 8;

  const paintable = $derived.by(() => {
    doc.version;
    const shapes = editor.selectedShapes();
    return shapes.length > 0 && !editor.selectionLocked() && shapes.every(shape => ['shape', 'connector'].includes(getShapeKind(shape)));
  });

  function toggleList(kind: 'bullet' | 'number') {
    const on = kind === 'bullet' ? bulleted : numbered;
    editTargetParagraphs(editor, (shape, index) => setParagraphBullet(shape, index, on ? 'none' : kind));
  }
  function changeLevel(delta: 1 | -1) {
    const levels = paragraphs.map(p => p.level);
    editTargetParagraphs(editor, (shape, index) => {
      const level = levels[index] ?? 0;
      setParagraphLevel(shape, index, Math.max(0, Math.min(MAX_LEVEL, level + delta)));
    });
  }
  function applyLayout(layout: SlideLayoutData) {
    openMenu = null;
    const slides = selectedSlideIndices(doc.selection).map(index => doc.slideAt(index)).filter(slide => slide !== null);
    if (!slides.length) return;
    doc.transact(t('Slide layout'), () => { for (const slide of slides) setSlideLayout(slide, layout); });
  }
  function insertSlide(layout: SlideLayoutData) {
    openMenu = null;
    editor.invoke('addSlide', { options: { layout } });
  }
  function outline(color: Color) { editor.invoke('setShapeStroke', { options: { color } }); }

  async function toggleGroup(group: Group) {
    openMenu = null;
    openGroup = openGroup === group ? null : group;
    if (openGroup) { await tick(); popup?.querySelector<HTMLElement>('button:not(:disabled), input')?.focus(); }
  }
  async function toggleMenu(menu: 'newSlide' | 'layout' | 'paste') {
    openMenu = openMenu === menu ? null : menu;
    if (openMenu) { await tick(); editor.shell?.querySelector<HTMLElement>('.home-menu [aria-checked="true"], .home-menu button')?.focus(); }
  }
  // Keep Text Only: the system clipboard's plain text, at the text cursor or
  // as a new text box.
  async function pasteTextOnly() {
    let text = '';
    try {
      text = await navigator.clipboard.readText();
    } catch (cause) {
      editor.toast('error', `${t('Paste failed')}: ${(cause as Error).message}`);
      return;
    }
    if (!text) return;
    if (editor.inlineTextFormat?.insertText) editor.inlineTextFormat.insertText(text);
    else editor.runOrPrompt('addSlideTextBox', { opts: { ...PRESET.textBox.opts, text } });
  }
  function closeAfterCommand(event: MouseEvent) {
    const button = (event.target as Element).closest('button');
    if (!button || button.matches('[aria-haspopup], .inline-more')) return;
    // Keep focus inside the ribbon: if the removed button's focus fell to the
    // page, in-progress text editing on the canvas would be committed.
    ((event.currentTarget as HTMLElement).previousElementSibling as HTMLElement | null)?.focus();
    openGroup = null;
  }
  function dismiss(event: PointerEvent) {
    const target = eventTarget(event);
    if (!(target instanceof Element)) return;
    if (openGroup && !target.closest('.group-popup, .group-trigger, .palette, .menu')) openGroup = null;
    if (openMenu && !target.closest('.home-menu, .menu-trigger')) openMenu = null;
  }
  function keydown(event: KeyboardEvent) {
    if (event.key !== 'Escape' || (!openGroup && !openMenu)) return;
    event.preventDefault();
    event.stopPropagation();
    if (openMenu) openMenu = null;
    else openGroup = null;
  }
</script>

<svelte:window onpointerdown={dismiss} onkeydown={keydown} />

{#snippet layoutMenu(kind: 'newSlide' | 'layout')}
  <div class="home-menu" role="menu" tabindex="-1" aria-label={t(kind === 'newSlide' ? 'New Slide' : 'Layout')} use:placeBelowTrigger>
    <div class="heading">{t('Layouts')}</div>
    <!-- The reference desktop app shows the layouts as a gallery of thumbnails. -->
    <div class="layout-grid">
      {#each layouts as layout (getSlideLayoutPartName(layout))}
        {#if kind === 'newSlide'}
          <button class="layout-item" role="menuitem" onclick={() => insertSlide(layout)}><LayoutThumbnail pres={doc.pres} {layout} /><span>{t(getSlideLayoutName(layout))}</span></button>
        {:else}
          <button class="layout-item" role="menuitemradio" aria-checked={getSlideLayoutPartName(layout) === currentLayout} onclick={() => applyLayout(layout)}><LayoutThumbnail pres={doc.pres} {layout} /><span>{t(getSlideLayoutName(layout))}</span></button>
        {/if}
      {/each}
    </div>
    {#if kind === 'newSlide'}
      <hr />
      <button role="menuitem" disabled={!hasSlide} onclick={() => { openMenu = null; editor.invoke('duplicateSlide'); }}>{t('Duplicate Selected Slides')}</button>
    {/if}
  </div>
{/snippet}

{#snippet slides()}
  <div class="split large">
    <button class="big" aria-label={t('New Slide')} onclick={() => editor.addNewSlide()}><Icon name="new-slide" size={32} /><span>{t('New Slide')}</span></button>
    <button class="arrow menu-trigger" aria-label={t('New Slide options')} aria-haspopup="menu" aria-expanded={openMenu === 'newSlide'} onclick={() => toggleMenu('newSlide')}>⌄</button>
    {#if openMenu === 'newSlide'}{@render layoutMenu('newSlide')}{/if}
  </div>
  <div class="stack" class:small>
    <div class="anchor">
      <button class="menu-trigger" class:big={!small} class:row={small} disabled={!hasSlide} aria-label={t('Layout')} aria-haspopup="menu" aria-expanded={openMenu === 'layout'} onclick={() => toggleMenu('layout')}>{#if small}<Icon name="layout" size={18} /><span>{t('Layout')}</span><span aria-hidden="true">⌄</span>{:else}<span class="icon-row"><Icon name="layout" size={32} /><span aria-hidden="true">⌄</span></span><span>{t('Layout')}</span>{/if}</button>
      {#if openMenu === 'layout'}{@render layoutMenu('layout')}{/if}
    </div>
    <button class:big={!small} class:narrow={!small} class:row={small} disabled={!editor.canRun('resetSlideLayout')} aria-label={t('Reset')} title={t('Reset the position, size, and formatting of the slide placeholders to their default settings.')} onclick={() => editor.invoke('resetSlideLayout')}><Icon name="reset" size={small ? 18 : 32} /><span>{t('Reset')}</span></button>
    <SectionMenu {small} />
  </div>
{/snippet}

{#snippet paragraph()}
  <div class="rows">
    <div class="row-controls">
      <button class="tool" aria-label={t('Bullets')} aria-pressed={bulleted} disabled={!paragraphEnabled} onclick={() => toggleList('bullet')}><Icon name="bullets" size={18} /></button>
      <button class="tool" aria-label={t('Numbering')} aria-pressed={numbered} disabled={!paragraphEnabled} onclick={() => toggleList('number')}><Icon name="numbering" size={18} /></button>
      <span class="sep" aria-hidden="true"></span>
      <button class="tool" aria-label={t('Decrease List Level')} disabled={!paragraphEnabled} onclick={() => changeLevel(-1)}><Icon name="indent-less" size={18} /></button>
      <button class="tool" aria-label={t('Increase List Level')} disabled={!paragraphEnabled} onclick={() => changeLevel(1)}><Icon name="indent-more" size={18} /></button>
      <span class="sep" aria-hidden="true"></span>
      <LineSpacingMenu onoptions={() => (paragraphOptions = true)} />
      <span class="sep" aria-hidden="true"></span>
      <ParagraphLayoutMenus row={1} />
    </div>
    <div class="row-controls"><ParagraphAlignment /><span class="sep" aria-hidden="true"></span><ParagraphLayoutMenus row={2} /></div>
  </div>
{/snippet}

{#snippet insert()}
  <button class="big" aria-label={t('Picture')} disabled={!editor.canRun('addSlideImage')} onclick={() => editor.runOrPrompt('addSlideImage')}><Icon name="picture" size={32} /><span>{t('Picture')}</span></button>
  <div class="stack" class:small>
    <button class:big={!small} class:row={small} aria-label={t('Shapes')} aria-haspopup="menu" aria-expanded={!!editor.shapeGallery} disabled={!editor.canRun('addSlideShape')} onclick={(event) => editor.openShapeGallery(event.currentTarget)}><Icon name="shapes" size={small ? 18 : 32} /><span>{t('Shapes')}</span></button>
    <button class:big={!small} class:row={small} aria-label={t('Text Box')} disabled={!editor.canRun('addSlideTextBox')} onclick={() => editor.runOrPrompt('addSlideTextBox', PRESET.textBox)}><Icon name="textbox" size={small ? 18 : 32} /><span>{t('Text Box')}</span></button>
  </div>
{/snippet}

{#snippet drawing()}
  <ArrangeMenu />
  <ShapeQuickStyles />
  <div class="stack small fill-outline">
    <span class="paint-row"><Icon name="fill" size={18} /><span class="label">{t('Shape Fill')}</span><ShapeFillPicker disabled={!paintable} /></span>
    <span class="paint-row"><Icon name="outline" size={18} /><span class="label">{t('Shape Outline')}</span><ColorPicker compact label={t('Shape Outline')} disabled={!paintable} choose={outline} /></span>
  </div>
{/snippet}

{#snippet group(name: Group, icon: string, body: Snippet)}
  <section class="cluster" aria-label={t(name)}>
    {#if collapsed.has(name)}
      <button class="big group-trigger" aria-haspopup="dialog" aria-expanded={openGroup === name} onclick={() => toggleGroup(name)}><Icon name={icon} size={32} /><span>{t(name)} <span aria-hidden="true">⌄</span></span></button>
      {#if openGroup === name}
        <!-- Choosing a command closes the popup, as the reference desktop app's collapsed groups do. -->
        <!-- svelte-ignore a11y_click_events_have_key_events -->
        <div class="group-popup" role="dialog" tabindex="-1" aria-label={t(name)} bind:this={popup} use:placeBelowTrigger onclick={closeAfterCommand}>{@render body()}</div>
      {/if}
    {:else}
      {@render body()}
    {/if}
  </section>
{/snippet}

{#snippet font()}<FontRibbon />{/snippet}

<div class="home" bind:this={home} bind:clientWidth={width}>
  <section class="cluster" aria-label={t('Clipboard')}>
    <div class="split large">
      <button class="big" aria-label={t('Paste')} onclick={() => editor.paste()}><Icon name="paste" size={32} /><span>{t('Paste')}</span></button>
      <button class="arrow menu-trigger" aria-label={t('Paste options')} aria-haspopup="menu" aria-expanded={openMenu === 'paste'} onclick={() => toggleMenu('paste')}>⌄</button>
      {#if openMenu === 'paste'}
        <div class="home-menu" role="menu" tabindex="-1" aria-label={t('Paste options')} use:placeBelowTrigger>
          <button role="menuitem" onclick={() => { openMenu = null; void editor.paste(); }}>{t('Paste')}</button>
          <button role="menuitem" onclick={() => { openMenu = null; void pasteTextOnly(); }}>{t('Keep Text Only')}</button>
          <hr />
          <button role="menuitem" title={t('Paste Special needs the system clipboard formats, which the browser does not expose.')} disabled>{t('Paste Special...')}</button>
        </div>
      {/if}
    </div>
    <div class="stack small tools">
      <button class="tool" aria-label={t('Cut')} title={t('Cut')} disabled={!canCopy} onclick={() => editor.cutSelection()}><Icon name="cut" size={18} /></button>
      <button class="tool" aria-label={t('Copy')} title={t('Copy')} disabled={!canCopy} onclick={() => editor.copySelection()}><Icon name="copy" size={18} /></button>
      <button class="tool" aria-label={t('Format Painter')} title={t('Format Painter')} aria-pressed={editor.formatPainterSource !== null} disabled={!shapeIds.length && !editor.formatPainterSource} onclick={() => editor.toggleFormatPainter()}><Icon name="format-painter" size={18} /></button>
    </div>
  </section>
  {@render group('Slides', 'layout', slides)}
  {@render group('Font', 'font', font)}
  {@render group('Paragraph', 'align', paragraph)}
  {@render group('Insert', 'textbox', insert)}
  {@render group('Drawing', 'quick-styles', drawing)}
  <section class="cluster" aria-label={t('Add-ins')}>
    <button class={small ? 'tool' : 'big narrow'} aria-label={t('Add-ins')} title={t('Add-ins are not available in this editor.')} disabled><Icon name="add-ins" size={small ? 18 : 32} />{#if !small}<span>{t('Add-ins')}</span>{/if}</button>
  </section>
  <!-- The reference desktop app separates Add-ins and Designer into two groups. -->
  <section class="cluster" aria-label={t('Designer')}>
    <button class={small ? 'tool' : 'big'} aria-label={t('Designer')} title={t('Designer needs an online design service.')} disabled><Icon name="designer" size={small ? 18 : 32} />{#if !small}<span>{t('Designer')}</span>{/if}</button>
  </section>
</div>
{#if paragraphOptions}
  <ParagraphDialog properties={paragraphs} apply={edit => editTargetParagraphs(editor, edit)} onclose={() => {
    paragraphOptions = false;
    (editor.shell?.querySelector<HTMLElement>('.home [aria-label="' + t('Line spacing') + '"]') ?? editor.shell?.querySelector<HTMLElement>('.home .group-trigger[aria-expanded]'))?.focus();
  }} />
{/if}

<style>
  /* The reference desktop app's (Mac) Home commands: a 72 pt row of groups separated by a
     rule with 10 pt on either side; large buttons are 50 pt wide and fill the
     row, small ones 26 × 26 pt, and menu buttons with an arrow 38 × 26 pt. */
  .home { display: flex; align-items: stretch; min-width: 0; width: 100%; height: 72px; gap: 0; }
  .cluster { position: relative; display: flex; align-items: center; gap: 0; padding: 0 10px; border-right: 1px solid var(--ok-border); flex: none; }
  .cluster:first-child { padding-left: 4px; }
  .cluster:last-child { border-right: none; }
  button { font: inherit; color: var(--ok-text); background: none; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  button:hover:not(:disabled) { background: var(--ok-hover); }
  button:disabled { opacity: 0.4; cursor: default; }
  button[aria-pressed='true'] { background: var(--ok-selected); border-color: var(--ok-selected-border); }
  .big { display: flex; flex-direction: column; align-items: center; justify-content: flex-start; align-self: stretch; gap: 2px; min-width: 50px; padding: 4px 2px; font-size: 11px; line-height: 1.15; }
  .big > span { max-width: 52px; text-align: center; }
  .split > .big { min-width: 36px; }
  .split > .big > span { max-width: 36px; }
  .big.narrow { min-width: 38px; }
  .icon-row { display: flex; align-items: center; gap: 2px; max-width: none !important; }
  .row { display: flex; align-items: center; gap: 3px; padding: 2px 3px; font-size: 11px; white-space: nowrap; }
  .tool { display: flex; align-items: center; justify-content: center; width: 26px; height: 26px; padding: 0; }
  .stack { display: flex; align-items: center; gap: 0; }
  .stack:not(.small) { align-self: stretch; align-items: stretch; }
  .stack:not(.small) > .anchor, .stack:not(.small) > :global(.section-menu) { display: flex; }
  .stack.small { flex-direction: column; align-items: flex-start; gap: 2px; }
  .stack.tools { align-items: center; }
  .split { position: relative; display: flex; align-items: flex-start; }
  .split.large { align-self: stretch; }
  .arrow { align-self: center; padding: 2px; font-size: 12px; }
  .anchor { position: relative; }
  .rows { display: flex; flex-direction: column; align-self: flex-start; gap: 6px; margin-top: 4px; }
  .row-controls { display: flex; align-items: center; gap: 0; height: 26px; }
  .sep { width: 1px; height: 20px; margin: 0 7px; background: var(--ok-border); }
  .paint-row { display: flex; align-items: center; gap: 5px; font-size: 11px; white-space: nowrap; }
  .home-menu, .group-popup { position: fixed; z-index: 400; padding: 6px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .group-popup { display: flex; align-items: center; gap: 4px; }
  .home-menu { display: flex; flex-direction: column; min-width: 220px; max-height: 70vh; overflow-y: auto; }
  .home-menu button { display: flex; align-items: center; gap: 8px; padding: 5px 8px; text-align: left; font-size: 12px; }
  .home-menu button[aria-checked='true'] { background: var(--ok-selected); }
  .layout-grid { display: grid; grid-template-columns: repeat(3, 120px); gap: 6px; padding: 2px 4px; }
  .home-menu .layout-item { flex-direction: column; align-items: stretch; gap: 3px; padding: 4px; text-align: center; font-size: 11px; }
  .layout-item span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .home-menu .layout-item[aria-checked='true'] { background: var(--ok-selected); }
  .home-menu .heading { padding: 4px 8px; font-size: 11px; font-weight: 600; color: var(--ok-text-2); }
  .home-menu hr { width: 100%; border: none; border-top: 1px solid var(--ok-border); margin: 4px 0; }
</style>

<script lang="ts">
  // Mac PowerPoint 16's Insert tab: its groups, order, names and ▾ menus. At
  // 1512 pt every command is a large button; at 1200 pt 3D Models, SmartArt
  // and Chart become small rows and Date & Time, Slide Number and Object
  // icon-only buttons, which is where PowerPoint stops shrinking this tab.
  import { tick } from 'svelte';
  import { eventTarget } from '../core/dom-root.ts';
  import {
    getSlideLayoutName,
    getSlideLayoutPartName,
    getSlideLayouts,
    type SlideLayoutData,
  } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { insertMedia, insertScreenshot, insertTable, insertTextBox } from '../core/insert-objects.ts';
  import { capabilityById } from '../manifest/index.ts';
  import { capLabel, getLocale, t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import LayoutThumbnail from '../ui/LayoutThumbnail.svelte';
  import { captionLines } from './caption.ts';
  import { placeBelowTrigger } from './place-menu.ts';
  import { PRESET } from './config.ts';

  const editor = getEditor();
  const doc = editor.doc;

  // Below this ribbon width (CSS px = pt) the tab takes PowerPoint's 1200 pt
  // layout. Longer labels (Japanese) also take it while the row overflows.
  const SMALL_BELOW = 1350;
  let width = $state(typeof window === 'undefined' ? Number.POSITIVE_INFINITY : window.innerWidth);
  let overflowSmall = $state(false);
  let root = $state<HTMLDivElement>();
  const small = $derived(width < SMALL_BELOW || overflowSmall);
  $effect.pre(() => {
    width;
    getLocale();
    overflowSmall = false;
  });
  $effect(() => {
    width;
    getLocale();
    if (root && !small && root.scrollWidth > root.clientWidth + 1) overflowSmall = true;
  });

  type Menu = 'newSlide' | 'table' | 'pictures' | 'screenshot' | 'chart' | 'textBox' | 'video' | 'audio';
  let open = $state<Menu | null>(null);
  const hasSlide = $derived(doc.currentSlide !== null);
  const layouts = $derived.by(() => { doc.version; return getSlideLayouts(doc.pres); });

  // PowerPoint's Insert Table grid.
  const GRID_COLUMNS = 10;
  const GRID_ROWS = 8;
  let hover = $state<{ rows: number; columns: number } | null>(null);

  // Mac PowerPoint's Chart ▾ categories. The dialog writes the first ones;
  // the library has no writer for the others yet.
  const CHARTS: readonly { readonly label: string; readonly kind?: string }[] = [
    { label: 'Column', kind: 'column' },
    { label: 'Line', kind: 'line' },
    { label: 'Pie', kind: 'pie' },
    { label: 'Bar', kind: 'bar' },
    { label: 'Area', kind: 'area' },
    { label: 'X Y (Scatter)' },
    { label: 'Stock' },
    { label: 'Surface' },
    { label: 'Radar', kind: 'radar' },
    { label: 'Treemap' },
    { label: 'Sunburst' },
    { label: 'Histogram' },
    { label: 'Box and Whisker' },
    { label: 'Waterfall' },
    { label: 'Funnel' },
  ];

  // The ribbon's tooltip for a command backed by a capability (see Ribbon.svelte).
  function capTip(id: string): string {
    const cap = capabilityById.get(id);
    return cap ? `${capLabel(cap)} — ${cap.id}` : id;
  }

  async function toggle(menu: Menu) {
    open = open === menu ? null : menu;
    hover = null;
    if (open) {
      await tick();
      root?.querySelector<HTMLElement>('.menu button:not(:disabled)')?.focus();
    }
  }
  function run(action: () => void) {
    open = null;
    action();
  }
  function insertSlide(layout: SlideLayoutData) {
    run(() => editor.invoke('addSlide', { options: { layout } }));
  }
  function table(rows: number, columns: number) {
    run(() => insertTable(editor, t('Insert table'), rows, columns, { header: true, banded: true }));
  }
  function dismiss(event: PointerEvent) {
    if (open && !(eventTarget(event) as Element).closest?.('.insert .anchor')) open = null;
  }
  function keydown(event: KeyboardEvent) {
    if (event.key !== 'Escape' || !open) return;
    event.preventDefault();
    event.stopPropagation();
    const trigger = root?.querySelector<HTMLElement>(`[data-menu="${open}"]`);
    open = null;
    trigger?.focus();
  }
</script>

<svelte:window onpointerdown={dismiss} onkeydown={keydown} />

<!-- A large button: 32 pt icon over a caption; `menu` puts ▾ beside the icon. -->
{#snippet big(label: string, icon: string, options: { menu?: Menu; gallery?: boolean; arrow?: boolean; disabled?: boolean; tip?: string; cap?: string; onclick?: (button: HTMLElement) => void })}
  {@const arrow = options.menu !== undefined || options.gallery || options.arrow}
  <button
    class="big"
    class:has-arrow={arrow}
    data-menu={options.menu}
    aria-label={t(label)}
    title={options.tip ? t(options.tip) : options.cap ? capTip(options.cap) : undefined}
    aria-haspopup={options.menu || options.gallery ? 'menu' : undefined}
    aria-expanded={options.menu ? open === options.menu : undefined}
    disabled={options.disabled}
    onclick={(event) => {
      if (options.menu) void toggle(options.menu);
      else options.onclick?.(event.currentTarget);
    }}
  >
    {#if arrow}<span class="icon-row"><Icon name={icon} size={32} /><span class="arrow" aria-hidden="true">⌄</span></span>{:else}<Icon name={icon} size={32} />{/if}
    <span class="caption">{captionLines(t(label))}</span>
  </button>
{/snippet}

<!-- PowerPoint's 1200 pt rows: an 18 pt icon, the caption and ▾. -->
{#snippet row(label: string, icon: string, options: { menu?: Menu; disabled?: boolean; tip?: string })}
  <button class="row" data-menu={options.menu} aria-label={t(label)} title={options.tip ? t(options.tip) : undefined} disabled={options.disabled} aria-haspopup={options.menu ? 'menu' : undefined} aria-expanded={options.menu ? open === options.menu : undefined} onclick={() => { if (options.menu) void toggle(options.menu); }}><Icon name={icon} size={18} /><span>{t(label)}</span><span class="arrow" aria-hidden="true">⌄</span></button>
{/snippet}

{#snippet menuItem(label: string, action: (() => void) | null, tip?: string, cap?: string)}
  <button role="menuitem" disabled={!action} title={tip ? t(tip) : cap ? capTip(cap) : undefined} onclick={() => action && run(action)}>{t(label)}</button>
{/snippet}

<div class="insert" bind:this={root} bind:clientWidth={width}>
  <section class="cluster" aria-label={t('Slides')}>
    <div class="anchor overlay">
      <button class="big has-arrow" aria-label={t('New Slide')} disabled={!editor.canRun('addSlide') && !editor.canRun('addBlankSlide')} onclick={() => editor.addNewSlide()}><span class="icon-row"><Icon name="new-slide" size={32} /><span class="arrow-space" aria-hidden="true"></span></span><span class="caption">{captionLines(t('New Slide'))}</span></button>
      <button class="side" data-menu="newSlide" aria-label={t('New Slide options')} aria-haspopup="menu" aria-expanded={open === 'newSlide'} onclick={() => toggle('newSlide')}>⌄</button>
      {#if open === 'newSlide'}
        <div class="menu" role="menu" aria-label={t('New Slide')} use:placeBelowTrigger>
          <div class="heading">{t('Layouts')}</div>
          <div class="layout-grid">
            {#each layouts as layout (getSlideLayoutPartName(layout))}
              <button class="layout-item" role="menuitem" onclick={() => insertSlide(layout)}><LayoutThumbnail pres={doc.pres} {layout} /><span>{t(getSlideLayoutName(layout))}</span></button>
            {/each}
          </div>
          <hr />
          {@render menuItem('Duplicate Selected Slides', hasSlide ? () => editor.invoke('duplicateSlide') : null)}
        </div>
      {/if}
    </div>
  </section>

  <section class="cluster" aria-label={t('Tables')}>
    <div class="anchor">
      {@render big('Table', 'table', { menu: 'table', disabled: !hasSlide })}
      {#if open === 'table'}
        <div class="menu" role="menu" aria-label={t('Table')} use:placeBelowTrigger>
          <div class="heading">{hover ? `${hover.columns}x${hover.rows} ${t('Table')}` : t('Insert Table')}</div>
          <div class="grid" role="group" aria-label={t('Insert Table')} style:grid-template-columns="repeat({GRID_COLUMNS}, 16px)" onpointerleave={() => (hover = null)}>
            {#each { length: GRID_ROWS } as _, r (r)}
              {#each { length: GRID_COLUMNS } as _, c (c)}
                <button
                  class="cell"
                  role="menuitem"
                  aria-label={`${c + 1}x${r + 1} ${t('Table')}`}
                  class:on={hover !== null && r < hover.rows && c < hover.columns}
                  onpointerenter={() => (hover = { rows: r + 1, columns: c + 1 })}
                  onfocus={() => (hover = { rows: r + 1, columns: c + 1 })}
                  onclick={() => table(r + 1, c + 1)}
                ></button>
              {/each}
            {/each}
          </div>
          <hr />
          {@render menuItem('Insert Table...', () => editor.runOrPrompt('addSlideTable'), undefined, 'addSlideTable')}
          {@render menuItem('Draw Table', null, 'Drawing a table with the pointer is not available in this editor yet.')}
        </div>
      {/if}
    </div>
  </section>

  <section class="cluster" aria-label={t('Images')}>
    <div class="anchor">
      {@render big('Pictures', 'picture', { menu: 'pictures', disabled: !editor.canRun('addSlideImage') })}
      {#if open === 'pictures'}
        <div class="menu" role="menu" aria-label={t('Pictures')} use:placeBelowTrigger>
          {@render menuItem('Photo Browser...', null, 'The Photos library is not available in the browser.')}
          {@render menuItem('Picture from File...', () => editor.runOrPrompt('addSlideImage'), undefined, 'addSlideImage')}
          {@render menuItem('Stock Images...', null, 'Stock images need the Microsoft 365 service.')}
          {@render menuItem('Online Pictures...', null, 'Online pictures need the Microsoft 365 service.')}
        </div>
      {/if}
    </div>
    <div class="anchor">
      {@render big('Screenshot', 'screenshot', { menu: 'screenshot', disabled: !hasSlide || typeof navigator === 'undefined' || !navigator.mediaDevices?.getDisplayMedia })}
      {#if open === 'screenshot'}
        <div class="menu" role="menu" aria-label={t('Screenshot')} use:placeBelowTrigger>
          <div class="heading">{t('Available Windows')}</div>
          <!-- The browser lists the windows itself, in its screen-sharing picker. -->
          {@render menuItem('Choose a Window or Screen...', () => void insertScreenshot(editor, t('Screenshot')))}
          <hr />
          {@render menuItem('Screen Clipping', null, 'The browser cannot clip a region of the screen.')}
        </div>
      {/if}
    </div>
  </section>

  <section class="cluster" aria-label={t('Camera')}>
    {@render big('Cameo', 'cameo', { disabled: true, tip: 'Recording is not available in the browser.' })}
  </section>

  <section class="cluster" aria-label={t('Illustrations')}>
    {@render big('Shapes', 'shapes', { gallery: true, disabled: !editor.canRun('addSlideShape'), onclick: (button) => editor.openShapeGallery(button) })}
    {@render big('Icons', 'icons', { disabled: true, tip: 'The Office icon library is not available here.' })}
    {#if small}
      <div class="stack">
        {@render row('3D Models', 'cube', { disabled: true, tip: '3D models are not supported by the library yet.' })}
        {@render row('SmartArt', 'smartart', { disabled: true, tip: 'SmartArt is not supported by the library yet.' })}
        <div class="anchor">
          {@render row('Chart', 'chart', { menu: 'chart', disabled: !editor.canRun('addSlideChart') })}
          {#if open === 'chart'}{@render chartMenu()}{/if}
        </div>
      </div>
    {:else}
      {@render big('3D Models', 'cube', { arrow: true, disabled: true, tip: '3D models are not supported by the library yet.' })}
      {@render big('SmartArt', 'smartart', { arrow: true, disabled: true, tip: 'SmartArt is not supported by the library yet.' })}
      <div class="anchor">
        {@render big('Chart', 'chart', { menu: 'chart', disabled: !editor.canRun('addSlideChart') })}
        {#if open === 'chart'}{@render chartMenu()}{/if}
      </div>
    {/if}
  </section>

  <section class="cluster" aria-label={t('Links')}>
    {@render big('Zoom', 'zoom-slide', { arrow: true, disabled: true, tip: 'Slide zoom is not supported by the library yet.' })}
    {@render big('Link', 'link', { disabled: !editor.canRun('setShapeHyperlink'), cap: 'setShapeHyperlink', onclick: () => editor.runOrPrompt('setShapeHyperlink') })}
    {@render big('Action', 'action', { disabled: !editor.canRun('setShapeClickAction'), cap: 'setShapeClickAction', onclick: () => editor.runOrPrompt('setShapeClickAction') })}
  </section>

  <section class="cluster" aria-label={t('Comments')}>
    {@render big('Comment', 'comment', { disabled: !editor.canRun('addSlideComment'), cap: 'addSlideComment', onclick: () => editor.runOrPrompt('addSlideComment') })}
  </section>

  <section class="cluster" aria-label={t('Text')}>
    <div class="anchor overlay">
      <button class="big has-arrow" aria-label={t('Text Box')} title={capTip('addSlideTextBox')} disabled={!editor.canRun('addSlideTextBox')} onclick={() => editor.runOrPrompt('addSlideTextBox', PRESET.textBox)}><span class="icon-row"><Icon name="textbox" size={32} /><span class="arrow-space" aria-hidden="true"></span></span><span class="caption">{t('Text Box')}</span></button>
      <button class="side" data-menu="textBox" aria-label={t('Text Box options')} aria-haspopup="menu" aria-expanded={open === 'textBox'} disabled={!hasSlide} onclick={() => toggle('textBox')}>⌄</button>
      {#if open === 'textBox'}
        <div class="menu" role="menu" aria-label={t('Text Box')} use:placeBelowTrigger>
          {@render menuItem('Draw Horizontal Text Box', () => insertTextBox(editor, t('Text Box'), t('Text'), false))}
          {@render menuItem('Draw Vertical Text Box', () => insertTextBox(editor, t('Text Box'), t('Text'), true))}
        </div>
      {/if}
    </div>
    {@render big('Header & Footer', 'header-footer', { disabled: !hasSlide, onclick: () => (editor.activeDialog = 'headerFooter') })}
    {@render big('WordArt', 'wordart', { gallery: true, disabled: !hasSlide, onclick: (button) => editor.openWordArtGallery(button) })}
    {#if small}
      <div class="stack tools">
        <button class="tool" aria-label={t('Date & Time')} title={t('Date & Time')} disabled={!editor.canRun('setShapeTextField')} onclick={() => editor.runOrPrompt('setShapeTextField')}><Icon name="calendar" size={18} /></button>
        <button class="tool" aria-label={t('Slide Number')} title={t('Slide Number')} disabled={!editor.canRun('setShapeTextField')} onclick={() => editor.runOrPrompt('setShapeTextField', { type: 'slidenum' })}><Icon name="slide-number" size={18} /></button>
        <button class="tool" aria-label={t('Object')} title={t('Embedded OLE objects are not supported by the library yet.')} disabled><Icon name="object" size={18} /></button>
      </div>
    {:else}
      <!-- Both insert a field PowerPoint keeps up to date into the selected
           box; Date & Time asks for the format first, as the native dialog does. -->
      {@render big('Date & Time', 'calendar', { cap: 'setShapeTextField', disabled: !editor.canRun('setShapeTextField'), onclick: () => editor.runOrPrompt('setShapeTextField') })}
      {@render big('Slide Number', 'slide-number', { disabled: !editor.canRun('setShapeTextField'), onclick: () => editor.runOrPrompt('setShapeTextField', { type: 'slidenum' }) })}
      {@render big('Object', 'object', { disabled: true, tip: 'Embedded OLE objects are not supported by the library yet.' })}
    {/if}
  </section>

  <section class="cluster" aria-label={t('Symbols')}>
    {@render big('Equation', 'equation', { arrow: true, disabled: true, tip: 'Equations are not supported by the library yet.' })}
    <!-- Like PowerPoint, Symbol needs a text cursor to insert at. -->
    {@render big('Symbol', 'symbol', { disabled: !editor.inlineTextFormat?.insertText, onclick: (button) => editor.openSymbolPicker(button) })}
  </section>

  <section class="cluster" aria-label={t('Media')}>
    <div class="anchor">
      {@render big('Video', 'video', { menu: 'video', disabled: !hasSlide })}
      {#if open === 'video'}
        <div class="menu" role="menu" aria-label={t('Video')} use:placeBelowTrigger>
          {@render menuItem('Movie Browser...', null, 'The Photos and Music libraries are not available in the browser.')}
          {@render menuItem('Movie from File...', () => void insertMedia(editor, 'video', t('Video')))}
          {@render menuItem('Online Movie...', null, 'Online videos are not supported by the library yet.')}
        </div>
      {/if}
    </div>
    <div class="anchor">
      {@render big('Audio', 'audio', { menu: 'audio', disabled: !hasSlide })}
      {#if open === 'audio'}
        <div class="menu" role="menu" aria-label={t('Audio')} use:placeBelowTrigger>
          {@render menuItem('Audio Browser...', null, 'The Photos and Music libraries are not available in the browser.')}
          {@render menuItem('Audio from File...', () => void insertMedia(editor, 'audio', t('Audio')))}
          {@render menuItem('Record Audio...', null, 'Recording is not available in the browser.')}
        </div>
      {/if}
    </div>
  </section>
</div>

{#snippet chartMenu()}
  <div class="menu" role="menu" aria-label={t('Chart')} use:placeBelowTrigger>
    {#each CHARTS as chart (chart.label)}
      {@render menuItem(chart.label, chart.kind ? () => editor.runOrPrompt('addSlideChart', { kind: chart.kind }) : null, chart.kind ? undefined : 'This chart type is not supported by the library yet.')}
    {/each}
  </div>
{/snippet}

<style>
  /* The Home tab's metrics (see HomeRibbon.svelte): a 72 pt row of groups
     separated by a rule with 10 pt either side. Large buttons fill the row and
     are at least 38 pt wide, or 50 pt with ▾ beside the icon; captions of more
     than one word take two lines. */
  .insert { display: flex; align-items: stretch; min-width: 0; width: 100%; height: 72px; }
  .cluster { position: relative; display: flex; align-items: center; padding: 0 9px; border-right: 1px solid var(--ok-border); flex: none; }
  .cluster:first-child { padding-left: 3px; }
  .cluster:last-child { border-right: none; }
  button { font: inherit; color: var(--ok-text); background: none; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  button:hover:not(:disabled) { background: var(--ok-hover); }
  button:disabled { opacity: 0.4; cursor: default; }
  .big { display: flex; flex-direction: column; align-items: center; justify-content: flex-start; align-self: stretch; gap: 2px; min-width: 38px; padding: 4px 1px; font-size: 11px; line-height: 1.15; }
  .big.has-arrow { min-width: 50px; }
  .caption { white-space: pre-line; text-align: center; }
  .icon-row { display: flex; align-items: center; gap: 1px; }
  .arrow { font-size: 10px; }
  .anchor { position: relative; display: flex; align-self: stretch; }
  .side { padding: 0; font-size: 12px; }
  .arrow-space { width: 9px; }
  /* A split button: the ▾ sits over the space beside the icon, as in
     PowerPoint's 50 pt menu buttons. */
  .overlay > .side { position: absolute; top: 12px; right: 2px; width: 12px; height: 22px; }
  .stack { display: flex; flex-direction: column; align-items: flex-start; justify-content: center; gap: 0; align-self: stretch; }
  .stack .anchor { align-self: auto; }
  .row { display: flex; align-items: center; gap: 4px; height: 22px; padding: 0 4px; font-size: 11px; white-space: nowrap; }
  .tool { display: flex; align-items: center; justify-content: center; width: 24px; height: 22px; padding: 0; }
  .menu { position: fixed; z-index: 400; display: flex; flex-direction: column; min-width: 200px; max-height: 70vh; overflow-y: auto; padding: 4px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .menu > button { display: flex; align-items: center; height: 24px; padding: 0 10px; text-align: left; font-size: 12px; white-space: nowrap; }
  .menu .heading { padding: 4px 8px; font-size: 11px; font-weight: 600; color: var(--ok-text-2); }
  .menu hr { width: 100%; border: none; border-top: 1px solid var(--ok-border); margin: 5px 0; }
  .layout-grid { display: grid; grid-template-columns: repeat(3, 120px); gap: 6px; padding: 2px 4px; }
  .layout-item { display: flex; flex-direction: column; align-items: stretch; gap: 3px; padding: 4px; font-size: 11px; text-align: center; }
  .layout-item span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .grid { display: grid; gap: 2px; padding: 2px 8px 4px; }
  .cell { width: 16px; height: 16px; padding: 0; border: 1px solid var(--ok-border-strong); border-radius: 1px; background: var(--ok-panel); }
  .cell.on, .cell:focus-visible { border-color: var(--ok-accent); background: var(--ok-selected); }
</style>

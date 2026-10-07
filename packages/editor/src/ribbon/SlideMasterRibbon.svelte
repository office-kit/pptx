<script lang="ts">
  import { eventTarget } from '../core/dom-root.ts';
  // Mac PowerPoint 16's Slide Master tab, measured in a 1512 × 900 pt window:
  // Insert Slide Master, Insert Layout and a Delete / Rename / Preserve column
  // | Master Layout, Insert Placeholder and the Title / Footers checkboxes
  // | Themes | Colors, Fonts, Background Styles, Hide Background Graphics
  // | Slide Size | Close Master. Edits act on the master or layout selected in
  // the pane. As natively, the layout commands (Insert Placeholder, Title,
  // Footers, Hide Background Graphics) are disabled with the master selected,
  // Master Layout and Preserve with a layout selected, and Delete while slides
  // use the selection.
  import { tick } from 'svelte';
  import {
    findSlideLayoutByPartName,
    findSlidesByLayoutPartName,
    getSlideLayoutPartName,
    getSlideLayoutPlaceholders,
    getSlideMasterLayouts,
    getSlideMasterPartNames,
    getSlideMasterPlaceholders,
    getSlideMasterUsageCounts,
    getSlideSize,
    isSlideLayoutBackgroundGraphicsHidden,
    isSlideMasterPreserved,
    setPresentationFonts,
    setPresentationTheme,
    setSlideMasterPlaceholderIncluded,
    type LayoutPlaceholderKind,
    type MasterPlaceholderType,
    type SlideLayoutData,
  } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { THEMES } from '../core/design-presets.ts';
  import { getLocale, t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import BackgroundStyles from './BackgroundStyles.svelte';
  import { captionLines } from './caption.ts';
  import { placeBelowTrigger } from './place-menu.ts';
  import SlideSizeMenu from './SlideSizeMenu.svelte';

  const editor = getEditor();
  const doc = editor.doc;
  const master = $derived(doc.layoutTarget?.master ?? null);
  const layout = $derived.by(() => {
    doc.version;
    const partName = doc.layoutTarget?.partName;
    return partName ? findSlideLayoutByPartName(doc.pres, partName) : null;
  });
  const types = $derived(new Set(layout ? getSlideLayoutPlaceholders(layout).map((placeholder) => placeholder.type) : []));
  const hasTitle = $derived(types.has('title') || types.has('ctrTitle'));
  const hasFooters = $derived(types.has('dt') || types.has('ftr') || types.has('sldNum'));
  const backgroundHidden = $derived(layout ? isSlideLayoutBackgroundGraphicsHidden(layout) : false);
  const preserved = $derived.by(() => { doc.version; return master !== null && isSlideMasterPreserved(doc.pres, master); });
  const masterLayouts = $derived.by<readonly SlideLayoutData[]>(() => { doc.version; return master ? getSlideMasterLayouts(doc.pres, master) : []; });

  // Why Delete is unavailable for the selection, or null when it can run.
  const deleteBlocked = $derived.by(() => {
    doc.version;
    if (master === null) return t('Select a master or layout.');
    if (layout) {
      if (findSlidesByLayoutPartName(doc.pres, getSlideLayoutPartName(layout)).length > 0) return t('Slides use this layout.');
      if (masterLayouts.length < 2) return t('A slide master needs at least one layout.');
      return null;
    }
    if (getSlideMasterPartNames(doc.pres).length < 2) return t('A presentation needs at least one slide master.');
    if ((getSlideMasterUsageCounts(doc.pres)[master] ?? 0) > 0) return t('Slides use this master.');
    return null;
  });

  // Japanese PowerPoint words these differently here than on the Design and Insert tabs.
  const MASTER_TAB_JA: Record<string, string> = { Colors: '色', 'Background Styles': '背景スタイル', 'Hide Background Graphics': '背景グラフィックを表示しない', Picture: '図' };
  const label = (key: string) => (getLocale() === 'ja' ? MASTER_TAB_JA[key] ?? t(key) : key);
  const ON_LAYOUT = 'Select a layout to change it.';
  const ON_MASTER = 'Select the slide master to change it.';

  // Insert Placeholder's menu, in PowerPoint's order.
  const PLACEHOLDER_KINDS: ReadonlyArray<readonly [LayoutPlaceholderKind, string]> = [
    ['content', 'Content'],
    ['verticalContent', 'Content (Vertical)'],
    ['text', 'Text'],
    ['verticalText', 'Text (Vertical)'],
    ['picture', 'Picture'],
    ['chart', 'Chart'],
    ['table', 'Table'],
    ['smartArt', 'SmartArt'],
    ['media', 'Media'],
    ['onlineImage', 'Online Image'],
  ];
  // The Master Layout dialog's checkboxes, in PowerPoint's order.
  const MASTER_PLACEHOLDERS: ReadonlyArray<readonly [MasterPlaceholderType, string]> = [
    ['title', 'Title'],
    ['body', 'Text'],
    ['dt', 'Date'],
    ['sldNum', 'Slide number'],
    ['ftr', 'Footer'],
  ];

  type Menu = 'themes' | 'placeholder' | 'masterLayout';
  let open = $state<Menu | null>(null);
  let root = $state<HTMLDivElement>();
  let masterChoice = $state<Record<string, boolean>>({});
  async function toggle(menu: Menu) {
    open = open === menu ? null : menu;
    if (open === 'masterLayout' && master) {
      const present = new Set(getSlideMasterPlaceholders(doc.pres, master).map((placeholder) => placeholder.type));
      masterChoice = Object.fromEntries(MASTER_PLACEHOLDERS.map(([type]) => [type, present.has(type) || (type === 'title' && present.has('ctrTitle'))]));
    }
    if (open) {
      await tick();
      root?.querySelector<HTMLElement>('.menu [aria-checked="true"], .menu button:not(:disabled), .menu input')?.focus();
    }
  }
  function applyTheme(index: number) {
    open = null;
    const preset = THEMES[index]!;
    const { name: _name, ...fonts } = preset.fonts;
    doc.transact(t('Themes'), () => {
      setPresentationTheme(doc.pres, preset.colors);
      setPresentationFonts(doc.pres, fonts);
    });
  }

  function insertMaster() {
    const name = editor.invoke('addSlideMaster');
    if (typeof name === 'string') editor.selectMasterCell(name, null);
  }
  function insertLayout() {
    if (!master) return;
    // PowerPoint adds the layout after the selected one, or last with the master selected.
    const at = layout ? masterLayouts.findIndex((item) => getSlideLayoutPartName(item) === getSlideLayoutPartName(layout)) + 1 : masterLayouts.length;
    const added = editor.invoke('addSlideLayout', { options: { index: at } }) as SlideLayoutData | undefined;
    if (added) editor.selectMasterCell(master, getSlideLayoutPartName(added));
  }
  function remove() {
    if (!master || deleteBlocked) return;
    if (layout) {
      const index = masterLayouts.findIndex((item) => getSlideLayoutPartName(item) === getSlideLayoutPartName(layout));
      editor.invoke('removeSlideLayout');
      // The selection moves up to the layout above, as in PowerPoint's pane.
      const remaining = getSlideMasterLayouts(doc.pres, master);
      const next = index > 0 ? remaining[index - 1] : null;
      editor.selectMasterCell(master, next ? getSlideLayoutPartName(next) : null);
    } else {
      const masters = getSlideMasterPartNames(doc.pres);
      const index = masters.indexOf(master);
      editor.invoke('removeSlideMaster');
      const left = getSlideMasterPartNames(doc.pres);
      editor.selectMasterCell(left[Math.max(0, index - 1)] ?? null, null);
    }
  }
  function rename() {
    editor.runOrPrompt(layout ? 'setSlideLayoutName' : 'setSlideMasterName');
  }
  function insertPlaceholder(kind: LayoutPlaceholderKind) {
    open = null;
    const size = getSlideSize(doc.pres) ?? { width: 12192000, height: 6858000 };
    // PowerPoint draws the placeholder with the pointer; without a drag the
    // editor places a box of 40% of the slide in the middle, to move after.
    const w = Math.round(size.width * 0.4);
    const h = Math.round(size.height * 0.4);
    editor.invoke('addSlideLayoutPlaceholder', { kind, bounds: { x: Math.round((size.width - w) / 2), y: Math.round((size.height - h) / 2), w, h } });
  }
  function applyMasterLayout() {
    if (!master) return;
    const current = new Set(getSlideMasterPlaceholders(doc.pres, master).map((placeholder) => placeholder.type));
    const changes = MASTER_PLACEHOLDERS.filter(([type]) => masterChoice[type] !== (current.has(type) || (type === 'title' && current.has('ctrTitle'))));
    open = null;
    if (changes.length === 0) return;
    // One undo step for the dialog, as in PowerPoint.
    doc.transact(t('Master Layout'), () => {
      for (const [type] of changes) setSlideMasterPlaceholderIncluded(doc.pres, master, type, masterChoice[type] === true);
    });
  }
</script>

<svelte:window onpointerdown={(event) => { if (open && !(eventTarget(event) as Element).closest?.('.slide-master-tab .anchor')) open = null; }} onkeydown={(event) => { if (open && event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); open = null; } }} />

<div class="slide-master-tab" bind:this={root}>
  <section class="cluster" role="group" aria-label={t('Edit Master')}>
    <button class="big" aria-label={t('Insert Slide Master')} style:--w="63px" onclick={insertMaster}><Icon name="new-slide" size={32} /><span class="caption">{captionLines(t('Insert Slide Master'))}</span></button>
    <button class="big" aria-label={t('Insert Layout')} style:--w="40px" disabled={master === null} onclick={insertLayout}><Icon name="slide-content" size={32} /><span class="caption">{captionLines(t('Insert Layout'))}</span></button>
    <div class="column">
      <button class="small" disabled={deleteBlocked !== null} title={deleteBlocked ?? undefined} onclick={remove}><Icon name="trash" size={16} />{t('Delete')}</button>
      <button class="small" disabled={layout ? !editor.canRun('setSlideLayoutName') : master === null} onclick={rename}><Icon name="rename" size={16} />{t('Rename')}</button>
      <button class="small" aria-pressed={preserved} disabled={layout !== null || master === null} title={layout ? t(ON_MASTER) : undefined} onclick={() => editor.invoke('setSlideMasterPreserved', { preserved: !preserved })}><Icon name="reset" size={16} />{t('Preserve')}</button>
    </div>
  </section>
  <section class="cluster" role="group" aria-label={t('Master Layout')}>
    <div class="anchor">
      <button class="big" aria-label={t('Master Layout')} style:--w="50px" aria-haspopup="dialog" aria-expanded={open === 'masterLayout'} disabled={layout !== null || master === null} title={layout ? t(ON_MASTER) : undefined} onclick={() => toggle('masterLayout')}><span class="icon-row"><Icon name="layout" size={32} /><span class="arrow" aria-hidden="true">⌄</span></span><span class="caption">{captionLines(t('Master Layout'))}</span></button>
      {#if open === 'masterLayout'}
        <div class="menu dialog" role="dialog" use:placeBelowTrigger aria-label={t('Master Layout')}>
          <span class="heading">{t('Placeholders')}</span>
          {#each MASTER_PLACEHOLDERS as [type, name] (type)}
            <label class="option"><input type="checkbox" bind:checked={masterChoice[type]} />{t(name)}</label>
          {/each}
          <div class="buttons">
            <button class="push" onclick={() => (open = null)}>{t('Cancel')}</button>
            <button class="push default" onclick={applyMasterLayout}>{t('OK')}</button>
          </div>
        </div>
      {/if}
    </div>
    <div class="anchor">
      <button class="big" aria-label={t('Insert Placeholder')} style:--w="66px" aria-haspopup="menu" aria-expanded={open === 'placeholder'} disabled={layout === null} title={layout ? undefined : t(ON_LAYOUT)} onclick={() => toggle('placeholder')}><span class="icon-row"><Icon name="textbox" size={32} /><span class="arrow" aria-hidden="true">⌄</span></span><span class="caption">{captionLines(t('Insert Placeholder'))}</span></button>
      {#if open === 'placeholder'}
        <div class="menu" role="menu" use:placeBelowTrigger aria-label={t('Insert Placeholder')}>
          {#each PLACEHOLDER_KINDS as [kind, name] (kind)}<button class="action" role="menuitem" onclick={() => insertPlaceholder(kind)}>{label(name)}</button>{/each}
        </div>
      {/if}
    </div>
    <div class="checks">
      <label title={layout ? undefined : t(ON_LAYOUT)}><input type="checkbox" checked={layout ? hasTitle : true} disabled={layout === null} onchange={(event) => editor.invoke('setSlideLayoutTitleIncluded', { included: event.currentTarget.checked })} />{t('Title')}</label>
      <label title={layout ? undefined : t(ON_LAYOUT)}><input type="checkbox" checked={layout ? hasFooters : true} disabled={layout === null} onchange={(event) => editor.invoke('setSlideLayoutFootersIncluded', { included: event.currentTarget.checked })} />{t('Footers')}</label>
    </div>
  </section>
  <section class="cluster" role="group" aria-label={t('Edit Theme')}>
    <div class="anchor">
      <button class="big" style:--w="50px" aria-haspopup="menu" aria-expanded={open === 'themes'} onclick={() => toggle('themes')}><span class="icon-row"><Icon name="theme" size={32} /><span class="arrow" aria-hidden="true">⌄</span></span><span class="caption">{t('Themes')}</span></button>
      {#if open === 'themes'}
        <div class="menu" role="menu" use:placeBelowTrigger aria-label={t('Themes')}>
          {#each THEMES as preset, i (preset.name)}<button class="action" role="menuitem" onclick={() => applyTheme(i)}>{t(preset.name)}</button>{/each}
        </div>
      {/if}
    </div>
  </section>
  <section class="cluster" role="group" aria-label={t('Background')}>
    <div class="column wide">
      <button class="row" disabled={!editor.canRun('setPresentationTheme')} onclick={() => editor.runOrPrompt('setPresentationTheme')}><Icon name="theme" size={16} />{label('Colors')}<span class="arrow" aria-hidden="true">⌄</span></button>
      <button class="row" disabled={!editor.canRun('setPresentationFonts')} onclick={() => editor.runOrPrompt('setPresentationFonts')}><Icon name="font" size={16} />{t('Fonts')}<span class="arrow" aria-hidden="true">⌄</span></button>
    </div>
    <div class="column wide">
      <BackgroundStyles small label={label('Background Styles')} />
      <label class="row" title={layout ? undefined : t(ON_LAYOUT)}><input type="checkbox" checked={backgroundHidden} disabled={layout === null} onchange={(event) => editor.invoke('setSlideLayoutBackgroundGraphicsHidden', { hidden: event.currentTarget.checked })} />{label('Hide Background Graphics')}</label>
    </div>
  </section>
  <section class="cluster" role="group" aria-label={t('Size')}>
    <SlideSizeMenu />
  </section>
  <section class="cluster" role="group" aria-label={t('Close')}>
    <button class="big" aria-label={t('Close Master')} style:--w="41px" onclick={() => editor.setViewMode('normal')}><Icon name="close-master" size={32} /><span class="caption">{captionLines(t('Close Master'))}</span></button>
  </section>
</div>

<style>
  /* The View tab's metrics: a 72 pt row, 10 pt cluster padding, 22 pt small
     rows and 26 pt checkbox and ▾ rows. */
  .slide-master-tab { display: flex; align-items: stretch; width: 100%; min-width: 0; height: 72px; }
  .cluster { display: flex; flex: none; align-items: stretch; padding: 0 10px; border-right: 1px solid var(--ok-border); }
  .cluster:first-child { padding-left: 0; }
  .cluster:last-child { border-right: none; }
  button, label { font: inherit; color: var(--ok-text); }
  button { background: none; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  button:hover:not(:disabled) { background: var(--ok-hover); }
  button:disabled, label:has(input:disabled) { opacity: 0.4; cursor: default; }
  button[aria-pressed='true'] { background: var(--ok-selected); border-color: var(--ok-selected-border); }
  .big { display: flex; flex-direction: column; align-items: center; justify-content: flex-start; gap: 2px; min-width: var(--w); padding: 4px 1px; font-size: 11px; line-height: 1.15; }
  .caption { white-space: pre-line; text-align: center; }
  .icon-row { display: flex; align-items: center; gap: 1px; }
  .arrow { font-size: 10px; }
  .anchor { position: relative; display: flex; }
  .column { display: flex; flex-direction: column; justify-content: flex-start; }
  .small { display: flex; align-items: center; gap: 5px; height: 22px; padding: 0 6px; font-size: 12px; white-space: nowrap; }
  .checks { display: flex; flex-direction: column; justify-content: flex-start; gap: 6px; margin-top: 4px; }
  .checks label, label.row { display: flex; align-items: center; gap: 6px; height: 26px; padding: 0 6px; font-size: 12px; white-space: nowrap; }
  .column.wide { gap: 6px; margin-top: 4px; }
  .row { display: flex; align-items: center; gap: 5px; height: 26px; padding: 0 6px; font-size: 12px; white-space: nowrap; }
  input { margin: 0; accent-color: var(--ok-accent); }
  .menu { position: fixed; z-index: 400; display: flex; flex-direction: column; min-width: 150px; max-height: 70vh; overflow-y: auto; padding: 4px 0; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .action { display: flex; align-items: center; width: 100%; height: 24px; padding: 0 20px; border: 0; border-radius: 0; text-align: left; font-size: 12px; white-space: nowrap; }
  .dialog { gap: 4px; padding: 10px 14px; min-width: 190px; }
  .heading { font-size: 12px; font-weight: 600; }
  .option { display: flex; align-items: center; gap: 6px; height: 22px; font-size: 12px; }
  .buttons { display: flex; justify-content: flex-end; gap: 8px; margin-top: 6px; }
  .push { min-width: 64px; height: 24px; padding: 0 10px; border: 1px solid var(--ok-border-strong); font-size: 12px; }
  .push.default { background: var(--ok-accent); border-color: var(--ok-accent); color: var(--ok-on-accent, #fff); }
</style>

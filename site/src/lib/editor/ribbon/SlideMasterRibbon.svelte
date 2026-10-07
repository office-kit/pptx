<script lang="ts">
  // Mac PowerPoint 16's Slide Master tab, measured in a 1512 × 900 pt window:
  // Insert Slide Master, Insert Layout and a Delete / Rename / Preserve column
  // | Master Layout, Insert Placeholder and the Title / Footers checkboxes
  // | Themes | Colors, Fonts, Background Styles, Hide Background Graphics
  // | Slide Size | Close Master. Edits act on the master or layout selected in
  // the pane; what the library cannot change is shown disabled with the reason.
  import { tick } from 'svelte';
  import {
    findSlideLayoutByPartName,
    getSlideLayoutPlaceholders,
    isSlideLayoutBackgroundGraphicsHidden,
    setPresentationFonts,
    setPresentationTheme,
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
  const layout = $derived.by(() => {
    doc.version;
    const partName = doc.layoutTarget?.partName;
    return partName ? findSlideLayoutByPartName(doc.pres, partName) : null;
  });
  const types = $derived(new Set(layout ? getSlideLayoutPlaceholders(layout).map((placeholder) => placeholder.type) : []));
  const hasTitle = $derived(types.has('title') || types.has('ctrTitle'));
  const hasFooters = $derived(types.has('dt') || types.has('ftr') || types.has('sldNum'));
  const backgroundHidden = $derived(layout ? isSlideLayoutBackgroundGraphicsHidden(layout) : false);

  // Japanese PowerPoint words these three differently here than on the Design tab.
  const MASTER_TAB_JA: Record<string, string> = { Colors: '色', 'Background Styles': '背景スタイル', 'Hide Background Graphics': '背景グラフィックを表示しない' };
  const label = (key: string) => (getLocale() === 'ja' ? MASTER_TAB_JA[key] ?? t(key) : key);
  const NOT_ADDING = 'Adding masters and layouts is not supported by the library yet.';
  const NOT_PLACEHOLDERS = 'Adding or removing layout placeholders is not supported by the library yet.';

  type Menu = 'themes';
  let open = $state<Menu | null>(null);
  let root = $state<HTMLDivElement>();
  async function toggle(menu: Menu) {
    open = open === menu ? null : menu;
    if (open) {
      await tick();
      root?.querySelector<HTMLElement>('.menu [aria-checked="true"], .menu button:not(:disabled)')?.focus();
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
</script>

<svelte:window onpointerdown={(event) => { if (open && !(event.target as Element).closest?.('.slide-master-tab .anchor')) open = null; }} onkeydown={(event) => { if (open && event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); open = null; } }} />

<div class="slide-master-tab" bind:this={root}>
  <section class="cluster" role="group" aria-label={t('Edit Master')}>
    <button class="big" aria-label={t('Insert Slide Master')} style:--w="63px" disabled title={t(NOT_ADDING)}><Icon name="new-slide" size={32} /><span class="caption">{captionLines(t('Insert Slide Master'))}</span></button>
    <button class="big" aria-label={t('Insert Layout')} style:--w="40px" disabled title={t(NOT_ADDING)}><Icon name="slide-content" size={32} /><span class="caption">{captionLines(t('Insert Layout'))}</span></button>
    <div class="column">
      <button class="small" disabled title={t('Deleting masters and layouts is not supported by the library yet.')}><Icon name="trash" size={16} />{t('Delete')}</button>
      <button class="small" disabled={!editor.canRun('setSlideLayoutName')} title={layout ? undefined : t('Renaming a slide master is not supported by the library yet.')} onclick={() => editor.runOrPrompt('setSlideLayoutName')}><Icon name="rename" size={16} />{t('Rename')}</button>
      <button class="small" disabled title={t('Preserving masters is not supported by the library yet.')}><Icon name="reset" size={16} />{t('Preserve')}</button>
    </div>
  </section>
  <section class="cluster" role="group" aria-label={t('Master Layout')}>
    <button class="big" aria-label={t('Master Layout')} style:--w="50px" disabled title={t('Choosing the master placeholders is not supported by the library yet.')}><span class="icon-row"><Icon name="layout" size={32} /><span class="arrow" aria-hidden="true">⌄</span></span><span class="caption">{captionLines(t('Master Layout'))}</span></button>
    <button class="big" aria-label={t('Insert Placeholder')} style:--w="66px" disabled title={t(NOT_PLACEHOLDERS)}><span class="icon-row"><Icon name="textbox" size={32} /><span class="arrow" aria-hidden="true">⌄</span></span><span class="caption">{captionLines(t('Insert Placeholder'))}</span></button>
    <div class="checks">
      <label title={t(NOT_PLACEHOLDERS)}><input type="checkbox" checked={layout ? hasTitle : true} disabled />{t('Title')}</label>
      <label title={t(NOT_PLACEHOLDERS)}><input type="checkbox" checked={layout ? hasFooters : true} disabled />{t('Footers')}</label>
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
      <label class="row" title={t('Hiding background graphics on a layout is not supported by the library yet.')}><input type="checkbox" checked={backgroundHidden} disabled />{label('Hide Background Graphics')}</label>
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
</style>

<script lang="ts">
  // Mac PowerPoint 16's Handout Master and Notes Master tabs: page orientation
  // and Slide Size | the placeholder checkboxes (Header, Footer, Date, Page
  // Number; the notes master adds Slide Image and Body) | the handout's slides
  // per page | Close Master. A deck without the master shows PowerPoint's
  // defaults (every placeholder, six slides per page); the first edit writes
  // the default master.
  import { tick } from 'svelte';
  import {
    getHandoutMasterPlaceholders,
    getHandoutSlidesPerPage,
    getNotesMasterPlaceholders,
    getNotesPageSize,
    type HandoutSlidesPerPage,
    type NotesMasterPlaceholderType,
  } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import { captionLines } from './caption.ts';
  import { placeBelowTrigger } from './place-menu.ts';
  import SlideSizeMenu from './SlideSizeMenu.svelte';

  let { kind }: { kind: 'handoutMaster' | 'notesMaster' } = $props();
  const editor = getEditor();
  const doc = editor.doc;
  type Check = readonly [NotesMasterPlaceholderType, string];
  const checks = $derived<ReadonlyArray<readonly Check[]>>(
    kind === 'handoutMaster'
      ? [[['hdr', 'Header'], ['ftr', 'Footer']], [['dt', 'Date'], ['sldNum', 'Page Number']]]
      : [[['hdr', 'Header'], ['sldImg', 'Slide Image'], ['ftr', 'Footer']], [['dt', 'Date'], ['body', 'Body'], ['sldNum', 'Page Number']]],
  );
  const PER_PAGE: ReadonlyArray<ReadonlyArray<readonly [HandoutSlidesPerPage, string]>> = [
    [[2, '2 Slides'], [3, '3 Slides'], [4, '4 Slides']],
    [[6, '6 Slides'], [9, '9 Slides'], ['outline', 'Slide Outline']],
  ];
  const present = $derived.by<ReadonlySet<string> | null>(() => {
    doc.version;
    const placeholders = kind === 'handoutMaster' ? getHandoutMasterPlaceholders(doc.pres) : getNotesMasterPlaceholders(doc.pres);
    return placeholders ? new Set(placeholders.map((placeholder) => placeholder.type ?? '')) : null;
  });
  const perPage = $derived.by(() => { doc.version; return getHandoutSlidesPerPage(doc.pres); });
  const portrait = $derived.by(() => { doc.version; const size = getNotesPageSize(doc.pres); return size.height >= size.width; });
  const orientationLabel = $derived(kind === 'handoutMaster' ? 'Handout Orientation' : 'Notes Page Orientation');

  let open = $state(false);
  let root = $state<HTMLDivElement>();
  async function toggle() {
    open = !open;
    if (open) {
      await tick();
      root?.querySelector<HTMLElement>('.menu [aria-checked="true"]')?.focus();
    }
  }
  function orient(orientation: 'portrait' | 'landscape') {
    open = false;
    if ((orientation === 'portrait') !== portrait) editor.invoke('setNotesPageOrientation', { orientation });
  }
  function include(type: NotesMasterPlaceholderType, included: boolean) {
    editor.invoke(kind === 'handoutMaster' ? 'setHandoutMasterPlaceholderIncluded' : 'setNotesMasterPlaceholderIncluded', { type, included });
  }
</script>

<svelte:window onpointerdown={(event) => { if (open && !(event.target as Element).closest?.('.page-master-tab .anchor')) open = false; }} onkeydown={(event) => { if (open && event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); open = false; } }} />

<div class="page-master-tab" class:notes={kind === 'notesMaster'} bind:this={root}>
  <section class="cluster" role="group" aria-label={t('Page Setup')}>
    <div class="anchor">
      <button class="big" aria-label={t(orientationLabel)} style:--w={kind === 'handoutMaster' ? '63px' : '64px'} aria-haspopup="menu" aria-expanded={open} onclick={toggle}><span class="icon-row"><Icon name="layout" size={32} /><span class="arrow" aria-hidden="true">⌄</span></span><span class="caption">{captionLines(t(orientationLabel))}</span></button>
      {#if open}
        <div class="menu" role="menu" use:placeBelowTrigger aria-label={t(orientationLabel)}>
          <button class="action" role="menuitemradio" aria-checked={portrait} onclick={() => orient('portrait')}>{t('Portrait')}</button>
          <button class="action" role="menuitemradio" aria-checked={!portrait} onclick={() => orient('landscape')}>{t('Landscape')}</button>
        </div>
      {/if}
    </div>
    <SlideSizeMenu />
  </section>
  <section class="cluster" role="group" aria-label={t('Placeholders')}>
    {#each checks as column, i (i)}
      <div class="checks">
        {#each column as [type, label] (type)}<label><input type="checkbox" checked={present === null || present.has(type)} onchange={(event) => include(type, event.currentTarget.checked)} />{t(label)}</label>{/each}
      </div>
    {/each}
  </section>
  {#if kind === 'handoutMaster'}
    <section class="cluster" role="group" aria-label={t('Slides Per Page')}>
      {#each PER_PAGE as column, i (i)}
        <div class="column">
          {#each column as [value, label] (value)}<button class="small" aria-pressed={perPage === value} onclick={() => editor.invoke('setHandoutSlidesPerPage', { slidesPerPage: value })}><Icon name="layout" size={16} />{t(label)}</button>{/each}
        </div>
      {/each}
    </section>
  {/if}
  <section class="cluster" role="group" aria-label={t('Close')}>
    <button class="big" aria-label={t('Close Master')} style:--w="41px" onclick={() => editor.setViewMode('normal')}><Icon name="close-master" size={32} /><span class="caption">{captionLines(t('Close Master'))}</span></button>
  </section>
</div>

<style>
  /* The Slide Master tab's metrics; the notes master stacks three 22 pt
     checkbox rows where the handout master has two 26 pt ones. */
  .page-master-tab { display: flex; align-items: stretch; width: 100%; min-width: 0; height: 72px; }
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
  .checks { display: flex; flex-direction: column; gap: 6px; margin-top: 4px; }
  .checks label { display: flex; align-items: center; gap: 6px; height: 26px; padding: 0 6px; font-size: 12px; white-space: nowrap; }
  .notes .checks { gap: 0; margin-top: 3px; }
  .notes .checks label { height: 22px; margin-bottom: -3px; }
  .column { display: flex; flex-direction: column; }
  .small { display: flex; align-items: center; gap: 5px; height: 22px; padding: 0 6px; font-size: 12px; white-space: nowrap; }
  input { margin: 0; accent-color: var(--ok-accent); }
  .anchor { position: relative; display: flex; }
  .menu { position: fixed; z-index: 400; display: flex; flex-direction: column; min-width: 120px; padding: 4px 0; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .action { position: relative; display: flex; align-items: center; width: 100%; height: 24px; padding: 0 20px; border: 0; border-radius: 0; text-align: left; font-size: 12px; white-space: nowrap; }
  .action[aria-checked='true']::before { content: '✓'; position: absolute; left: 6px; }
</style>

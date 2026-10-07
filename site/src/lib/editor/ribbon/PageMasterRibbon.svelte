<script lang="ts">
  // Mac PowerPoint 16's Handout Master and Notes Master tabs: page orientation
  // and Slide Size | the placeholder checkboxes (Header, Footer, Date, Page
  // Number; the notes master adds Slide Image and Body) | the handout's slides
  // per page | Close Master. The library does not read or write either master
  // yet, so everything but Slide Size and Close Master is shown disabled.
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import { captionLines } from './caption.ts';
  import SlideSizeMenu from './SlideSizeMenu.svelte';

  let { kind }: { kind: 'handoutMaster' | 'notesMaster' } = $props();
  const editor = getEditor();
  const reason = $derived(t(kind === 'handoutMaster' ? 'Handout masters are not supported by the library yet.' : 'Notes masters are not supported by the library yet.'));
  // PowerPoint's defaults: every placeholder shown, six slides per handout page.
  const checks = $derived(kind === 'handoutMaster' ? [['Header', 'Footer'], ['Date', 'Page Number']] : [['Header', 'Slide Image', 'Footer'], ['Date', 'Body', 'Page Number']]);
  const PER_PAGE = [['2 Slides', '3 Slides', '4 Slides'], ['6 Slides', '9 Slides', 'Slide Outline']];
</script>

<div class="page-master-tab" class:notes={kind === 'notesMaster'}>
  <section class="cluster" role="group" aria-label={t('Page Setup')}>
    <button class="big" aria-label={t(kind === 'handoutMaster' ? 'Handout Orientation' : 'Notes Page Orientation')} style:--w={kind === 'handoutMaster' ? '63px' : '64px'} disabled title={reason}><span class="icon-row"><Icon name="layout" size={32} /><span class="arrow" aria-hidden="true">⌄</span></span><span class="caption">{captionLines(t(kind === 'handoutMaster' ? 'Handout Orientation' : 'Notes Page Orientation'))}</span></button>
    <SlideSizeMenu />
  </section>
  <section class="cluster" role="group" aria-label={t('Placeholders')}>
    {#each checks as column, i (i)}
      <div class="checks">
        {#each column as label (label)}<label title={reason}><input type="checkbox" checked disabled />{t(label)}</label>{/each}
      </div>
    {/each}
  </section>
  {#if kind === 'handoutMaster'}
    <section class="cluster" role="group" aria-label={t('Slides Per Page')}>
      {#each PER_PAGE as column, i (i)}
        <div class="column">
          {#each column as label (label)}<button class="small" aria-pressed={label === '6 Slides'} disabled title={reason}><Icon name="layout" size={16} />{t(label)}</button>{/each}
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
</style>

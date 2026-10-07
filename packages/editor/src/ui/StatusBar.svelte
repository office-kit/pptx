<script lang="ts">
  import { eventTarget } from '../core/dom-root.ts';
  import { getEditor } from '../core/context.ts';
  import { getLocale, t } from '../i18n/i18n.svelte.ts';
  import { accessibilityIssues } from '../core/accessibility.ts';
  import { isSlideEditingView } from '../core/view-modes.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const zoomPercent = $derived(Math.round(editor.viewZoom * 100));
  // Mac PowerPoint names the master being edited instead of a slide position,
  // and counts notes pages in Notes Page view.
  const MASTER_STATUS = { slideMaster: 'Slide Master', handoutMaster: 'Handout Master', notesMaster: 'Notes Master' } as const;
  const position = $derived.by(() => {
    const mode = editor.viewMode;
    if (mode === 'slideMaster' || mode === 'handoutMaster' || mode === 'notesMaster')
      return { text: t(MASTER_STATUS[mode]), help: t(`Currently in ${MASTER_STATUS[mode]} View.`) };
    const template = mode === 'notesPage' ? 'Notes {n} of {count}' : 'Slide {n} of {count}';
    return { text: t(template).replace('{n}', String(doc.selection.slideIndex + 1)).replace('{count}', String(doc.slides.length)), help: undefined };
  });
  const editing = $derived(isSlideEditingView(editor.viewMode));
  const sliderPosition = $derived(zoomPercent <= 100 ? (zoomPercent - editor.minZoomPercent) / (100 - editor.minZoomPercent) * 1000 : 1000 + (zoomPercent - 100) / (editor.maxZoomPercent - 100) * 1000);
  // Mac PowerPoint shows the proofing language, which follows the system's
  // language and region (e.g. "English (Japan)").
  const language = $derived.by(() => {
    if (typeof navigator === 'undefined') return '';
    try {
      return new Intl.DisplayNames([getLocale()], { type: 'language' }).of(navigator.language) ?? navigator.language;
    } catch {
      return navigator.language;
    }
  });
  const issues = $derived.by(() => { doc.version; return accessibilityIssues(doc.pres); });
  function goTo(issue: (typeof issues)[number]) {
    editor.accessibilityOpen = false;
    if (issue.kind === 'alt-text') doc.selectShape(issue.slide, issue.shapeId);
    else doc.selectSlide(issue.slide);
  }
  function slideZoom(event: Event) {
    const position = Number((event.currentTarget as HTMLInputElement).value);
    editor.setZoom(Math.round(position <= 1000 ? editor.minZoomPercent + position / 1000 * (100 - editor.minZoomPercent) : 100 + (position - 1000) / 1000 * (editor.maxZoomPercent - 100)) / 100);
  }
  function zoomKeys(event: KeyboardEvent) {
    if (!['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault(); event.stopPropagation();
    const amount = event.shiftKey ? 10 : 1;
    editor.setZoom((event.key === 'Home' ? editor.minZoomPercent : event.key === 'End' ? editor.maxZoomPercent : zoomPercent + (['ArrowLeft', 'ArrowDown'].includes(event.key) ? -amount : amount)) / 100);
  }

</script>

<svelte:window
  onkeydown={(event) => { if (editor.accessibilityOpen && event.key === 'Escape') editor.accessibilityOpen = false; }}
  onpointerdown={(event) => { if (editor.accessibilityOpen && !(eventTarget(event) as Element).closest?.('.a11y')) editor.accessibilityOpen = false; }}
/>

<div class="statusbar">
  <span class="position" title={position.help}>{position.text}</span>
  <span class="language">{language}</span>
  <div class="a11y">
    <button class="labelled" aria-haspopup="dialog" aria-expanded={editor.accessibilityOpen} onclick={() => (editor.accessibilityOpen = !editor.accessibilityOpen)}><svg width="14" height="14" viewBox="0 0 14 14" aria-hidden="true"><circle cx="7" cy="2.5" r="1.3"/><path d="M2 5h10M7 5v4M5 13l2-4 2 4"/></svg>{t(issues.length ? 'Accessibility: Investigate' : 'Accessibility: Good to go')}</button>
    {#if editor.accessibilityOpen}
      <div class="issues" role="dialog" aria-label={t('Accessibility')}>
        {#if issues.length === 0}<p>{t('No accessibility issues found.')}</p>{/if}
        {#each issues as issue (issue.kind + issue.slide + (issue.kind === 'alt-text' ? issue.shapeId : ''))}
          <button onclick={() => goTo(issue)}>
            {#if issue.kind === 'alt-text'}{t('Missing alternative text')}: {issue.name}{:else}{t('Missing slide title')}{/if}
            <span>{t('Slide {n}').replace('{n}', String(issue.slide + 1))}</span>
          </button>
        {/each}
      </div>
    {/if}
  </div>
  <span class="spacer"></span>
  <!-- Notes and Comments only exist where the slide editor does. -->
  {#if editing}
    <button class="labelled" aria-label={t('Notes')} aria-pressed={editor.notesVisible} onclick={() => { if (editor.notesVisible) editor.notesVisible = false; else editor.showNotes(); }}><svg width="14" height="14" viewBox="0 0 14 14" aria-hidden="true"><path d="M1.5 1.5h11v7l-4 4h-7z M8.5 12.5v-4h4"/></svg>{t('Notes')}</button>
    <button class="labelled" aria-label={t('Comments')} aria-pressed={editor.activeDialog === 'addSlideComment'} disabled={!doc.currentSlide} onclick={() => { if (editor.activeDialog === 'addSlideComment') editor.activeDialog = null; else editor.runOrPrompt('addSlideComment'); }}><svg width="15" height="14" viewBox="0 0 15 14" aria-hidden="true"><path d="M1.5 1.5h12v8h-7l-3 3v-3h-2z"/></svg>{t('Comments')}</button>
  {/if}
  <!-- Mac PowerPoint's 147 pt segmented switcher has exactly these four
       segments; Outline counts as Normal, and Notes Page and the master views
       select none. -->
  <div class="views" role="group" aria-label={t('Presentation views')}>
    <button title={t('Normal')} aria-label={t('Normal')} aria-pressed={editing} onclick={() => editor.setViewMode('normal')}><svg width="16" height="12" viewBox="0 0 16 12" aria-hidden="true"><rect x=".5" y=".5" width="15" height="11" rx="1"/><path d="M5 1v10"/></svg></button>
    <button title={t('Slide Sorter')} aria-label={t('Slide Sorter')} aria-pressed={editor.viewMode === 'sorter'} onclick={() => editor.setViewMode('sorter')}><svg width="16" height="12" viewBox="0 0 16 12" aria-hidden="true"><path d="M1 1h5v4H1zM9 1h5v4H9zM1 7h5v4H1zM9 7h5v4H9z"/></svg></button>
    <button title={t('Reading View')} aria-label={t('Reading View')} aria-pressed={editor.readingView} onclick={() => editor.openReadingView()}><svg width="16" height="12" viewBox="0 0 16 12" aria-hidden="true"><path d="M1 1.5h6v9H1zM9 1.5h6v9H9z"/></svg></button>
    <button title={t('Slide Show')} aria-label={t('Slide Show')} aria-pressed="false" disabled={!editor.canPresent} onclick={() => editor.present('current')}><svg width="16" height="12" viewBox="0 0 16 12" aria-hidden="true"><path d="M1 1h14v7H1zM8 8v3M5 11h6M6.5 3l3 1.5-3 1.5z"/></svg></button>
  </div>
  <div class="zoom">
    <button class="zbtn" title={t('Zoom out (Ctrl+-)')} onclick={() => editor.zoomOut()}>−</button>
    <input type="range" min="0" max="2000" step="1" value={sliderPosition} aria-label={t('Zoom percentage')} aria-valuetext="{zoomPercent}%" oninput={slideZoom} onkeydown={zoomKeys} />
    <button class="zbtn" title={t('Zoom in (Ctrl+=)')} onclick={() => editor.zoomIn()}>+</button>
    <button class="zpct" title={t('Zoom...')} onclick={() => editor.activeDialog = 'zoom'}>{zoomPercent}%</button>
    <button class="zfit" title={t('Fit (Ctrl+0)')} aria-label={t('Fit slide to current window')} onclick={() => editor.zoomFit()}><svg width="14" height="14" viewBox="0 0 14 14" aria-hidden="true"><path d="M1 5V1h4M9 1h4v4M13 9v4H9M5 13H1V9M1 1l4 4M13 1L9 5M13 13L9 9M1 13l4-4"/></svg></button>
  </div>
</div>

<style>
  /* Mac PowerPoint's status bar sits on the window chrome color, not the
     accent: 28 pt tall, 12 pt text, 36 pt view buttons, a 102 pt zoom slider
     between 14 pt − / + buttons, a 54 pt percentage and a 38 pt Fit button. */
  .statusbar {
    display: flex;
    align-items: center;
    gap: 2px;
    height: 28px;
    padding: 0 13px;
    background: var(--ok-ribbon);
    border-top: 1px solid var(--ok-border);
    color: var(--ok-text-2);
    font-size: 12px;
  }
  .spacer { flex: 1; }
  .language { padding: 0 6px; white-space: nowrap; }
  .a11y { position: relative; }
  .issues { position: absolute; bottom: calc(100% + 6px); left: 0; z-index: 300; display: flex; flex-direction: column; min-width: 280px; max-height: 50vh; overflow-y: auto; padding: 6px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .issues button { justify-content: space-between; gap: 12px; padding: 4px 8px; text-align: left; }
  .issues p { margin: 4px 8px; }
  .zoom { display: flex; align-items: center; gap: 2px; }
  .zoom input[type='range'] { width: 102px; height: 12px; accent-color: var(--ok-text-2); margin: 0 3px; }
  .position { padding-right: 5px; white-space: nowrap; }
  .views { display: flex; gap: 0; width: 147px; margin: 0 3px 0 0; }
  .views button { justify-content: center; width: 36px; height: 24px; padding: 0; }
  .views button:first-child { width: 37px; }
  .labelled { padding: 2px 12px; }
  svg { fill: none; stroke: currentColor; stroke-width: 1.1; }
  button {
    display: inline-flex;
    align-items: center;
    gap: 5px;
    background: transparent;
    border: none;
    color: inherit;
    font: inherit;
    font-size: 12px;
    cursor: pointer;
    padding: 2px 6px;
    border-radius: 3px;
  }
  button:hover:not(:disabled) { background: var(--ok-hover); }
  button:disabled { opacity: 0.45; cursor: default; }
  button[aria-pressed='true'] { background: var(--ok-selected); color: var(--ok-text); }
  .zbtn { font-size: 15px; line-height: 1; width: 14px; padding: 2px 0; justify-content: center; }
  .zpct { min-width: 54px; justify-content: center; }
  .zfit { justify-content: center; width: 38px; }
</style>

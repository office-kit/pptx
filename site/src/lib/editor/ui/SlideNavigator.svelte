<script lang="ts">
  // Left rail: slide thumbnails (rendered with the same preview renderer as the
  // canvas) plus slide-level operations. Selecting a thumbnail sets the active
  // slide; the buttons dispatch through the controller so they participate in
  // undo/redo like everything else.
  import { getEditor } from '../core/context.ts';
  import { renderSlideToSvg } from '@office-kit/pptx-preview';
  import { getSlideSize, isSlideHidden } from '@office-kit/pptx';
  import { tick } from 'svelte';
  import { selectedSlideIndices } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let { mode = 'normal' }: { mode?: 'normal' | 'sorter' } = $props();
  const editor = getEditor();
  const doc = editor.doc;
  const selected = $derived(selectedSlideIndices(doc.selection));
  const firstSelected = $derived(selected[0] ?? 0);

  const skippedSlides = $derived.by(() => { doc.version; return doc.slides.map(isSlideHidden); });
  const size = $derived.by(() => { doc.version; return getSlideSize(doc.pres); });

  function thumb(index: number): { svg: string; error: string } {
    doc.version; // reactive
    const slide = doc.slideAt(index);
    if (!slide) return { svg: '', error: '' };
    try {
      return { svg: renderSlideToSvg(doc.pres, slide), error: '' };
    } catch (error) {
      return { svg: '', error: error instanceof Error ? error.message : String(error) };
    }
  }

  async function focusSlide(index: number, range = false, preserve = false) {
    if (!preserve) doc.selectSlide(index, { range });
    await tick();
    rail?.querySelector<HTMLElement>(`[data-slide-index="${doc.selection.slideIndex}"]`)?.focus();
  }

  function reorder(from: number, to: number) {
    if ((selected.includes(from) ? firstSelected : from) === to || to < 0 || to >= doc.slides.length) return;
    if (doc.selection.kind !== 'slide' || !selected.includes(from)) doc.selectSlide(from);
    editor.invoke('moveSlide', { toIndex: to });
    void focusSlide(doc.selection.slideIndex, false, true);
  }

  function onKeydown(event: KeyboardEvent, index: number) {
    if (event.isComposing) return;
    const mod = event.ctrlKey || event.metaKey;
    if (mod && event.key.toLowerCase() === 'a') {
      event.preventDefault();
      const indices = doc.slides.map((_, i) => i);
      doc.select({ kind: 'slide', slideIndex: index, slideIndices: indices, anchorIndex: index });
      return;
    }
    if (mod && event.key.toLowerCase() === 'd') {
      event.preventDefault();
      editor.invoke('duplicateSlide');
      void focusSlide(doc.selection.slideIndex, false, true);
      return;
    }
    if (event.key === 'Delete' || event.key === 'Backspace') {
      event.preventDefault();
      editor.invoke('removeSlide');
      void focusSlide(doc.selection.slideIndex, false, true);
      return;
    }
    const thumbs = [...rail.querySelectorAll<HTMLElement>('[data-slide-index]')];
    const columns = mode === 'sorter' && thumbs.length ? thumbs.filter(item => item.offsetTop === thumbs[0]!.offsetTop).length : 1;
    const delta = event.key === 'ArrowUp' ? -columns : event.key === 'ArrowDown' ? columns : mode === 'sorter' && event.key === 'ArrowLeft' ? -1 : mode === 'sorter' && event.key === 'ArrowRight' ? 1 : 0;
    if (delta) {
      event.preventDefault();
      if (event.altKey) reorder(index, firstSelected + delta);
      else void focusSlide(index + delta, event.shiftKey);
    } else if (event.key === 'Home' || event.key === 'End') {
      event.preventDefault();
      void focusSlide(event.key === 'Home' ? 0 : doc.slides.length - 1, event.shiftKey);
    } else if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      doc.selectSlide(index);
      if (mode === 'sorter' && event.key === 'Enter') editor.setViewMode('normal');
    }
  }

  let pointerSelecting = false;
  let rail: HTMLDivElement;
  let dragIndex = $state<number | null>(null);

  // Keep keyboard focus aligned when history restores a different selection.
  $effect(() => {
    const selection = doc.selection;
    if (selection.kind === 'slide' && rail?.contains(document.activeElement) && document.activeElement?.classList.contains('thumb-row')) {
      rail.querySelector<HTMLElement>(`[data-slide-index="${selection.slideIndex}"]`)?.focus();
    }
  });
</script>

<svelte:window onpointerup={() => pointerSelecting = false} onpointercancel={() => pointerSelecting = false} />

<div class="nav ok-scroll" class:sorter={mode === 'sorter'} style:--sorter-thumb={`${Math.round(230 * editor.sorterZoom)}px`} bind:this={rail}>
  <div class="nav-actions">
    <button class="ok-btn add" onclick={() => editor.invoke('addBlankSlide')} title={t('New slide')}>＋ {t('Slide')}</button>
    <button class="ok-btn add from-layout" onclick={() => editor.runOrPrompt('addSlide')} title={t('New slide from layout')}>{t('New slide from layout')}</button>
    <div class="slide-actions">
      <button class="ok-btn" title={t('Duplicate slide')} aria-label={t('Duplicate slide')} disabled={!doc.slides.length} onclick={() => editor.invoke('duplicateSlide')}>⧉</button>
      <button class="ok-btn" title={t('Delete slide')} aria-label={t('Delete slide')} disabled={!doc.slides.length} onclick={() => editor.invoke('removeSlide')}>×</button>
      <button class="ok-btn" title={t('Move slide up')} aria-label={t('Move slide up')} disabled={firstSelected === 0} onclick={() => reorder(doc.selection.slideIndex, firstSelected - 1)}>↑</button>
      <button class="ok-btn" title={t('Move slide down')} aria-label={t('Move slide down')} disabled={firstSelected >= doc.slides.length - selected.length} onclick={() => reorder(doc.selection.slideIndex, firstSelected + 1)}>↓</button>
    </div>
  </div>

  {#if doc.selection.kind === 'slide' && selected.length > 1}<div class="selection-count" aria-live="polite">{t('Selected slides')}: {selected.length}</div>{/if}

  {#each doc.slides as _slide, i (i)}
    {@const preview = thumb(i)}
    <div
      class="thumb-row"
      class:active={doc.selection.kind === 'slide' ? selected.includes(i) : doc.selection.slideIndex === i}
      class:skipped={skippedSlides[i]}
      draggable="true"
      role="button"
      data-slide-index={i}
      aria-label={`${t('Slide')} ${i + 1}`}
      title={skippedSlides[i] ? t('Skipped during presentation') : undefined}
      aria-current={doc.selection.slideIndex === i ? 'true' : undefined}
      tabindex={doc.selection.slideIndex === i ? 0 : -1}
      aria-pressed={doc.selection.kind === 'slide' && selected.includes(i)}
      onpointerdown={() => pointerSelecting = true}
      onfocus={() => { if (!pointerSelecting && (doc.selection.kind !== 'slide' || !selected.includes(i))) doc.selectSlide(i); }}
      ondragstart={(event) => {
        dragIndex = i;
        if (doc.selection.kind !== 'slide' || !selected.includes(i)) doc.selectSlide(i);
        if (event.dataTransfer) {
          event.dataTransfer.effectAllowed = 'move';
          event.dataTransfer.setData('text/plain', String(i));
        }
      }}
      ondragend={() => { dragIndex = null; pointerSelecting = false; }}
      ondragover={(e) => e.preventDefault()}
      ondrop={(event) => {
        event.preventDefault();
        if (dragIndex !== null && dragIndex !== i) reorder(dragIndex, i);
        dragIndex = null;
      }}
      oncontextmenu={(event) => {
        event.preventDefault();
        if (doc.selection.kind !== 'slide' || !selected.includes(i)) doc.selectSlide(i);
        editor.openContextMenu(event.clientX, event.clientY);
      }}
      onclick={(event) => {
        doc.selectSlide(i, { additive: event.ctrlKey || event.metaKey, range: event.shiftKey });
        void focusSlide(doc.selection.slideIndex, false, true);
      }}
      ondblclick={() => { doc.selectSlide(i); editor.setViewMode('normal'); }}
      onkeydown={(event) => onKeydown(event, i)}
    >
      <span class="num">{i + 1}{#if skippedSlides[i]}<span class="skip-mark" aria-label={t('Skipped during presentation')}>⊘</span>{/if}</span>
      <div class="thumb" style:aspect-ratio={size ? `${size.width} / ${size.height}` : '16 / 9'}>
        {#if preview.error}<span class="render-error" title={preview.error}>{t('Preview unavailable')}</span>{:else}{@html preview.svg}{/if}
      </div>
    </div>
  {/each}
</div>

<style>
  .sorter { display: grid; grid-template-columns: repeat(auto-fill, minmax(min(100%, var(--sorter-thumb)), 1fr)); align-content: start; gap: 20px; padding: 24px; border-right: 0; }
  .sorter .nav-actions, .sorter .selection-count { grid-column: 1 / -1; }
  .sorter .nav-actions { display: flex; gap: 8px; align-items: center; }
  .sorter .nav-actions .add { width: auto; margin: 0; }
  .sorter .slide-actions { margin: 0; }
  .sorter .thumb-row { min-width: 0; }
  .selection-count { padding: 6px 4px; font-size: 11px; color: var(--ok-muted); }
  .skipped .num { text-decoration: line-through; }
  .skip-mark { display: block; text-decoration: none; }
  .skipped .thumb { opacity: .6; }
  .nav {
    background: var(--ok-panel-2);
    border-right: 1px solid var(--ok-border);
    padding: 8px;
    overflow-y: auto;
  }
  .nav-actions {
    margin-bottom: 8px;
  }
  .from-layout { margin-top: 6px; white-space: normal; }
  .slide-actions { display: flex; gap: 2px; margin-top: 6px; }
  .slide-actions button { flex: 1; justify-content: center; }
  .render-error { font-size: 11px; color: var(--ok-text-2); }
  .thumb-row:focus-visible { outline: 2px solid var(--ok-selected-border); outline-offset: 1px; }
  .add {
    width: 100%;
    justify-content: center;
    border: 1px solid var(--ok-border);
    background: var(--ok-panel);
  }
  .thumb-row {
    display: flex;
    gap: 6px;
    align-items: flex-start;
    padding: 4px;
    border-radius: var(--ok-radius);
    cursor: pointer;
  }
  .thumb-row:hover {
    background: var(--ok-hover);
  }
  .thumb-row.active {
    background: var(--ok-selected);
  }
  .num {
    font-size: 11px;
    color: var(--ok-text-2);
    width: 16px;
    text-align: right;
    padding-top: 2px;
  }
  .thumb {
    flex: 1;
    background: #fff;
    border: 1px solid var(--ok-border-strong);
    border-radius: 2px;
    overflow: hidden;
  }
  .thumb-row.active .thumb {
    border-color: var(--ok-selected-border);
    box-shadow: 0 0 0 1px var(--ok-selected-border);
  }
  .thumb :global(svg) {
    width: 100%;
    height: 100%;
    display: block;
  }
</style>

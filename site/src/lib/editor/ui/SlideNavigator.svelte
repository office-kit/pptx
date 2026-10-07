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
  import { sectionRanges } from '../core/sections.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let { mode = 'normal' }: { mode?: 'normal' | 'sorter' } = $props();
  // Mac PowerPoint's Slide Sorter at its default 80%: 200 pt thumbnails
  // (border included) in 206 × 147 pt cells, a 39 pt gutter (245 pt pitch), and
  // as many columns as fit the list — floor((width + 39) / 245), six in a
  // 1512 pt window — centred. The thumbnail scales with the zoom; the gutter
  // and the slide number's 31 pt line below do not.
  const SORTER_THUMB_AT_100 = 250;
  const editor = getEditor();
  const doc = editor.doc;
  const sectionStarts = $derived.by(() => { doc.version; return new Map(sectionRanges(doc.pres).map(range => [range.start, range.name])); });
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

<div class="nav ok-scroll" class:sorter={mode === 'sorter'} style:--sorter-thumb={`${Math.round(SORTER_THUMB_AT_100 * editor.sorterZoom)}px`} bind:this={rail}>
  {#if doc.selection.kind === 'slide' && selected.length > 1}<div class="selection-count" aria-live="polite">{t('Selected slides')}: {selected.length}</div>{/if}

  {#each doc.slides as _slide, i (i)}
    {@const preview = thumb(i)}
    {#if sectionStarts.has(i)}<div class="section-header" role="heading" aria-level="2">{sectionStarts.get(i)}</div>{/if}
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
  .nav.sorter { display: grid; grid-template-columns: repeat(auto-fill, calc(var(--sorter-thumb) + 6px)); justify-content: center; align-content: start; column-gap: 39px; row-gap: 24px; padding: 5px 0 24px; border-right: 0; background: var(--ok-canvas-bg); scrollbar-gutter: stable; }
  .sorter .thumb-row { flex-direction: column; gap: 0; min-width: 0; padding: 3px 3px 0; }
  .sorter .thumb { flex: none; box-sizing: border-box; width: var(--sorter-thumb); }
  .sorter .num { order: 1; box-sizing: border-box; width: auto; height: 31px; padding-top: 11px; font-size: 13px; line-height: 16px; text-align: left; }
  .sorter .selection-count { grid-column: 1 / -1; }
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
  .section-header { grid-column: 1 / -1; padding: 6px 4px 2px; font-size: 11px; font-weight: 600; color: var(--ok-text-2); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .render-error { font-size: 11px; color: var(--ok-text-2); }
  .thumb-row:focus-visible { outline: 2px solid var(--ok-selected-border); outline-offset: 1px; }
  .thumb-row {
    display: flex;
    gap: 6px;
    align-items: flex-start;
    padding: 6px 6px 8px 4px;
    cursor: pointer;
  }
  /* Mac PowerPoint frames the selected thumbnail with a rounded accent ring
     (and hovered ones with a gray ring) instead of tinting the row. */
  .thumb-row:hover .thumb {
    outline: 3px solid var(--ok-border-strong);
    outline-offset: 2px;
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
    border: 1px solid var(--ok-border);
    border-radius: 4px;
    overflow: hidden;
  }
  .thumb-row.active .thumb {
    outline: 3px solid var(--ok-accent);
    outline-offset: 2px;
  }
  .thumb :global(svg) {
    width: 100%;
    height: 100%;
    display: block;
  }
</style>

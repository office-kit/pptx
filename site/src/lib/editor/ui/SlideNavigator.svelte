<script lang="ts">
  // Left rail: slide thumbnails (rendered with the same preview renderer as the
  // canvas) plus slide-level operations. Selecting a thumbnail sets the active
  // slide; the buttons dispatch through the controller so they participate in
  // undo/redo like everything else.
  import { getEditor } from '../core/context.ts';
  import { renderSlideToSvg } from '@office-kit/pptx-preview';
  import { getSlideSize } from '@office-kit/pptx';
  import { tick } from 'svelte';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;

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

  async function focusSlide(index: number) {
    doc.selectSlide(index);
    await tick();
    rail?.querySelector<HTMLElement>(`[data-slide-index="${doc.selection.slideIndex}"]`)?.focus();
  }

  function reorder(from: number, to: number) {
    if (from === to || to < 0 || to >= doc.slides.length) return;
    doc.selectSlide(from);
    editor.invoke('moveSlide', { toIndex: to });
    void focusSlide(to);
  }

  function onKeydown(event: KeyboardEvent, index: number) {
    if (event.isComposing) return;
    const mod = event.ctrlKey || event.metaKey;
    if (mod && event.key.toLowerCase() === 'd') {
      event.preventDefault();
      editor.invoke('duplicateSlide');
      void focusSlide(doc.selection.slideIndex);
      return;
    }
    if (event.key === 'Delete' || event.key === 'Backspace') {
      event.preventDefault();
      editor.invoke('removeSlide');
      void focusSlide(doc.selection.slideIndex);
      return;
    }
    const delta = event.key === 'ArrowUp' ? -1 : event.key === 'ArrowDown' ? 1 : 0;
    if (delta) {
      event.preventDefault();
      if (event.altKey) reorder(index, index + delta);
      else void focusSlide(index + delta);
    } else if (event.key === 'Home' || event.key === 'End') {
      event.preventDefault();
      void focusSlide(event.key === 'Home' ? 0 : doc.slides.length - 1);
    } else if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      doc.selectSlide(index);
    }
  }

  let rail: HTMLDivElement;
  let dragIndex = $state<number | null>(null);
</script>

<div class="nav ok-scroll" bind:this={rail}>
  <div class="nav-actions">
    <button class="ok-btn add" onclick={() => editor.invoke('addBlankSlide')} title={t('New slide')}>＋ {t('Slide')}</button>
    <div class="slide-actions">
      <button class="ok-btn" title={t('Duplicate slide')} aria-label={t('Duplicate slide')} disabled={!doc.slides.length} onclick={() => editor.invoke('duplicateSlide')}>⧉</button>
      <button class="ok-btn" title={t('Delete slide')} aria-label={t('Delete slide')} disabled={!doc.slides.length} onclick={() => editor.invoke('removeSlide')}>×</button>
      <button class="ok-btn" title={t('Move slide up')} aria-label={t('Move slide up')} disabled={doc.selection.slideIndex === 0} onclick={() => reorder(doc.selection.slideIndex, doc.selection.slideIndex - 1)}>↑</button>
      <button class="ok-btn" title={t('Move slide down')} aria-label={t('Move slide down')} disabled={doc.selection.slideIndex >= doc.slides.length - 1} onclick={() => reorder(doc.selection.slideIndex, doc.selection.slideIndex + 1)}>↓</button>
    </div>
  </div>

  {#each doc.slides as _slide, i (i)}
    {@const preview = thumb(i)}
    <div
      class="thumb-row"
      class:active={doc.selection.slideIndex === i}
      draggable="true"
      role="button"
      data-slide-index={i}
      aria-label={`${t('Slide')} ${i + 1}`}
      aria-current={doc.selection.slideIndex === i ? 'true' : undefined}
      tabindex={doc.selection.slideIndex === i ? 0 : -1}
      onfocus={() => doc.selectSlide(i)}
      ondragstart={(event) => {
        dragIndex = i;
        doc.selectSlide(i);
        if (event.dataTransfer) {
          event.dataTransfer.effectAllowed = 'move';
          event.dataTransfer.setData('text/plain', String(i));
        }
      }}
      ondragend={() => dragIndex = null}
      ondragover={(e) => e.preventDefault()}
      ondrop={(event) => {
        event.preventDefault();
        if (dragIndex !== null && dragIndex !== i) reorder(dragIndex, i);
        dragIndex = null;
      }}
      oncontextmenu={(event) => {
        event.preventDefault();
        doc.selectSlide(i);
        editor.openContextMenu(event.clientX, event.clientY);
      }}
      onclick={() => doc.selectSlide(i)}
      onkeydown={(event) => onKeydown(event, i)}
    >
      <span class="num">{i + 1}</span>
      <div class="thumb" style:aspect-ratio={size ? `${size.width} / ${size.height}` : '16 / 9'}>
        {#if preview.error}<span class="render-error" title={preview.error}>{t('Preview unavailable')}</span>{:else}{@html preview.svg}{/if}
      </div>
    </div>
  {/each}
</div>

<style>
  .nav {
    background: var(--ok-panel-2);
    border-right: 1px solid var(--ok-border);
    padding: 8px;
    overflow-y: auto;
  }
  .nav-actions {
    margin-bottom: 8px;
  }
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

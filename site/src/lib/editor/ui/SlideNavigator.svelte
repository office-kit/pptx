<script lang="ts">
  // Left rail: slide thumbnails (rendered with the same preview renderer as the
  // canvas) plus slide-level operations. Selecting a thumbnail sets the active
  // slide; the buttons dispatch through the controller so they participate in
  // undo/redo like everything else.
  import { getEditor } from '../core/context.ts';
  import { renderSlideToSvg } from '@office-kit/pptx-preview';
  import { moveSlide } from '@office-kit/pptx';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;

  function thumb(index: number): string {
    doc.version; // reactive
    const slide = doc.slideAt(index);
    if (!slide) return '';
    try {
      return renderSlideToSvg(doc.pres, slide);
    } catch {
      return '';
    }
  }

  function reorder(from: number, to: number) {
    const slide = doc.slideAt(from);
    if (!slide || to < 0 || to >= doc.slides.length) return;
    doc.transact(t('Move slide'), () => moveSlide(doc.pres, slide, to));
    doc.selectSlide(to);
  }

  let dragIndex = $state<number | null>(null);
</script>

<div class="nav ok-scroll">
  <div class="nav-actions">
    <button class="ok-btn add" onclick={() => editor.invoke('addBlankSlide')} title={t('New slide')}>＋ {t('Slide')}</button>
  </div>

  {#each doc.slides as _slide, i (i)}
    <div
      class="thumb-row"
      class:active={doc.selection.slideIndex === i}
      draggable="true"
      role="button"
      tabindex="0"
      ondragstart={() => (dragIndex = i)}
      ondragover={(e) => e.preventDefault()}
      ondrop={() => {
        if (dragIndex !== null && dragIndex !== i) reorder(dragIndex, i);
        dragIndex = null;
      }}
      onclick={() => doc.selectSlide(i)}
      onkeydown={(e) => e.key === 'Enter' && doc.selectSlide(i)}
    >
      <span class="num">{i + 1}</span>
      <div class="thumb">
        {@html thumb(i)}
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
    aspect-ratio: 16 / 9;
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

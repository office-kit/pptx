<script lang="ts">
  import { tick } from 'svelte';
  import { getEditor } from '../core/context.ts';
  import { outlineShapes } from '../core/outline.ts';
  import { selectedSlideIndices } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import OutlineText from './OutlineText.svelte';
  const editor = getEditor();
  const doc = editor.doc;
  const selected = $derived(selectedSlideIndices(doc.selection));
  const entries = $derived.by(() => { doc.version; return doc.slides.map(outlineShapes); });
  let pane: HTMLElement;
  async function keys(event: KeyboardEvent, index: number) {
    if (event.isComposing) return;
    if (event.key === 'ArrowUp' || event.key === 'ArrowDown' || event.key === 'Home' || event.key === 'End') {
      event.preventDefault(); event.stopPropagation();
      doc.selectSlide(event.key === 'Home' ? 0 : event.key === 'End' ? doc.slides.length - 1 : index + (event.key === 'ArrowDown' ? 1 : -1), { range: event.shiftKey });
      await tick();
      pane.querySelector<HTMLButtonElement>(`[data-outline-slide="${doc.selection.slideIndex}"] button`)?.focus();
    } else if (event.key === 'Enter') {
      event.preventDefault();
      pane.querySelector<HTMLTextAreaElement>(`[data-outline-slide="${index}"] textarea`)?.focus();
    }
  }
</script>

<nav class="outline-pane" bind:this={pane} aria-label={t('Outline View')}>
  {#key doc.pres}
    {#each entries as shapes, index (doc.slides[index])}
      <div class="outline-slide" data-outline-slide={index}>
        <button class:selected={selected.includes(index)} aria-label={`${t('Slide')} ${index + 1}`} aria-pressed={selected.includes(index)} onclick={event => doc.selectSlide(index, { additive: event.metaKey || event.ctrlKey, range: event.shiftKey })} onkeydown={event => keys(event, index)} oncontextmenu={event => { event.preventDefault(); if (!selected.includes(index)) doc.selectSlide(index); editor.openContextMenu(event.clientX, event.clientY); }}><span>{index + 1}</span><svg viewBox="0 0 20 16" aria-hidden="true"><rect x="1.5" y="1.5" width="17" height="13" /></svg></button>
        <div class="text">
          {#each shapes as shape (shape.id)}<OutlineText slideIndex={index} shapeId={shape.id} title={shape.title} />{/each}
        </div>
      </div>
    {/each}
  {/key}
</nav>

<style>
  .outline-pane { overflow: auto; min-height: 0; padding: 12px 5px; background: var(--ok-panel); border-right: 1px solid var(--ok-border); }
  .outline-slide { display: grid; grid-template-columns: 43px minmax(0, 1fr); gap: 2px; margin-bottom: 14px; min-height: 24px; }
  button { display: flex; align-items: center; gap: 3px; align-self: start; padding: 2px 0; background: transparent; color: var(--ok-text); border: 0; font: 12px Arial, sans-serif; cursor: pointer; }
  button span { min-width: 17px; text-align: right; }
  svg { width: 20px; height: 16px; fill: var(--ok-bg); stroke: var(--ok-text-2); stroke-width: 2; }
  button.selected svg { stroke: var(--ok-accent); stroke-width: 3; }
  button:focus-visible { outline: 2px solid var(--ok-accent); }
  .text { min-width: 0; }
</style>

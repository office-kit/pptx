<script lang="ts">
  import { onDestroy, onMount, tick } from 'svelte';
  import { getCollapsedOutlineSlides, setSlideOutlineCollapsed, setParagraphAlignment } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { outlineShapes } from '../core/outline.ts';
  import { selectedSlideIndices } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import OutlineText from './OutlineText.svelte';
  import { OutlineSelectionModel } from '../core/outline-selection.ts';
  import { toggleTextFormat, type TextFormatToggle } from '../core/text-format-toggle.ts';
  import type { TextCase, TextFormat } from '@office-kit/pptx';
  const editor = getEditor();
  const doc = editor.doc;
  const selection = new OutlineSelectionModel();
  let selectionRevision = $state(0);
  const unsubscribeSelection = selection.subscribe(() => selectionRevision++);
  onDestroy(unsubscribeSelection);
  onDestroy(() => selection.clear());
  const selected = $derived(selectedSlideIndices(doc.selection));
  const entries = $derived.by(() => { doc.version; const collapsed = new Set(getCollapsedOutlineSlides(doc.pres)); return doc.slides.map(slide => ({shapes: outlineShapes(slide), collapsed: collapsed.has(slide)})); });
  let pane: HTMLElement;
  onMount(() => {
    const ownerDocument = pane.ownerDocument;
    const leaveOutline = (event: Event) => {
      if (event.target instanceof Element && event.target.closest('.canvas-viewport')) {
        const active = ownerDocument.activeElement;
        if (active instanceof HTMLElement && pane.contains(active)) active.blur();
        selection.clear();
      }
    };
    ownerDocument.addEventListener('pointerdown', leaveOutline, true);
    ownerDocument.addEventListener('focusin', leaveOutline, true);
    return () => {
      ownerDocument.removeEventListener('pointerdown', leaveOutline, true);
      ownerDocument.removeEventListener('focusin', leaveOutline, true);
    };
  });
  let dragging: { presentation: typeof doc.pres; version: number; indices: number[] } | null = null;
  let insertion = $state<{ index: number; after: boolean } | null>(null);
  function endDrag() { dragging = null; insertion = null; }
  function dragStart(event: DragEvent, index: number) {
    selection.clear();
    if (doc.selection.kind !== 'slide' || !selected.includes(index)) doc.selectSlide(index);
    dragging = { presentation: doc.pres, version: doc.version, indices: selectedSlideIndices(doc.selection) };
    if (event.dataTransfer) {
      event.dataTransfer.effectAllowed = 'move';
      event.dataTransfer.setData('text/plain', dragging.indices.map(i => String(i + 1)).join(', '));
    }
  }
  function dragOver(event: DragEvent, index: number) {
    if (!dragging || dragging.presentation !== doc.pres || dragging.version !== doc.version || dragging.indices.includes(index)) { insertion = null; return; }
    event.preventDefault();
    if (event.dataTransfer) event.dataTransfer.dropEffect = 'move';
    const bounds = (event.currentTarget as HTMLElement).getBoundingClientRect();
    insertion = { index, after: event.clientY > bounds.top + bounds.height / 2 };
  }
  async function drop(event: DragEvent, index: number) {
    if (!dragging || !insertion || insertion.index !== index || dragging.presentation !== doc.pres || dragging.version !== doc.version) { endDrag(); return; }
    event.preventDefault(); event.stopPropagation();
    const indices = dragging.indices;
    const boundary = index + (insertion.after ? 1 : 0);
    const destination = boundary - indices.filter(i => i < boundary).length;
    endDrag();
    if (destination === indices[0] && indices.every((value, i) => value === destination + i)) return;
    doc.select({ kind: 'slide', slideIndex: indices[0]!, slideIndices: indices });
    editor.invoke('moveSlide', { toIndex: destination });
    await tick();
    pane.querySelector<HTMLButtonElement>(`[data-outline-slide="${doc.selection.slideIndex}"] > button`)?.focus();
  }
  function toggleCollapse(index: number) {
    const slide = doc.slideAt(index)!;
    const collapsed = entries[index].collapsed;
    doc.transact(t(collapsed ? 'Expand' : 'Collapse'), () => setSlideOutlineCollapsed(slide, !collapsed));
  }

  $effect(() => {
    selectionRevision;
    const range = selection.current();
    if (!range) return;
    const api = {
      formats: selection.formats(),
      paragraphs: selection.paragraphs(),
      editParagraphs: (edit: (shape: Parameters<typeof setParagraphAlignment>[0], index: number) => void) => selection.editParagraphs(edit, t('Format paragraphs')),
      alignment: (() => {
        const values = new Set(selection.paragraphs().map(paragraph => paragraph.align).filter(Boolean));
        return values.size === 1 ? [...values][0]! : '';
      })(),
      align: (value: string) => {
        if (value === 'left' || value === 'center' || value === 'right' || value === 'justify' || value === 'distribute')
          selection.editParagraphs((shape, index) => setParagraphAlignment(shape, index, value), t('Format paragraphs'));
      },
      apply: (format: TextFormat, reset = false) => selection.format(format, reset, t('Format selected text')),
      fontSize: (direction: 1 | -1) => selection.fontSize(direction, t(direction > 0 ? 'Increase Font Size' : 'Decrease Font Size')),
      changeCase: (value: TextCase) => selection.changeCase(value, t('Change Case')),
      toggle: (property: TextFormatToggle) => selection.format(formats => toggleTextFormat(formats, property), false, t('Format selected text')),
    };
    editor.inlineTextFormat = api;
    return () => {
      // $state proxies the API object; its callback retains its identity.
      if (editor.inlineTextFormat?.apply === api.apply) editor.inlineTextFormat = null;
    };
  });
  async function keys(event: KeyboardEvent, index: number) {
    if (event.isComposing) return;
    if (event.key === 'ArrowUp' || event.key === 'ArrowDown' || event.key === 'Home' || event.key === 'End') {
      event.preventDefault(); event.stopPropagation();
      selection.clear();
      doc.selectSlide(event.key === 'Home' ? 0 : event.key === 'End' ? doc.slides.length - 1 : index + (event.key === 'ArrowDown' ? 1 : -1), { range: event.shiftKey });
      await tick();
      pane.querySelector<HTMLButtonElement>(`[data-outline-slide="${doc.selection.slideIndex}"] button`)?.focus();
    } else if (event.key === 'Enter') {
      event.preventDefault();
      pane.querySelector<HTMLElement>(`[data-outline-slide="${index}"] [role="textbox"]`)?.focus();
    }
  }
</script>

<svelte:window ondragend={endDrag} ondrop={endDrag} onblur={endDrag} />

<nav class="outline-pane" bind:this={pane} aria-label={t('Outline View')}>
  {#key doc.pres}
    {#each entries as entry, index (doc.slides[index])}
      <div class="outline-slide" role="group" aria-label={`${t('Slide')} ${index + 1}`} class:insert-before={insertion?.index === index && !insertion.after} class:insert-after={insertion?.index === index && insertion.after} data-outline-slide={index} ondragover={event => dragOver(event, index)} ondragleave={() => insertion = null} ondrop={event => { void drop(event, index); }}>
        <button draggable="true" ondragstart={event => dragStart(event, index)} class:selected={selected.includes(index)} aria-label={`${t('Slide')} ${index + 1}`} aria-pressed={selected.includes(index)} aria-expanded={!entry.collapsed} ondblclick={() => toggleCollapse(index)} onclick={event => { selection.clear(); doc.selectSlide(index, { additive: event.metaKey || event.ctrlKey, range: event.shiftKey }); }} onkeydown={event => keys(event, index)} oncontextmenu={event => { event.preventDefault(); if (doc.selection.kind !== 'slide' || !selected.includes(index)) { selection.clear(); doc.selectSlide(index); } editor.openContextMenu(event.clientX, event.clientY, { source: 'outline' }); }}><span>{index + 1}</span><svg viewBox="0 0 20 16" aria-hidden="true"><rect x="1.5" y="1.5" width="17" height="13" /></svg></button>
        <div class="text">
          {#each entry.shapes as shape (shape.id)}{#if shape.title || !entry.collapsed}<OutlineText slideIndex={index} shapeId={shape.id} title={shape.title} {selection} />{/if}{/each}
        </div>
      </div>
    {/each}
  {/key}
</nav>

<style>
  .outline-pane { overflow: auto; min-height: 0; padding: 12px 5px; background: var(--ok-panel); border-right: 1px solid var(--ok-border); }
  .outline-slide { display: grid; grid-template-columns: 43px minmax(0, 1fr); gap: 2px; margin-bottom: 14px; min-height: 24px; }
  .outline-slide { position: relative; }
  .insert-before::before, .insert-after::after { content: ''; position: absolute; left: 0; right: 0; border-top: 2px solid var(--ok-accent); pointer-events: none; }
  .insert-before::before { top: -7px; }
  .insert-after::after { bottom: -7px; }
  button { display: flex; align-items: center; gap: 3px; align-self: start; padding: 2px 0; background: transparent; color: var(--ok-text); border: 0; font: 12px Arial, sans-serif; cursor: pointer; }
  button span { min-width: 17px; text-align: right; }
  svg { width: 20px; height: 16px; fill: var(--ok-bg); stroke: var(--ok-text-2); stroke-width: 2; }
  button.selected svg { stroke: var(--ok-accent); stroke-width: 3; }
  button:focus-visible { outline: 2px solid var(--ok-accent); }
  .text { min-width: 0; }
</style>

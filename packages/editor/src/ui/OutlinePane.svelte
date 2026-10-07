<script lang="ts">
  import { onDestroy, onMount, tick } from 'svelte';
  import { getCollapsedOutlineSlides, setSlideOutlineCollapsed, setParagraphAlignment, findShapeById, getParagraphLevel, getShapeId, getSlidePartName, type SlideData } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { outlineShapes, outlineSlots, outlineParagraphBlock, outlineParagraphRange, moveOutlineParagraphs } from '../core/outline.ts';
  import { bulletParagraphAt, boundaryBefore, outlineParagraphBoxes, OUTLINE_DRAG_THRESHOLD_PX, OUTLINE_LEVEL_STEP_PX, type OutlineParagraphBox } from '../core/outline-drag.ts';
  import { textClipboardHtml } from '../core/html-text-clipboard.ts';
  import { TEXT_CLIPBOARD_TYPE } from '../core/text-clipboard.ts';
  import { selectedSlideIndices } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import OutlineText from './OutlineText.svelte';
  import { OutlineSelectionModel, type OutlinePoint, type OutlineSelectionField } from '../core/outline-selection.ts';
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
  let dragging: { presentation: typeof doc.pres; version: number; indices: number[]; x: number } | null = null;
  let insertion = $state<{ index: number; after: boolean } | null>(null);
  // Dragging a single slide icon right by two outline levels demotes its
  // title into the previous slide's body, as Demote does.
  let demoting = $state<number | null>(null);
  const DEMOTE_DRAG_LEVELS = 2;
  function endDrag() { dragging = null; insertion = null; demoting = null; }
  function dragStart(event: DragEvent, index: number) {
    selection.clear();
    if (doc.selection.kind !== 'slide' || !selected.includes(index)) doc.selectSlide(index);
    dragging = { presentation: doc.pres, version: doc.version, indices: selectedSlideIndices(doc.selection), x: event.clientX };
    if (event.dataTransfer) {
      event.dataTransfer.effectAllowed = 'move';
      event.dataTransfer.setData('text/plain', dragging.indices.map(i => String(i + 1)).join(', '));
    }
  }
  function dragOver(event: DragEvent, index: number) {
    if (!dragging || dragging.presentation !== doc.pres || dragging.version !== doc.version) { insertion = null; demoting = null; return; }
    if (dragging.indices.includes(index)) {
      insertion = null;
      demoting = dragging.indices.length === 1 && index > 0 &&
        event.clientX - dragging.x >= DEMOTE_DRAG_LEVELS * OUTLINE_LEVEL_STEP_PX ? index : null;
      if (demoting !== null) {
        event.preventDefault();
        if (event.dataTransfer) event.dataTransfer.dropEffect = 'move';
      }
      return;
    }
    demoting = null;
    event.preventDefault();
    if (event.dataTransfer) event.dataTransfer.dropEffect = 'move';
    const bounds = (event.currentTarget as HTMLElement).getBoundingClientRect();
    insertion = { index, after: event.clientY > bounds.top + bounds.height / 2 };
  }
  async function drop(event: DragEvent, index: number) {
    if (dragging && demoting === index && dragging.presentation === doc.pres && dragging.version === doc.version) {
      event.preventDefault(); event.stopPropagation();
      endDrag();
      const part = getSlidePartName(doc.slideAt(index)!);
      await selection.fields().find(field => field.title && field.key.startsWith(part + ':'))?.changeLevel(false);
      return;
    }
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

  // Each outline field is its own contenteditable, so the browser cannot
  // select across them. Pointer gestures that leave a field are driven
  // through the shared selection model instead.
  let press: { origin: OutlinePoint; crossing: boolean } | null = null;
  let pressOnSelection: { point: OutlinePoint; dragged: boolean } | null = null;
  type ParagraphDrag = {
    field: OutlineSelectionField; slide: SlideData; id: number; first: number; last: number; level: number;
    x: number; y: number; moved: boolean; boxes: OutlineParagraphBox[];
    target: { slide: SlideData; id: number | null; index: number } | null; levelOffset: number;
  };
  let paragraphDrag: ParagraphDrag | null = null;
  let dropLine = $state<{ top: number; left: number } | null>(null);

  function fieldOf(target: Element): OutlineSelectionField | null {
    const root = target.closest('[role="textbox"]');
    return root ? selection.fields().find(field => field.root === root) ?? null : null;
  }
  function slotOf(key: string) {
    return outlineSlots(doc.pres).find(item => item.key === key) ?? null;
  }

  function pointerDown(event: PointerEvent) {
    press = null; pressOnSelection = null;
    if (event.button !== 0 || !(event.target instanceof Element)) return;
    const field = fieldOf(event.target);
    if (!field) return;
    const bullet = bulletParagraphAt(field, event.target, event.clientX);
    if (bullet !== null && !event.shiftKey) {
      event.preventDefault(); event.stopPropagation();
      startParagraphDrag(field, bullet, event);
      return;
    }
    const point = selection.pointAt(event.clientX, event.clientY);
    if (!point) return;
    const anchor = selection.anchor();
    const range = selection.current();
    if (event.shiftKey && anchor && range && (anchor.key !== point.key || range.start.key !== range.end.key)) {
      // Shift-click extends from the existing anchor into another field.
      event.preventDefault(); event.stopPropagation();
      selection.extendTo(point);
      return;
    }
    if (event.shiftKey) return;
    if (selection.contains(point)) {
      // Pressing on the selection may start a text drag; keep the range.
      event.stopPropagation();
      pressOnSelection = { point, dragged: false };
      return;
    }
    press = { origin: point, crossing: false };
  }

  function startParagraphDrag(field: OutlineSelectionField, index: number, event: PointerEvent) {
    const slot = slotOf(field.key);
    if (!slot) return;
    // Paragraph indices below address committed XML, so commit drafts first.
    const drafts = selection.fields().map(item => ({ item, edits: item.flush() })).filter(draft => draft.edits.length);
    if (drafts.length) doc.transact(t('Edit text'), () => { for (const { item, edits } of drafts) item.apply(edits); });
    const shape = findShapeById(slot.slide, slot.id)!;
    const block = outlineParagraphBlock(shape, index);
    const range = outlineParagraphRange(shape, block.first, block.last);
    selection.clear();
    field.select(range.start, range.end);
    paragraphDrag = {
      field, slide: slot.slide, id: slot.id, ...block, level: getParagraphLevel(shape, block.first),
      x: event.clientX, y: event.clientY, moved: false, boxes: [], target: null, levelOffset: 0,
    };
  }

  function updateParagraphDrag(drag: ParagraphDrag, event: PointerEvent) {
    if (!drag.moved) {
      if (Math.hypot(event.clientX - drag.x, event.clientY - drag.y) < OUTLINE_DRAG_THRESHOLD_PX) return;
      drag.moved = true;
      drag.boxes = outlineParagraphBoxes(selection.fields());
    }
    drag.levelOffset = Math.trunc((event.clientX - drag.x) / OUTLINE_LEVEL_STEP_PX);
    const after = boundaryBefore(drag.boxes, event.clientY);
    const slot = after && slotOf(after.field.key);
    if (!after || !slot) { drag.target = null; dropLine = null; return; }
    drag.target = slot.title
      ? { slide: slot.slide, id: outlineShapes(slot.slide).find(item => !item.title)?.id ?? null, index: 0 }
      : { slide: slot.slide, id: slot.id, index: after.index + 1 };
    const bounds = pane.getBoundingClientRect();
    const next = drag.boxes[drag.boxes.indexOf(after) + 1];
    const level = Math.max(0, Math.min(8, drag.level + drag.levelOffset));
    dropLine = {
      top: (next ? (after.bottom + next.top) / 2 : after.bottom) - bounds.top + pane.scrollTop,
      left: after.left - bounds.left + level * OUTLINE_LEVEL_STEP_PX,
    };
  }

  async function finishParagraphDrag(drag: ParagraphDrag) {
    dropLine = null;
    if (!drag.moved || !drag.target) return;
    const staying = drag.target.slide === drag.slide && drag.target.id === drag.id &&
      drag.target.index >= drag.first && drag.target.index <= drag.last + 1;
    if (staying && drag.levelOffset < 0 && drag.level === 0) {
      // Dragging a top-level bullet left promotes it to a slide title, as Promote does.
      await drag.field.changeLevel(true);
      return;
    }
    if (staying && !drag.levelOffset) return;
    const label = staying ? t(drag.levelOffset < 0 ? 'Promote' : 'Demote') : t('Move paragraphs');
    const target = drag.target;
    let result: ReturnType<typeof moveOutlineParagraphs>;
    try {
      result = doc.transact(label, () => moveOutlineParagraphs(
        { slide: drag.slide, id: drag.id, first: drag.first, last: drag.last }, target, drag.levelOffset));
    } catch (error) { editor.toast('error', error instanceof Error ? error.message : String(error)); return; }
    if (!result) return;
    await tick();
    const key = `${getSlidePartName(target.slide)}:${getShapeId(result.shape)}`;
    const field = selection.fields().find(item => item.key === key);
    const range = outlineParagraphRange(result.shape, result.first, result.last);
    field?.select(range.start, range.end);
  }

  // pointerMove and pointerUp listen on the window in the capture phase:
  // RichTextInput stops pointerup propagation, and a release over a field
  // must still end the gesture.
  function pointerMove(event: PointerEvent) {
    if (paragraphDrag) { updateParagraphDrag(paragraphDrag, event); return; }
    if (!press || !(event.buttons & 1)) return;
    const point = selection.pointAt(event.clientX, event.clientY);
    if (!point) return;
    if (point.key !== press.origin.key) press.crossing = true;
    if (press.crossing) selection.extendTo(point, press.origin);
  }

  function pointerUp() {
    const drag = paragraphDrag;
    paragraphDrag = null;
    if (drag) void finishParagraphDrag(drag);
    const pressed = pressOnSelection;
    pressOnSelection = null;
    press = null;
    if (pressed && !pressed.dragged) {
      // A click inside the selection without dragging places the caret there.
      const field = selection.fields().find(item => item.key === pressed.point.key);
      selection.clear();
      if (field) { field.focus(pressed.point.offset); selection.setCaret(field, pressed.point.offset); }
    }
  }

  function textDragStart(event: DragEvent) {
    const node = event.target instanceof Node ? event.target : null;
    const element = node instanceof Element ? node : node?.parentElement;
    if (!element || !fieldOf(element)) return;
    if (pressOnSelection) pressOnSelection.dragged = true;
    const copied = selection.beginTextDrag();
    if (!copied || !event.dataTransfer) return;
    event.dataTransfer.effectAllowed = 'copyMove';
    event.dataTransfer.setData('text/plain', copied.text);
    event.dataTransfer.setData('text/html', textClipboardHtml(copied));
    event.dataTransfer.setData(TEXT_CLIPBOARD_TYPE, JSON.stringify({ version: 1, ...copied }));
  }

  function cancelPointer(event: KeyboardEvent) {
    if (event.key !== 'Escape' || !paragraphDrag) return;
    event.preventDefault();
    paragraphDrag = null;
    dropLine = null;
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

<svelte:window ondragend={() => { endDrag(); selection.endTextDrag(); }} ondrop={endDrag} onblur={endDrag} onpointermovecapture={pointerMove} onpointerupcapture={pointerUp} onkeydowncapture={cancelPointer} />

<nav class="outline-pane" bind:this={pane} aria-label={t('Outline View')} onpointerdowncapture={pointerDown} ondragstartcapture={textDragStart}>
  {#if dropLine}<div class="drop-line" style={`top: ${dropLine.top}px; left: ${dropLine.left}px`} aria-hidden="true"></div>{/if}
  {#key doc.pres}
    {#each entries as entry, index (doc.slides[index])}
      <div class="outline-slide" role="group" aria-label={`${t('Slide')} ${index + 1}`} class:insert-before={insertion?.index === index && !insertion.after} class:insert-after={insertion?.index === index && insertion.after} class:demoting={demoting === index} data-outline-slide={index} ondragover={event => dragOver(event, index)} ondragleave={() => insertion = null} ondrop={event => { void drop(event, index); }}>
        <button draggable="true" ondragstart={event => dragStart(event, index)} class:selected={selected.includes(index)} aria-label={`${t('Slide')} ${index + 1}`} aria-pressed={selected.includes(index)} aria-expanded={!entry.collapsed} ondblclick={() => toggleCollapse(index)} onclick={event => { selection.clear(); doc.selectSlide(index, { additive: event.metaKey || event.ctrlKey, range: event.shiftKey }); }} onkeydown={event => keys(event, index)} oncontextmenu={event => { event.preventDefault(); if (doc.selection.kind !== 'slide' || !selected.includes(index)) { selection.clear(); doc.selectSlide(index); } editor.openContextMenu(event.clientX, event.clientY, { source: 'outline' }); }}><span>{index + 1}</span><svg viewBox="0 0 17 12" aria-hidden="true"><rect x="1" y="1" width="15" height="10" /></svg></button>
        <div class="text">
          {#each entry.shapes as shape (shape.id)}{#if shape.title || !entry.collapsed}<OutlineText slideIndex={index} shapeId={shape.id} title={shape.title} {selection} />{/if}{/each}
        </div>
      </div>
    {/each}
  {/key}
</nav>

<style>
  /* Mac PowerPoint's outline: the slide number at x = 6, a 17 × 12 pt slide
     icon, titles at x = 36 and a 27 pt pitch between slides without body text. */
  .outline-pane { position: relative; overflow: auto; min-height: 0; padding: 4px 0 12px; background: var(--ok-panel-2); border-right: 1px solid var(--ok-border); }
  .outline-slide { display: grid; grid-template-columns: 32px minmax(0, 1fr); align-items: start; min-height: 27px; }
  .outline-slide { position: relative; }
  .insert-before::before, .insert-after::after { content: ''; position: absolute; left: 0; right: 0; border-top: 2px solid var(--ok-accent); pointer-events: none; }
  .insert-before::before { top: 0; }
  /* The demotion target: body text of the previous slide, one level in. */
  .demoting::after { content: ''; position: absolute; left: calc(32px + 6px + 11.5px); right: 0; bottom: 0; border-top: 2px solid var(--ok-accent); pointer-events: none; }
  .drop-line { position: absolute; right: 0; height: 0; border-top: 2px solid var(--ok-accent); pointer-events: none; z-index: 1; }
  .insert-after::after { bottom: 0; }
  button { display: flex; align-items: center; gap: 3px; align-self: start; height: 20px; padding: 0 0 0 4px; background: transparent; color: var(--ok-text); border: 0; font: 12px Arial, sans-serif; cursor: pointer; }
  button span { min-width: 8px; text-align: right; }
  svg { width: 17px; height: 12px; fill: var(--ok-bg); stroke: var(--ok-text-2); stroke-width: 2; }
  button.selected svg { stroke: var(--ok-accent); stroke-width: 3; }
  button:focus-visible { outline: 2px solid var(--ok-accent); }
  .text { min-width: 0; }
</style>

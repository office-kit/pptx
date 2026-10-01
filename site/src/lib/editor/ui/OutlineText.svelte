<script lang="ts">
  import { onDestroy, tick, untrack } from 'svelte';
  import { getShapeText, getParagraphLevel, setParagraphLevel, getSlides, getSlideLayout, addSlideAt, setShapeText, setShapeParagraphs, findShapeById, copyShape, removeShape } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { outlineShapes } from '../core/outline.ts';
  import { textEditDiff } from '../core/text-edit-diff.ts';
  import { replayTextEdits, type TextEdit } from '../core/text-edit-preview.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let { slideIndex, shapeId, title }: { slideIndex: number; shapeId: number; title: boolean } = $props();
  const editor = getEditor();
  const doc = editor.doc;
  const presentation = untrack(() => doc.pres);
  const slide = untrack(() => doc.slideAt(slideIndex)!);
  let value = $state(untrack(() => getShapeText(doc.shapeById(slideIndex, shapeId)!)));
  let input: HTMLTextAreaElement;
  let changes: TextEdit[] = [];
  let range = { start: 0, end: 0 };
  let timer: ReturnType<typeof setTimeout> | undefined;
  let composing = false;

  function commit() {
    clearTimeout(timer);
    if (!changes.length) return;
    const edits = changes;
    changes = [];
    // History restoration and external reloads replace the model; an old draft
    // must never be applied to a replacement that happens to reuse shape IDs.
    if (doc.pres !== presentation) return;
    const index = getSlides(presentation).indexOf(slide);
    const shape = index < 0 ? null : doc.shapeById(index, shapeId);
    if (shape) doc.transact(t('Edit text'), () => replayTextEdits(shape, edits));
  }
  // Selection events can arrive after Svelte has detached this input on Undo or view changes.
  function rememberRange(target: HTMLTextAreaElement = input) { range = { start: target.selectionStart, end: target.selectionEnd }; }
  function changed() {
    const next = input.value;
    const change = textEditDiff(value, next, range, input.selectionStart);
    if (change) changes.push(change);
    value = next;
    rememberRange();
    clearTimeout(timer);
    if (!composing) timer = setTimeout(commit, 600);
  }
  async function keys(event: KeyboardEvent) {
    if (event.isComposing) return;
    const mod = event.metaKey || event.ctrlKey;
    if (mod && ['z', 'y'].includes(event.key.toLowerCase())) {
      event.preventDefault(); event.stopPropagation(); commit();
      await (event.shiftKey || event.key.toLowerCase() === 'y' ? doc.redo() : doc.undo());
    } else if (mod && event.key.toLowerCase() === 's') commit();
    else if (event.key === 'Escape') { commit(); input.blur(); }
    else if (event.key === 'Tab' && !event.shiftKey && !mod && !event.altKey && !title) {
      event.preventDefault(); event.stopPropagation();
      rememberRange(); commit();
      const source = doc.shapeById(slideIndex, shapeId)!;
      const maxLevel = 8;
      if (getParagraphLevel(source, range).some(level => level < maxLevel)) {
        doc.transact(t('Indent'), () => setParagraphLevel(source, range, { offset: 1 }));
      }
    }
    else if (event.key === 'Enter' && !event.shiftKey && !mod && !event.altKey && title) {
      const layout = getSlideLayout(slide);
      if (!layout) return;
      const ownerDocument = input.ownerDocument;
      const start = input.selectionStart;
      const end = input.selectionEnd;
      event.preventDefault(); event.stopPropagation(); commit();
      const source = doc.shapeById(slideIndex, shapeId)!;
      const index = getSlides(doc.pres).indexOf(slide) + 1;
      doc.transact(t('New slide'), () => {
        const next = addSlideAt(doc.pres, index, { layout });
        const placeholders = outlineShapes(next);
        for (const item of placeholders) setShapeText(findShapeById(next, item.id)!, '');
        // Mac PowerPoint moves the following outline body when Enter splits a title.
        // Copy whole placeholders so paragraph levels, bullets and links survive.
        for (const item of placeholders) {
          if (!item.title) removeShape(findShapeById(next, item.id)!);
        }
        for (const item of outlineShapes(slide)) {
          if (item.title) continue;
          const body = findShapeById(slide, item.id)!;
          copyShape(next, body);
          setShapeText(body, '');
        }
        const heading = placeholders.find(item => item.title);
        if (heading) setShapeParagraphs(findShapeById(next, heading.id)!, { source, range: { start: end, end: value.length } });
        setShapeText(source, '', { range: { start, end: value.length } });
        doc.selectSlide(index);
      });
      await tick();
      ownerDocument.querySelector<HTMLTextAreaElement>(`[data-outline-slide="${index}"] textarea`)?.focus();
    }
  }
  $effect(() => {
    doc.version;
    if (!changes.length && doc.pres === presentation) {
      const shape = doc.shapeById(slideIndex, shapeId);
      if (shape) value = getShapeText(shape);
    }
  });
  $effect(() => { value; if (input) { input.style.height = '0'; input.style.height = `${input.scrollHeight}px`; } });
  onDestroy(() => untrack(commit));
</script>

<textarea bind:this={input} {value} class:title aria-label={`${t(title ? 'Outline title' : 'Outline text')} ${slideIndex + 1}`} rows="1" spellcheck="false" onfocus={() => doc.selectShape(slideIndex, shapeId)} onbeforeinput={event => rememberRange(event.currentTarget)} onselect={event => rememberRange(event.currentTarget)} oninput={changed} onblur={commit} onkeydown={keys} oncompositionstart={() => { composing = true; clearTimeout(timer); }} oncompositionend={() => { composing = false; changed(); }}></textarea>

<style>
  textarea { display: block; width: 100%; box-sizing: border-box; resize: none; overflow: hidden; min-height: 23px; padding: 2px 4px; border: 0; outline: none; color: var(--ok-text); background: transparent; font: 14px/1.4 Arial, sans-serif; }
  textarea.title { font-weight: bold; }
  textarea:not(.title) { padding-left: 20px; }
  textarea:focus { background: var(--ok-hover); }
</style>

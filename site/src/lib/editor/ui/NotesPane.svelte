<script lang="ts">
  import { onDestroy, onMount, untrack } from 'svelte';
  import { getSlideNotes, getSlideNotesParagraphEndFormat, getSlideNotesTextFormats, getSlides, setSlideNotes, setSlideNotesFormat, toWritableTextFormat, transformSlideNotesCase, type TextCase, type TextFormat } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { textEditDiff } from '../core/text-edit-diff.ts';
  import { stepFontSize } from '../core/font-size.ts';
  import type { TextEdit } from '../core/text-edit-preview.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import RichTextInput from './RichTextInput.svelte';
  import { parseHtmlTextClipboard, textClipboardHtml } from '../core/html-text-clipboard.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const slide = untrack(() => doc.currentSlide!);
  const presentation = untrack(() => doc.pres);
  let value = $state(getSlideNotes(slide) ?? '');
  let pending = false;
  let changes: TextEdit[] = [];
  let range = { start: 0, end: 0 };
  let composing = false;
  let timer: ReturnType<typeof setTimeout> | undefined;
  let typingFormat: { format: TextFormat; reset: boolean } | undefined;
  let input: { focus(): void; blur(): void; getElement(): HTMLElement | undefined; setSelectionRange(start: number, end: number): void; getSelection(): { start: number; end: number } };
  let pane: HTMLElement;
  let drag: { id: number; y: number; height: number } | null = null;
  let maxHeight = $state(400);
  let noteFormat = $state<TextFormat>({});
  let focused = $state(false);
  let formattedRanges = $state<{ start: number; end: number; format: TextFormat }[]>(
    getSlideNotesTextFormats(slide).map((item) => ({ ...item, format: toWritableTextFormat(item.format) })),
  );
  const noteHtml = $derived.by(() => {
    const text = value;
    if (!text) return '';
    const points = new Set([0, text.length]);
    for (const item of formattedRanges) { points.add(item.start); points.add(item.end); }
    const sorted = [...points].sort((a, b) => a - b);
    const formats: { start: number; end: number; format: TextFormat }[] = [];
    let rangeIndex = 0;
    for (let index = 0; index < sorted.length - 1; index++) {
      const start = sorted[index]!;
      const end = sorted[index + 1]!;
      while (rangeIndex < formattedRanges.length && formattedRanges[rangeIndex]!.end <= start) rangeIndex++;
      const current = formattedRanges[rangeIndex];
      formats.push({ start, end, format: current && current.start <= start && current.end >= end ? current.format : {} });
    }
    return textClipboardHtml({ text, formats }, { editing: true });
  });

  function updateHeightLimit() {
    maxHeight = Math.max(60, Math.min((pane.parentElement?.clientHeight ?? 600) - 100, pane.ownerDocument.defaultView!.innerHeight / 2));
    editor.notesHeight = Math.min(maxHeight, Math.max(60, editor.notesHeight));
  }
  onMount(() => {
    const observer = new ResizeObserver(updateHeightLimit);
    observer.observe(pane.parentElement!);
    const view = pane.ownerDocument.defaultView!;
    view.addEventListener('resize', updateHeightLimit);
    updateHeightLimit();
    const leaveNotes = (event: Event) => {
      const target = event.target;
      if (!(target instanceof Element) || !target.closest('.canvas-viewport')) return;
      focused = false;
      input?.blur();
      editor.notesInlineTextFormat = null;
    };
    pane.ownerDocument.addEventListener('pointerdown', leaveNotes, true);
    pane.ownerDocument.addEventListener('focusin', leaveNotes, true);
    return () => {
      observer.disconnect();
      view.removeEventListener('resize', updateHeightLimit);
      pane.ownerDocument.removeEventListener('pointerdown', leaveNotes, true);
      pane.ownerDocument.removeEventListener('focusin', leaveNotes, true);
    };
  });

  function commit() {
    clearTimeout(timer);
    if (!pending) return;
    pending = false;
    const edits = changes;
    changes = [];
    // An external replacement or deleted slide must never receive a stale draft.
    if (doc.pres !== presentation || !getSlides(presentation).includes(slide)) return;
    if (value !== (getSlideNotes(slide) ?? '')) {
      try { doc.transact(t('Speaker notes'), () => {
        for (const edit of edits) {
          setSlideNotes(slide, edit.text, { range: { start: edit.start, end: edit.end } });
          if (edit.typing && edit.text.length) {
            setSlideNotesFormat(slide, edit.typing.format, {
              range: { start: edit.start, end: edit.start + edit.text.length },
              reset: edit.typing.reset,
            });
          }
        }
      }); }
      catch (error) { editor.toast('error', String(error)); }
    }
  }
  function formatForRange(next: { start: number; end: number }) {
    const matching: TextFormat[] = [];
    if (next.start === next.end) {
      const item = formattedRanges.find((candidate) => candidate.start <= next.start && (candidate.end > next.start || (candidate.end === next.start && next.start === value.length)));
      matching.push(item?.format ?? {});
    } else {
      let cursor = next.start;
      for (const item of formattedRanges) {
        if (item.end <= next.start) continue;
        if (item.start >= next.end) break;
        if (item.start > cursor) matching.push({});
        const start = Math.max(cursor, item.start);
        const end = Math.min(next.end, item.end);
        if (start < end) {
          matching.push(item.format);
          cursor = end;
        }
      }
      if (cursor < next.end) matching.push({});
    }
    if (matching.length === 0) { noteFormat = {}; return; }
    const keys = new Set(Object.keys(matching[0]!));
    for (const item of matching.slice(1)) for (const key of [...keys]) {
      if (item[key as keyof TextFormat] !== matching[0]![key as keyof TextFormat]) keys.delete(key);
    }
    noteFormat = Object.fromEntries([...keys].map((key) => [key, matching[0]![key as keyof TextFormat]])) as TextFormat;
  }
  function changed(next: string) {
    const after = input?.getSelection?.() ?? range;
    const change = textEditDiff(value, next, range, after.end);
    if (change) {
      if (typingFormat && change.text.length) {
        change.typing = typingFormat;
      }
      changes.push(change);
    }
    value = next;
    range = after;
    pending = changes.length > 0;
    clearTimeout(timer);
    if (!composing) timer = setTimeout(commit, 600);
  }
  function keys(event: KeyboardEvent) {
    if (event.isComposing) return;
    const mod = event.metaKey || event.ctrlKey;
    if (mod && ['z', 'y'].includes(event.key.toLowerCase())) {
      event.preventDefault(); event.stopPropagation();
      commit();
      typingFormat = undefined;
      const backward = !(event.shiftKey || event.key.toLowerCase() === 'y');
      void (backward ? doc.undo() : doc.redo());
    } else if (mod && event.key.toLowerCase() === 's') commit();
    else if (event.key === 'Escape') { commit(); input.blur(); }
  }
  function readFormattedRanges() {
    return getSlideNotesTextFormats(slide).map((item) => ({ ...item, format: toWritableTextFormat(item.format) }));
  }
  function currentParagraphEnd() {
    return value.slice(0, range.start).split('\n').length - 1;
  }
  function currentParagraphIsEmpty() {
    const start = value.lastIndexOf('\n', Math.max(0, range.start - 1)) + 1;
    const newline = value.indexOf('\n', range.start);
    const end = newline === -1 ? value.length : newline;
    return value.slice(start, end).length === 0;
  }
  function applyNoteFormat(format: TextFormat, reset = false) {
    commit();
    if (range.start === range.end) {
      // PowerPoint persists paragraph-end formatting for an empty paragraph,
      // which makes the toggle itself part of the shared document history.
      // Once text exists, a middle/end caret format is only a typing state;
      // creating an OOXML endParaRPr there would add a phantom undo step and
      // incorrectly affect a later caret in the same paragraph.
      if (currentParagraphIsEmpty()) {
        const paragraphEnd = currentParagraphEnd();
        doc.transact(t(reset ? 'Clear notes typing format' : 'Format notes typing'), () => {
          setSlideNotesFormat(slide, format, { paragraphEnd, reset });
        });
      }
      const after = reset ? { ...format } : { ...(typingFormat?.format ?? noteFormat), ...format };
      typingFormat = {
        format: after,
        reset: typingFormat?.reset === true || reset,
      };
      noteFormat = reset ? {} : { ...noteFormat, ...format };
      return;
    }
    doc.transact(t(reset ? 'Clear text formatting' : 'Format selected notes'), () => {
      setSlideNotesFormat(slide, format, { range, reset });
    });
    formattedRanges = readFormattedRanges();
    formatForRange(range);
    input.focus();
    input.setSelectionRange(range.start, range.end);
  }
  function toggleNoteFormat(property: 'bold' | 'italic' | 'underline' | 'strike' | 'superscript' | 'subscript') {
    const format = property === 'underline' ? { underline: noteFormat.underline === 'sng' ? 'none' : 'sng' } : property === 'superscript' ? { baseline: noteFormat.baseline === 30 ? 0 : 30 } : property === 'subscript' ? { baseline: noteFormat.baseline === -25 ? 0 : -25 } : property === 'bold' ? { bold: !noteFormat.bold } : property === 'italic' ? { italic: !noteFormat.italic } : { strike: !noteFormat.strike };
    applyNoteFormat(format);
  }
  function changeCase(caseValue: TextCase) {
    commit();
    if (range.start === range.end) return;
    doc.transact(t('Change Case'), () => {
      transformSlideNotesCase(slide, caseValue, { range });
    });
    value = getSlideNotes(slide) ?? value;
    formattedRanges = readFormattedRanges();
    formatForRange(range);
    input.focus();
    input.setSelectionRange(range.start, range.end);
  }
  function fontSize(direction: 1 | -1) {
    applyNoteFormat({ size: stepFontSize(noteFormat.size ?? 12, direction) });
  }
  function pasteNotes(event: ClipboardEvent) {
    const data = event.clipboardData;
    if (!data) return;
    const plain = data.getData('text/plain');
    const parsed = parseHtmlTextClipboard(data.getData('text/html'), plain);
    if (!parsed || range.start === range.end && !plain) return;
    event.preventDefault();
    const start = range.start;
    const next = `${value.slice(0, start)}${parsed.text}${value.slice(range.end)}`;
    commit();
    if (doc.pres !== presentation || !getSlides(presentation).includes(slide)) return;
    doc.transact(t('Paste formatted notes'), () => {
      setSlideNotes(slide, parsed.text, { range: { start, end: range.end } });
      for (const item of parsed.formats) {
        const end = Math.min(parsed.text.length, item.end);
        if (item.start < end) setSlideNotesFormat(slide, item.format, { range: { start: start + item.start, end: start + end } });
      }
    });
    value = next;
    range = { start: start + parsed.text.length, end: start + parsed.text.length };
    formattedRanges = readFormattedRanges();
    formatForRange(range);
    input.focus();
    input.setSelectionRange(range.start, range.end);
  }
  $effect(() => {
    const api = {
      formats: [noteFormat],
      apply: applyNoteFormat, changeCase, fontSize,
      toggle: toggleNoteFormat,
    };
    if (focused) editor.notesInlineTextFormat = api;
    return () => { if (editor.notesInlineTextFormat === api) editor.notesInlineTextFormat = null; };
  });
  $effect(() => {
    doc.version;
    if (!pending) {
      value = getSlideNotes(slide) ?? '';
      formattedRanges = getSlideNotesTextFormats(slide).map((item) => ({ ...item, format: toWritableTextFormat(item.format) }));
      if (range.start === range.end && currentParagraphIsEmpty()) {
        const paragraphEnd = currentParagraphEnd();
        const paragraphFormat = toWritableTextFormat(getSlideNotesParagraphEndFormat(slide, paragraphEnd));
        const currentFormat = untrack(() => noteFormat);
        if (JSON.stringify(currentFormat) !== JSON.stringify(paragraphFormat)) noteFormat = paragraphFormat;
        typingFormat = { format: { ...paragraphFormat }, reset: false };
      }
    }
  });
  $effect(() => { if (editor.notesFocusRequest && input) { input.focus(); editor.notesFocusRequest = 0; } });
  onDestroy(() => untrack(commit));

  function resizeStart(event: PointerEvent) {
    if (event.button !== 0) return;
    event.preventDefault();
    updateHeightLimit();
    drag = { id: event.pointerId, y: event.clientY, height: pane.clientHeight };
    (event.currentTarget as HTMLElement).setPointerCapture(event.pointerId);
  }
  function resizeMove(event: PointerEvent) {
    if (drag?.id !== event.pointerId) return;
    editor.notesHeight = Math.min(maxHeight, Math.max(60, drag.height + drag.y - event.clientY));
  }
  function resizeKeys(event: KeyboardEvent) {
    if (!['ArrowUp', 'ArrowDown', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault(); event.stopPropagation();
    updateHeightLimit();
    editor.notesHeight = event.key === 'Home' ? 60 : event.key === 'End' ? maxHeight : Math.min(maxHeight, Math.max(60, editor.notesHeight + (event.key === 'ArrowUp' ? 10 : -10)));
  }
</script>

<section class="notes-pane" bind:this={pane} aria-label={t('Notes')} style:height="{editor.notesHeight}px">
  <!-- A focusable separator implements the ARIA window-splitter pattern. -->
  <!-- svelte-ignore a11y_no_noninteractive_tabindex, a11y_no_noninteractive_element_interactions -->
  <div class="resize" role="separator" tabindex="0" aria-label={t('Notes pane height')} aria-orientation="horizontal" aria-valuemin={60} aria-valuemax={maxHeight} aria-valuenow={editor.notesHeight} onpointerdown={resizeStart} onpointermove={resizeMove} onpointerup={() => drag = null} onpointercancel={() => drag = null} onlostpointercapture={() => drag = null} onkeydown={resizeKeys}></div>
  <RichTextInput bind:this={input} {value} html={noteHtml} label={t('Notes content')} style="position:static; width:100%; height:100%; min-height:40px; box-sizing:border-box;" textZoom={1}
    onfocus={() => { editor.inlineTextFormat = null; focused = true; }} onblur={() => { commit(); setTimeout(() => { if (document.activeElement !== input?.getElement?.()) focused = false; }, 0); }}
    onselect={(next) => { if (next.start !== range.start || next.end !== range.end) typingFormat = undefined; range = next; formatForRange(next); }} onbeforeinput={(next) => { range = next; formatForRange(next); }} oninput={changed} onkeydown={keys}
    onnewline={() => changed(`${value.slice(0, range.start)}\n${value.slice(range.end)}`)} oncomposition={(active) => { composing = active; if (active) clearTimeout(timer); }}
    onhistory={(backward) => { commit(); typingFormat = undefined; void (backward ? doc.undo() : doc.redo()); }} oncopy={() => {}} oncut={() => {}} onpaste={pasteNotes} />
</section>

<style>
  .notes-pane { position: relative; min-height: 60px; max-height: 50vh; box-sizing: border-box; border-top: 1px solid var(--ok-border); background: var(--ok-panel); padding: 10px 16px 8px; }
  .resize { position: absolute; left: 0; right: 0; top: -3px; height: 6px; cursor: ns-resize; touch-action: none; }
  .resize:focus-visible { outline: 2px solid var(--ok-accent); }
</style>

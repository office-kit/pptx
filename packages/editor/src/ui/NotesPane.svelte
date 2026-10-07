<script lang="ts">
  import { mergeTextFormat } from '../core/merge-text-format.ts';
  import { onDestroy, onMount, untrack } from 'svelte';
  import { asColor, getSlideNotes, getSlideNotesLineBreaks, getSlideNotesParagraphEndFormat, getSlideNotesTextFormats, getSlides, resolveSlideNotesTextColor, setSlideNotes, setSlideNotesFormat, toWritableTextFormat, transformSlideNotesCase, type TextCase, type TextFormat } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { activeElementOf } from '../core/dom-root.ts';
  import { textEditDiff } from '../core/text-edit-diff.ts';
  import { stepFontSize } from '../core/font-size.ts';
  import { toggleTextFormat, type TextFormatToggle } from '../core/text-format-toggle.ts';
  import type { TextEdit } from '../core/text-edit-preview.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import RichTextInput from './RichTextInput.svelte';
  import { parseHtmlTextClipboard, textClipboardHtml } from '../core/html-text-clipboard.ts';

  // Mac PowerPoint opens the notes pane one line tall: 39 pt plus its 5 pt splitter.
  const NOTES_MIN_HEIGHT = 44;
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
  let pendingParagraphs = new Set<number>();
  // `getSlideNotes` intentionally exposes both paragraph breaks and soft
  // breaks as `\n`. Keep the distinction while an edit is live so paragraph
  // end formatting is indexed by OOXML `<a:p>`, rather than every newline.
  let softBreaks = new Set(getSlideNotesLineBreaks(slide).filter((item) => item.kind === 'break').map((item) => item.position));
  let input: { focus(): void; blur(): void; getElement(): HTMLElement | undefined; setSelectionRange(start: number, end: number): void; getSelection(): { start: number; end: number } };
  let pane: HTMLElement;
  let drag: { id: number; y: number; height: number } | null = null;
  let maxHeight = $state(400);
  let noteFormat = $state<TextFormat>({});
  let displayNoteFormat = $state<TextFormat>({});
  let focused = $state(false);
  let formattedRanges = $state<{ start: number; end: number; format: TextFormat }[]>(
    getSlideNotesTextFormats(slide).map((item) => ({ ...item, format: toWritableTextFormat(item.format) })),
  );
  let resolvedFormattedRanges = $state<{ start: number; end: number; format: TextFormat }[]>(
    getSlideNotesTextFormats(slide, { resolveColors: true }).map((item) => ({ ...item, format: toWritableTextFormat(item.format) })),
  );
  const noteHtml = $derived.by(() => {
    const text = value;
    if (!text) return '';
    const displayRanges = pendingResolvedRanges();
    const points = new Set([0, text.length]);
    for (const item of displayRanges) { points.add(item.start); points.add(item.end); }
    const sorted = [...points].sort((a, b) => a - b);
    const formats: { start: number; end: number; format: TextFormat }[] = [];
    let rangeIndex = 0;
    for (let index = 0; index < sorted.length - 1; index++) {
      const start = sorted[index]!;
      const end = sorted[index + 1]!;
      while (rangeIndex < displayRanges.length && displayRanges[rangeIndex]!.end <= start) rangeIndex++;
      const current = displayRanges[rangeIndex];
      formats.push({ start, end, format: current && current.start <= start && current.end >= end ? current.format : {} });
    }
    return textClipboardHtml({ text, formats }, { editing: true });
  });

  function updateHeightLimit() {
    maxHeight = Math.max(NOTES_MIN_HEIGHT, Math.min((pane.parentElement?.clientHeight ?? 600) - 100, pane.ownerDocument.defaultView!.innerHeight / 2));
    editor.notesHeight = Math.min(maxHeight, Math.max(NOTES_MIN_HEIGHT, editor.notesHeight));
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
    pendingParagraphs.clear();
    // An external replacement or deleted slide must never receive a stale draft.
    if (doc.pres !== presentation || !getSlides(presentation).includes(slide)) return;
    try { doc.transact(t('Speaker notes'), () => {
      for (const edit of edits) {
        setSlideNotes(slide, edit.text, {
          range: { start: edit.start, end: edit.end },
          newlines: edit.newlines,
        });
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
  function formatInRanges(next: { start: number; end: number }, ranges: typeof formattedRanges): TextFormat {
    const matching: TextFormat[] = [];
    if (next.start === next.end) {
      const item = ranges.find((candidate) => candidate.start <= next.start && (candidate.end > next.start || (candidate.end === next.start && next.start === value.length)));
      matching.push(item?.format ?? {});
    } else {
      let cursor = next.start;
      for (const item of ranges) {
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
    if (matching.length === 0) return {};
    const keys = new Set(Object.keys(matching[0]!));
    for (const item of matching.slice(1)) for (const key of [...keys]) {
      if (item[key as keyof TextFormat] !== matching[0]![key as keyof TextFormat]) keys.delete(key);
    }
    return Object.fromEntries([...keys].map((key) => [key, matching[0]![key as keyof TextFormat]])) as TextFormat;
  }
  function formatForRange(next: { start: number; end: number }) {
    const nextFormat = formatInRanges(next, formattedRanges);
    if (JSON.stringify(noteFormat) !== JSON.stringify(nextFormat)) noteFormat = nextFormat;
    const nextDisplayFormat = formatInRanges(next, resolvedFormattedRanges);
    if (JSON.stringify(displayNoteFormat) !== JSON.stringify(nextDisplayFormat)) displayNoteFormat = nextDisplayFormat;
  }
  function reindexPendingParagraphs(change: TextEdit, insertedNewlinesAreParagraphs = true) {
    if (pendingParagraphs.size === 0) return;
    // TextEdit offsets are UTF-16 offsets (the same convention used by the
    // contenteditable selection). Iterate by code unit here: spreading a
    // string would count astral characters as one item and shift all later
    // soft-break/paragraph positions by one.
    let removedParagraphs = 0;
    for (let index = change.start; index < change.end; index++) {
      if (value[index] === '\n' && !softBreaks.has(index)) removedParagraphs++;
    }
    let insertedParagraphs = 0;
    if (insertedNewlinesAreParagraphs) {
      for (let index = 0; index < change.text.length; index++) {
        if (change.text[index] === '\n') insertedParagraphs++;
      }
    }
    if (removedParagraphs === 0 && insertedParagraphs === 0) return;
    let paragraph = 0;
    for (let index = 0; index < change.start; index++) {
      if (value[index] === '\n' && !softBreaks.has(index)) paragraph++;
    }
    const next = new Set<number>();
    for (const index of pendingParagraphs) {
      if (index <= paragraph) next.add(index);
      else if (index > paragraph + removedParagraphs) next.add(index - removedParagraphs + insertedParagraphs);
    }
    pendingParagraphs = next;
  }
  function changed(next: string, nextSelection?: { start: number; end: number }) {
    // RichTextInput cancels the browser paragraph insertion before this
    // callback, so the DOM selection still points at the pre-newline range.
    // Callers that synthesize an edit must provide the resulting caret rather
    // than making the next edit depend on a stale DOM selection.
    const after = nextSelection ?? input?.getSelection?.() ?? range;
    const change = textEditDiff(value, next, range, after.end);
    if (change) {
      reindexPendingParagraphs(change);
      softBreaks = new Set(
        [...softBreaks]
          .filter((position) => position < change.start || position >= change.end)
          .map((position) => position >= change.end ? position + change.text.length - (change.end - change.start) : position),
      );
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
      softBreaks.clear();
      const backward = !(event.shiftKey || event.key.toLowerCase() === 'y');
      void (backward ? doc.undo() : doc.redo());
    } else if (mod && event.key.toLowerCase() === 's') commit();
    else if (event.key === 'Escape') { commit(); input.blur(); }
  }
  function readFormattedRanges() {
    return getSlideNotesTextFormats(slide).map((item) => ({ ...item, format: toWritableTextFormat(item.format) }));
  }
  function readResolvedFormattedRanges() {
    return getSlideNotesTextFormats(slide, { resolveColors: true }).map((item) => ({ ...item, format: toWritableTextFormat(item.format) }));
  }
  function displayFormat(format: TextFormat): TextFormat {
    if (!format.color || /^#[\da-f]{6}$/i.test(format.color)) return { ...format };
    const color = resolveSlideNotesTextColor(slide, format.color);
    const parsed = color ? asColor(color) : null;
    return parsed ? { ...format, color: parsed } : { ...format };
  }
  function pendingResolvedRanges() {
    let ranges = resolvedFormattedRanges.map(item => ({ ...item, format: { ...item.format } }));
    for (const change of changes) {
      const delta = change.text.length - (change.end - change.start);
      const insertionCarrier = change.start === change.end
        ? ranges.find(item => item.start < change.start && item.end >= change.start) ??
          ranges.find(item => item.end === change.start) ??
          ranges.find(item => item.start === change.start)
        : undefined;
      const replacementCarrier = change.start !== change.end
        ? ranges.find(item => item.start <= change.start && item.end > change.start)
        : undefined;
      const inherited = insertionCarrier?.format ?? replacementCarrier?.format ?? {};
      const typed = change.typing && change.text.length
        ? mergeTextFormat(change.typing.reset ? undefined : inherited, change.typing.format)
        : null;
      if (typed?.color && !/^#[\da-f]{6}$/i.test(typed.color)) {
        const resolved = resolveSlideNotesTextColor(slide, typed.color);
        const parsed = resolved ? asColor(resolved) : null;
        if (parsed) typed.color = parsed;
        else if (displayNoteFormat.color && /^#[\da-f]{6}$/i.test(displayNoteFormat.color)) typed.color = displayNoteFormat.color;
        else delete typed.color;
      }
      const transformed: { start: number; end: number; format: TextFormat }[] = [];
      let insertedWithInheritedFormat = false;
      for (const item of ranges) {
        if (change.start === change.end) {
          const carriesInsertion = item === insertionCarrier;
          if (carriesInsertion) {
            if (typed) {
              if (item.start < change.start) transformed.push({ ...item, end: change.start });
              if (item.end > change.start) transformed.push({ ...item, start: change.start + change.text.length, end: item.end + delta });
            } else transformed.push({ ...item, end: item.end + delta });
          } else if (item.start >= change.start) {
            transformed.push({ ...item, start: item.start + delta, end: item.end + delta });
          } else transformed.push(item);
        } else {
          if (item.start < change.start) transformed.push({ ...item, end: Math.min(item.end, change.start) });
          if (item.end > change.end) {
            const start = item.start >= change.end ? item.start + delta : change.start + change.text.length;
            transformed.push({ ...item, start, end: item.end + delta });
          }
          if (!typed && item.start <= change.start && item.end >= change.end && change.text.length) {
            transformed.push({ start: change.start, end: change.start + change.text.length, format: item.format });
            insertedWithInheritedFormat = true;
          }
        }
      }
      if (typed && change.text.length) {
        transformed.push({ start: change.start, end: change.start + change.text.length, format: typed });
      } else if (!typed && change.text.length && !insertedWithInheritedFormat && replacementCarrier) {
        transformed.push({ start: change.start, end: change.start + change.text.length, format: replacementCarrier.format });
      }
      ranges = transformed;
    }
    return ranges.sort((left, right) => left.start - right.start || left.end - right.end);
  }
  function currentParagraphEnd() {
    return paragraphIndexAt(range.start);
  }
  function paragraphIndexAt(position: number) {
    let paragraph = 0;
    for (let index = 0; index < position; index++) {
      if (value[index] === '\n' && !softBreaks.has(index)) paragraph++;
    }
    return paragraph;
  }
  function paragraphBoundsAt(position: number) {
    let start = position;
    while (start > 0) {
      const newline = value.lastIndexOf('\n', start - 1);
      if (newline < 0) { start = 0; break; }
      if (softBreaks.has(newline)) { start = newline; continue; }
      start = newline + 1;
      break;
    }
    let end = position;
    while (end < value.length) {
      const newline = value.indexOf('\n', end);
      if (newline < 0 || !softBreaks.has(newline)) { end = newline < 0 ? value.length : newline; break; }
      end = newline + 1;
    }
    return { start, end };
  }
  function currentParagraphIsEmpty() {
    const { start, end } = paragraphBoundsAt(range.start);
    return value.slice(start, end).replace(/\n/g, '').length === 0;
  }
  function paragraphIsEmptyAt(position: number) {
    const { start, end } = paragraphBoundsAt(position);
    return value.slice(start, end).replace(/\n/g, '').length === 0;
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
        pendingParagraphs.clear();
        doc.transact(t(reset ? 'Clear notes typing format' : 'Format notes typing'), () => {
          setSlideNotesFormat(slide, format, { paragraphEnd, reset });
        });
      }
      const after = mergeTextFormat(reset ? undefined : (typingFormat?.format ?? noteFormat), format);
      typingFormat = {
        format: after,
        reset: typingFormat?.reset === true || reset,
      };
      noteFormat = reset ? {} : mergeTextFormat(noteFormat, format);
      // Keep the display-only theme-resolved state in lockstep with the
      // literal typing state. The next commit will re-read the ranges, but
      // the toolbar must update immediately at a collapsed caret as well.
      if (reset) displayNoteFormat = {};
      else {
        displayNoteFormat = mergeTextFormat(displayNoteFormat, displayFormat(format));
      }
      return;
    }
    doc.transact(t(reset ? 'Clear text formatting' : 'Format selected notes'), () => {
      setSlideNotesFormat(slide, format, { range, reset });
    });
    formattedRanges = readFormattedRanges();
    resolvedFormattedRanges = readResolvedFormattedRanges();
    formatForRange(range);
    input.focus();
    input.setSelectionRange(range.start, range.end);
  }
  function toggleNoteFormat(property: TextFormatToggle) {
    applyNoteFormat(toggleTextFormat([noteFormat], property));
  }
  function changeCase(caseValue: TextCase) {
    commit();
    if (range.start === range.end) return;
    doc.transact(t('Change Case'), () => {
      transformSlideNotesCase(slide, caseValue, { range });
    });
    value = getSlideNotes(slide) ?? value;
    formattedRanges = readFormattedRanges();
    resolvedFormattedRanges = readResolvedFormattedRanges();
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
    resolvedFormattedRanges = readResolvedFormattedRanges();
    formatForRange(range);
    input.focus();
    input.setSelectionRange(range.start, range.end);
  }
  $effect(() => {
    const api = {
      formats: [noteFormat],
      displayFormats: [displayNoteFormat],
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
      softBreaks = new Set(getSlideNotesLineBreaks(slide).filter((item) => item.kind === 'break').map((item) => item.position));
      formattedRanges = getSlideNotesTextFormats(slide).map((item) => ({ ...item, format: toWritableTextFormat(item.format) }));
      resolvedFormattedRanges = getSlideNotesTextFormats(slide, { resolveColors: true }).map((item) => ({ ...item, format: toWritableTextFormat(item.format) }));
      if (range.start === range.end && currentParagraphIsEmpty()) {
        const paragraphEnd = currentParagraphEnd();
        const paragraphFormat = toWritableTextFormat(getSlideNotesParagraphEndFormat(slide, paragraphEnd));
        const resolvedParagraphFormat = toWritableTextFormat(getSlideNotesParagraphEndFormat(slide, paragraphEnd, { resolveColors: true }));
        const currentFormat = untrack(() => noteFormat);
        if (JSON.stringify(currentFormat) !== JSON.stringify(paragraphFormat)) noteFormat = paragraphFormat;
        displayNoteFormat = resolvedParagraphFormat;
        typingFormat = { format: { ...paragraphFormat }, reset: false };
      } else {
        // Undo/redo and theme edits change the ranges without moving the
        // selection. Refresh both literal and resolved toolbar state even
        // when the caret is in non-empty text.
        untrack(() => formatForRange(range));
      }
    }
  });
  // The pane no longer focuses itself on mount (it is shown by default and
  // remounts on Undo); the first explicit request places the caret at the end,
  // as mounting used to.
  let caretPlaced = false;
  $effect(() => {
    if (!editor.notesFocusRequest || !input) return;
    input.focus();
    if (!caretPlaced) input.setSelectionRange(value.length, value.length);
    caretPlaced = true;
    editor.notesFocusRequest = 0;
  });
  onDestroy(() => {
    untrack(commit);
    // Undo and slide changes remount the pane; keep the caret in the notes
    // only when they had it, so canvas text editing is never interrupted.
    if (focused) editor.notesFocusRequest++;
  });

  function resizeStart(event: PointerEvent) {
    if (event.button !== 0) return;
    event.preventDefault();
    updateHeightLimit();
    drag = { id: event.pointerId, y: event.clientY, height: pane.clientHeight };
    (event.currentTarget as HTMLElement).setPointerCapture(event.pointerId);
  }
  function resizeMove(event: PointerEvent) {
    if (drag?.id !== event.pointerId) return;
    editor.notesHeight = Math.min(maxHeight, Math.max(NOTES_MIN_HEIGHT, drag.height + drag.y - event.clientY));
  }
  function resizeKeys(event: KeyboardEvent) {
    if (!['ArrowUp', 'ArrowDown', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault(); event.stopPropagation();
    updateHeightLimit();
    editor.notesHeight = event.key === 'Home' ? NOTES_MIN_HEIGHT : event.key === 'End' ? maxHeight : Math.min(maxHeight, Math.max(NOTES_MIN_HEIGHT, editor.notesHeight + (event.key === 'ArrowUp' ? 10 : -10)));
  }
</script>

<section class="notes-pane" bind:this={pane} aria-label={t('Notes')} style:height="{editor.notesHeight}px">
  <!-- A focusable separator implements the ARIA window-splitter pattern. -->
  <!-- svelte-ignore a11y_no_noninteractive_tabindex, a11y_no_noninteractive_element_interactions -->
  <div class="resize" role="separator" tabindex="0" aria-label={t('Notes pane height')} aria-orientation="horizontal" aria-valuemin={NOTES_MIN_HEIGHT} aria-valuemax={maxHeight} aria-valuenow={editor.notesHeight} onpointerdown={resizeStart} onpointermove={resizeMove} onpointerup={() => drag = null} onpointercancel={() => drag = null} onlostpointercapture={() => drag = null} onkeydown={resizeKeys}></div>
  <RichTextInput bind:this={input} {value} html={noteHtml} label={t('Notes content')} autofocus={false} style="position:static; width:100%; height:100%; min-height:20px; box-sizing:border-box;" textZoom={1}
    onfocus={() => { editor.inlineTextFormat = null; focused = true; }} onblur={() => { commit(); setTimeout(() => { const element = input?.getElement?.(); if (!element || activeElementOf(element) !== element) focused = false; }, 0); }}
    onselect={(next) => { const paragraph = paragraphIndexAt(next.start); const pendingEmptyParagraph = next.start === next.end && typingFormat && pendingParagraphs.has(paragraph) && paragraphIsEmptyAt(next.start); if ((next.start !== range.start || next.end !== range.end) && !pendingEmptyParagraph) { typingFormat = undefined; } range = next; formatForRange(next); }} onbeforeinput={(next) => { range = next; formatForRange(next); }} oninput={changed} onkeydown={keys}
    onnewline={(kind) => {
      // A newly-created paragraph is empty until its first character is
      // typed. Capture the current paragraph-end format before applying the
      // newline so that pending typing in the new paragraph inherits the
      // literal scheme token immediately, even before the debounce commits.
      // `range` is the logical caret maintained by changed(). Derive the
      // paragraph from it so Enter also works when the user moves back into
      // an earlier pending paragraph.
      const paragraph = currentParagraphEnd() + 1;
      const inherited = typingFormat?.format
        ? { ...typingFormat.format }
        : toWritableTextFormat(getSlideNotesParagraphEndFormat(slide, currentParagraphEnd()));
      const nextCaret = range.start + 1;
      const nextValue = `${value.slice(0, range.start)}\n${value.slice(range.end)}`;
      const change = textEditDiff(value, nextValue, range, nextCaret) ?? {
        start: range.start,
        end: range.end,
        text: value.slice(range.start, range.end),
      };
      change.newlines = kind;
      if (typingFormat) change.typing = typingFormat;
      reindexPendingParagraphs(change, kind === 'paragraph');
      softBreaks = new Set(
        [...softBreaks]
          .filter((position) => position < range.start || position >= range.end)
          .map((position) => position >= range.end ? position + 1 - (range.end - range.start) : position),
      );
      if (kind === 'break') softBreaks.add(range.start);
      changes.push(change);
      value = nextValue;
      range = { start: nextCaret, end: nextCaret };
      pending = changes.length > 0;
      clearTimeout(timer);
      if (!composing) timer = setTimeout(commit, 600);
      // The contenteditable DOM is not updated until Svelte flushes. Keep its
      // native caret aligned immediately so a second Enter or typed character
      // cannot be applied at the pre-newline position.
      input.setSelectionRange(nextCaret, nextCaret);
      typingFormat = { format: inherited, reset: false };
      // Paragraph indexes after the insertion move by one. Keep existing
      // pending paragraphs addressable when Enter is pressed in an earlier
      // empty paragraph before them.
      if (kind === 'paragraph') pendingParagraphs.add(paragraph);
    }} oncomposition={(active) => { composing = active; if (active) clearTimeout(timer); else if (pending) timer = setTimeout(commit, 600); }}
    onhistory={(backward) => { commit(); typingFormat = undefined; pendingParagraphs.clear(); void (backward ? doc.undo() : doc.redo()); }} oncopy={() => {}} oncut={() => {}} onpaste={pasteNotes} />
  <!-- PowerPoint's empty-notes prompt; clicks fall through to the text box. -->
  {#if value === '' && !focused}<span class="placeholder" aria-hidden="true">{t('Click to add notes')}</span>{/if}
</section>

<style>
  .placeholder { position: absolute; top: 12px; left: 18px; color: var(--ok-text-3); font-size: 14px; pointer-events: none; }
  .notes-pane { position: relative; min-height: 44px; max-height: 50vh; box-sizing: border-box; border-top: 1px solid var(--ok-border); background: var(--ok-panel); padding: 10px 16px 6px; }
  .resize { position: absolute; left: 0; right: 0; top: -3px; height: 6px; cursor: ns-resize; touch-action: none; }
  .resize:focus-visible { outline: 2px solid var(--ok-accent); }
</style>

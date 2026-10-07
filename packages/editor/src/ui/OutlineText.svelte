<script lang="ts">
  import { mergeTextFormat } from '../core/merge-text-format.ts';
  import { onDestroy, onMount, tick, untrack } from 'svelte';
  import { getShapeText, getParagraphLevel, setParagraphLevel, getSlides, getSlideLayout, setShapeText, setShapeParagraphs, findShapeById, getSlidePartName, setShapeTextFormat, getShapeParagraphCount, getShapeParagraphElements, getParagraphPropertiesEffective, type TextCase, type TextFormat } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { deleteOutlineRange, splitOutlineTitleRange, splitOutlineTitle, outlineSlots, outlineDemotionNeedsConfirmation, promoteOutlineBody, demoteOutlineTitle, outlineParagraphMove, outlineTitleMove, moveOutlineTitle } from '../core/outline.ts';
  import { textEditDiff } from '../core/text-edit-diff.ts';
  import { projectTextEdits, replayTextEdits, type TextEdit } from '../core/text-edit-preview.ts';
  import { copyTextRange, parseTextClipboard, TEXT_CLIPBOARD_TYPE } from '../core/text-clipboard.ts';
  import { parseHtmlTextClipboard, textClipboardHtml } from '../core/html-text-clipboard.ts';
  import RichTextInput from './RichTextInput.svelte';
  import { outlineTextHtml } from '../core/outline-text-html.ts';
  import { richTextValue, selectRichText } from '../core/rich-text-dom.ts';
  import { OutlineSelectionModel, type OutlineSelectionField, type OutlineClipboard, type OutlinePoint, type OutlineRange } from '../core/outline-selection.ts';
  import { textCaseSelection } from '../core/text-case.ts';
  import { textFormatsInRange } from '../core/text-format-selection.ts';
  import { stepFontSize, stepShapeFontSize } from '../core/font-size.ts';
  import { defaultTextMetrics } from '../core/text-layout-defaults.ts';
  import { paragraphsInTextRange } from '../core/paragraph-selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let { slideIndex, shapeId, title, selection }: { slideIndex: number; shapeId: number; title: boolean; selection: OutlineSelectionModel } = $props();
  const editor = getEditor();
  const doc = editor.doc;
  const presentation = untrack(() => doc.pres);
  const slide = untrack(() => doc.slideAt(slideIndex)!);
  let value = $state(untrack(() => getShapeText(doc.shapeById(slideIndex, shapeId)!)));
  let input: RichTextInput;
  // Blur and destruction can commit within the same reactive teardown batch.
  // Keep the queue synchronous so that the second callback sees it drained.
  let changes: TextEdit[] = [];
  let draftVersion = $state(0);
  const html = $derived.by(() => {
    doc.version;
    draftVersion;
    const source = doc.shapeById(slideIndex, shapeId);
    return source ? outlineTextHtml(doc.pres, projectTextEdits(source, changes, undefined, doc.pres), title, editor.outlineShowFormatting, source) : '';
  });
  let range = { start: 0, end: 0 };
  let timer: ReturnType<typeof setTimeout> | undefined;
  let composing = false;
  let titleBodyComposition = false;
  let deletionDialog = $state<HTMLDialogElement>();
  let pendingDeletion: (() => void) | undefined;
  let demotionDialog = $state<HTMLDialogElement>();
  let demotionVersion = 0;
  let selectionField: OutlineSelectionField;
  let typingFormat: { format: TextFormat; reset: boolean } | undefined;

  function flushDraft(): readonly TextEdit[] {
    if (!changes.length) return [];
    const edits = changes;
    changes = [];
    draftVersion++;
    return edits;
  }

  function commit() {
    if (selection.transitioning()) {
      // Keep the existing timer: a focus change across outline fields can
      // blur a draft while the shared native selection is being installed,
      // and clearing it here would strand that draft.
      return;
    }
    clearTimeout(timer);
    const edits = flushDraft();
    if (!edits.length) return;
    // History restoration and external reloads replace the model; an old draft
    // must never be applied to a replacement that happens to reuse shape IDs.
    if (doc.pres !== presentation) return;
    const index = getSlides(presentation).indexOf(slide);
    const shape = index < 0 ? null : doc.shapeById(index, shapeId);
    if (shape) doc.transact(t('Edit text'), () => replayTextEdits(shape, edits));
  }
  function rememberRange() { range = input.getSelection(); }
  function changed(next: string) {
    if (titleBodyComposition) return;
    const change = textEditDiff(value, next, range, input.getSelection().start);
    if (change) {
      if (typingFormat && change.text.length) change.typing = typingFormat;
      changes.push(change); draftVersion++;
    }
    value = next;
    rememberRange();
    clearTimeout(timer);
    if (!composing) timer = setTimeout(commit, 600);
  }
  function replaceSelection(text: string, formats?: TextEdit['formats']) {
    rememberRange();
    const { start, end } = range;
    changes.push({ start, end, text, formats });
    draftVersion++;
    value = value.slice(0, start) + text + value.slice(end);
    // The shared editor restores these offsets after rendering the new paragraphs.
    range = { start: start + text.length, end: start + text.length };
    input.setSelectionRange(range.start, range.end);
    clearTimeout(timer);
    timer = setTimeout(commit, 600);
  }
  /**
   * Opening the confirmation modal blurs the editor. Commit drafts before
   * taking the version snapshot so that this blur cannot invalidate it.
   */
  function commitDrafts(): number {
    const drafts = selection.fields().map(field => ({ field, edits: field.flush() }))
      .filter(draft => draft.edits.length);
    if (drafts.length) doc.transact(t('Edit text'), () => {
      for (const { field, edits } of drafts) field.apply(edits);
    });
    return doc.version;
  }
  /**
   * Replace the shared range in one history step. A range that crosses a slide
   * title merges slides; `split` then applies Enter at the caret, which starts
   * a new slide when the caret is in a title.
   */
  async function replaceOutlineRange(text: string, label: string, options: { formats?: OutlineClipboard['formats']; confirmed?: boolean; range?: OutlineRange; split?: boolean } = {}) {
    const { formats = [], confirmed = false, split = false } = options;
    const selected = options.range ?? selection.current();
    if (!selected || selected.start.key === selected.end.key) {
      selection.replace(text, formats, label);
      return;
    }
    const slots = outlineSlots(doc.pres);
    const start = slots.find(item => item.key === selected.start.key);
    const end = slots.find(item => item.key === selected.end.key);
    const joinsSlides = start && end && end.index > start.index;
    if (start && end && ((start.title && start.slide === end.slide && !end.title) || joinsSlides)) {
      const affected = getSlides(doc.pres).slice(start.index, end.index + 1);
      if (joinsSlides && !confirmed && affected.slice(1).some(outlineDemotionNeedsConfirmation)) {
        const version = commitDrafts();
        pendingDeletion = () => {
          if (doc.pres === presentation && doc.version === version) void replaceOutlineRange(text, label, { formats, confirmed: true, range: selected, split });
        };
        deletionDialog?.showModal();
        return;
      }
      const parts = new Set(affected.map(source => getSlidePartName(source)));

      // The transaction can remove this field's slide and unmount it.
      const owner = input.getElement()?.ownerDocument;
      let key = start.key;
      let caret = selected.start.offset + text.length;
      doc.transact(label, () => {
        for (const field of selection.fields()) {
          if (parts.has(field.key.slice(0, field.key.lastIndexOf(':')))) field.apply(field.flush());
        }
        deleteOutlineRange(doc.pres,
          { slide: start.slide, id: start.id, offset: selected.start.offset },
          { slide: end.slide, id: end.id, offset: selected.end.offset });
        const head = findShapeById(start.slide, start.id)!;
        const index = split && start.title
          ? splitOutlineTitle(doc.pres, start.slide, head, { start: selected.start.offset, end: selected.start.offset })
          : null;
        if (index !== null) {
          const next = getSlides(doc.pres)[index]!;
          key = outlineSlots(doc.pres).find(item => item.slide === next && item.title)?.key ?? key;
          caret = 0;
          doc.selectSlide(index);
        } else {
          if (text) replayTextEdits(head, [{ start: selected.start.offset, end: selected.start.offset, text, formats }]);
          doc.selectShape(start.index, start.id);
        }
      });
      selection.clear();
      await tick();
      const active = owner?.activeElement;
      const version = doc.version;
      const restore = () => {
        if (doc.version !== version || owner?.activeElement !== active) return;
        const field = selection.fields().find(item => item.key === key);
        if (field) {
          field.focus(caret);
          selection.setCaret(field, caret);
        }
      };
      if (owner?.defaultView) owner.defaultView.requestAnimationFrame(restore);
      else restore();
    } else selection.replace(text, formats, label);
  }
  /**
   * Backspace at the start or Delete at the end of a field joins it with the
   * neighbouring field when a title boundary lies between them, so the outline
   * behaves as one text (the joined slide is merged). Collapsed bodies take
   * part through the document, not the rendered fields.
   */
  function boundaryJoin(direction: -1 | 1): OutlineRange | null {
    const current = input.getSelection();
    const shared = selection.current();
    if (shared && shared.start.key !== shared.end.key) return null;
    if (current.start !== current.end || current.start !== (direction < 0 ? 0 : value.length)) return null;
    const slots = outlineSlots(doc.pres);
    const index = slots.findIndex(item => item.key === selectionField.key);
    const neighbour = slots[index + direction];
    if (index < 0 || !neighbour) return null;
    const [first, second] = direction < 0 ? [neighbour, slots[index]!] : [slots[index]!, neighbour];
    if (first.slide === second.slide && !(first.title && !second.title)) return null;
    const length = selection.fields().find(field => field.key === first.key)?.text().length
      ?? getShapeText(findShapeById(first.slide, first.id)!).length;
    return { start: { key: first.key, offset: length }, end: { key: second.key, offset: 0 } };
  }
  // Option-drag copies on the Mac; Control-drag copies elsewhere.
  function copyModifier(event: DragEvent): boolean { return event.altKey || event.ctrlKey; }
  /**
   * Dropping dragged outline text moves it, or copies it with the copy
   * modifier, in one history step. A moved range that crosses a slide title
   * merges slides exactly as deleting it would, with the same confirmation.
   */
  async function dropText(event: DragEvent) {
    const drag = selection.textDrag();
    if (!drag) return;
    event.preventDefault(); event.stopPropagation();
    selection.endTextDrag();
    const point = selection.pointAt(event.clientX, event.clientY);
    const copying = copyModifier(event);
    if (!point || (!copying && selection.contains(point))) return;
    const slots = outlineSlots(doc.pres);
    const locate = (target: OutlinePoint) => {
      const slot = slots.find(item => item.key === target.key);
      return slot && { slide: slot.slide, id: slot.id, offset: target.offset, index: slot.index, key: slot.key };
    };
    const start = locate(drag.range.start);
    const end = locate(drag.range.end);
    const destination = locate(point);
    if (!start || !end || !destination) return;
    const after = selection.precedes(drag.range.end, point);
    const apply = () => {
      doc.transact(t(copying ? 'Copy' : 'Move text'), () => {
        for (const field of selection.fields()) field.apply(field.flush());
        const insert = () => replayTextEdits(findShapeById(destination.slide, destination.id)!,
          [{ start: destination.offset, end: destination.offset, text: drag.clipboard.text, formats: drag.clipboard.formats }]);
        if (copying) insert();
        else if (after) { insert(); deleteOutlineRange(doc.pres, start, end); }
        else { deleteOutlineRange(doc.pres, start, end); insert(); }
      });
      selection.clear();
    };
    if (!copying && end.index > start.index &&
      getSlides(doc.pres).slice(start.index + 1, end.index + 1).some(outlineDemotionNeedsConfirmation)) {
      const version = commitDrafts();
      pendingDeletion = () => { if (doc.pres === presentation && doc.version === version) apply(); };
      deletionDialog?.showModal();
      return;
    }
    apply();
    // Put the caret after the dropped text when its position is unaffected by
    // the deletion: a drop before the range, on a later slide, or within the
    // range's own field (which only shifts by the removed length).
    let caret: number | null = destination.offset + drag.clipboard.text.length;
    if (!copying && after && destination.index <= end.index) {
      caret = start.key === end.key && destination.key === end.key ? caret - (end.offset - start.offset) : null;
    }
    await tick();
    const field = selection.fields().find(item => item.key === destination.key);
    if (field && caret !== null) {
      field.focus(caret);
      selection.setCaret(field, caret);
    }
  }
  function copy(event: ClipboardEvent, cut = false) {
    if (!event.clipboardData || composing) return;
    const copied = selection.copy();
    if (!copied || copied.text.length === 0) return;
    event.clipboardData.setData('text/plain', copied.text);
    event.clipboardData.setData('text/html', textClipboardHtml(copied));
    event.clipboardData.setData(TEXT_CLIPBOARD_TYPE, JSON.stringify({ version: 1, ...copied }));
    event.preventDefault(); event.stopPropagation();
    if (cut) void replaceOutlineRange('', t('Cut'));
  }
  function paste(event: ClipboardEvent) {
    if (!event.clipboardData || composing) return;
    const plain = event.clipboardData.getData('text/plain');
    const copied = parseTextClipboard(event.clipboardData.getData(TEXT_CLIPBOARD_TYPE), plain)
      ?? parseHtmlTextClipboard(event.clipboardData.getData('text/html'), plain);
    event.preventDefault(); event.stopPropagation();
    void replaceOutlineRange(copied?.text ?? plain, t('Paste'), { formats: copied?.formats });
  }
  async function menuClipboard(action: 'copy' | 'cut' | 'paste') {
    if (composing) return;
    const target = input.getElement()!;
    const selectedRange = { ...range };
    const version = doc.version;
    const original = value;
    // Clipboard permission may resolve after navigation, Undo or another edit.
    const current = () => target.isConnected && doc.pres === presentation && doc.version === version
      && target.ownerDocument.activeElement === target && value === original
      && input.getSelection().start === selectedRange.start && input.getSelection().end === selectedRange.end;
    try {
      if (action === 'paste') {
        const items = await navigator.clipboard.read();
        const item = items.find(item => item.types.includes('text/plain') || item.types.includes('text/html'));
        if (!item) return;
        const [plain, html] = await Promise.all(['text/plain', 'text/html'].map(async (mimeType) =>
          item.types.includes(mimeType) ? (await item.getType(mimeType)).text() : ''));
        if (!current()) return;
        const copied = parseHtmlTextClipboard(html!, plain!);
        await replaceOutlineRange(copied?.text ?? plain!, t('Paste'), { formats: copied?.formats });
      } else {
        if (selectedRange.start === selectedRange.end && !selection.current()) return;
        const copied = selection.copy();
        if (!copied) return;
        await navigator.clipboard.write([new ClipboardItem({
          'text/plain': new Blob([copied.text], { type: 'text/plain' }),
          'text/html': new Blob([textClipboardHtml(copied)], { type: 'text/html' }),
        })]);
        if (action === 'cut' && current()) await replaceOutlineRange('', t('Cut'));
      }
    } catch (error) { editor.toast('error', error instanceof Error ? error.message : String(error)); }
  }
  async function changeLevel(promote: boolean, confirmed = false) {
    rememberRange(); commit();
    const source = doc.shapeById(slideIndex, shapeId)!;
    const ownerDocument = input.getElement()!.ownerDocument;
    let focusIndex = slideIndex;
    let focusBody = false;
    if (title && (promote || slideIndex === 0)) return;
    if (title && !confirmed && outlineDemotionNeedsConfirmation(slide)) {
      demotionVersion = doc.version;
      demotionDialog?.showModal();
      return;
    }
    if (title || promote || getParagraphLevel(source, range).some(level => level < 8)) {
      try { doc.transact(t(promote ? 'Promote' : 'Demote'), () => {
        if (title) {
          const target = demoteOutlineTitle(doc.pres, slide);
          if (target) {
            focusIndex = slideIndex - 1;
            focusBody = true;
            doc.selectSlide(focusIndex);
          }
        } else if (promote) {
          const added = promoteOutlineBody(doc.pres, slide, source, range);
          if (added.length) {
            focusIndex = getSlides(doc.pres).indexOf(added[0]!);
            doc.selectSlide(focusIndex);
          }
        } else setParagraphLevel(source, range, { offset: 1 });
      }); } catch (error) { editor.toast('error', error instanceof Error ? error.message : String(error)); }
    }
    await tick();
    if (focusIndex !== slideIndex) ownerDocument.querySelector<HTMLElement>(`[data-outline-slide="${focusIndex}"] [role="textbox"]${focusBody ? ':has([data-outline-paragraph]:not([data-outline-title]))' : ''}`)?.focus();
    else input?.focus();
  }
  function confirmDemotion() {
    demotionDialog?.close();
    if (doc.pres === presentation && doc.version === demotionVersion) void changeLevel(false, true);
  }
  function moveAcrossTextboxes(direction: -1 | 1): boolean {
    const current = input.getElement();
    if (!current) return false;
    const pane = current.closest<HTMLElement>('.outline-pane');
    const textboxes = pane
      ? [...pane.querySelectorAll<HTMLElement>('[role="textbox"]')]
      : [];
    const position = textboxes.indexOf(current);
    const next = textboxes[position + direction];
    if (!next) return false;
    commit();
    next.focus();
    const offset = direction < 0 ? richTextValue(next).length : 0;
    selectRichText(next, offset);
    // RichTextInput keeps its UTF-16 range in sync from selectionchange. Dispatch
    // explicitly because focus followed by a programmatic range does not fire it
    // in every browser.
    next.ownerDocument.dispatchEvent(new Event('selectionchange'));
    return true;
  }
  async function moveParagraph(direction: -1 | 1) {
    rememberRange(); commit();
    const source = doc.shapeById(slideIndex, shapeId)!;
    if (title) {
      doc.transact(t(direction === -1 ? 'Move Up' : 'Move Down'), () => moveOutlineTitle(doc.pres, slide, direction));
      await tick();
      input.focus();
      input.setSelectionRange(0, value.length);
      rememberRange();
      return;
    }
    const move = outlineParagraphMove(source, range, direction);
    if (!move) return;
    doc.transact(t(direction === -1 ? 'Move Up' : 'Move Down'), () => setShapeParagraphs(source, { source, ranges: move.ranges }));
    await tick();
    input.focus();
    input.setSelectionRange(move.selection.start, move.selection.end);
    rememberRange();
  }
  function context(event: MouseEvent) {
    event.preventDefault(); event.stopPropagation();
    rememberRange(); commit();
    editor.openContextMenu(event.clientX, event.clientY, { source: 'outline', outlineText: {
      moveUp: () => { void moveParagraph(-1); },
      moveDown: () => { void moveParagraph(1); },
      canMoveUp: title ? outlineTitleMove(doc.pres, slide, -1) !== null : outlineParagraphMove(doc.shapeById(slideIndex, shapeId)!, range, -1) !== null,
      canMoveDown: title ? outlineTitleMove(doc.pres, slide, 1) !== null : outlineParagraphMove(doc.shapeById(slideIndex, shapeId)!, range, 1) !== null,
      promote: () => { void changeLevel(true); },
      demote: () => { void changeLevel(false); },
      copy: () => { void menuClipboard('copy'); },
      cut: () => { void menuClipboard('cut'); },
      paste: () => { void menuClipboard('paste'); },
      hasTextSelection: range.start !== range.end || (() => { const selected = selection.current(); return !!selected && selected.start.key !== selected.end.key; })(),
      canPromote: !title,
      canDemote: !title || slideIndex > 0,
      hyperlink: () => {
        if (range.start === range.end) return;
        const linked = { ...range };
        commit();
        doc.selectShape(slideIndex, shapeId);
        editor.linkTextRange = linked;
        editor.linkTableCell = null;
        editor.runOrPrompt('setShapeHyperlink');
      },
    } });
  }
  async function keys(event: KeyboardEvent) {
    if (event.isComposing) return;
    const mod = event.metaKey || event.ctrlKey;
    const join = !mod && !event.altKey && !event.shiftKey && (event.key === 'Backspace' || event.key === 'Delete')
      ? boundaryJoin(event.key === 'Backspace' ? -1 : 1) : null;
    if (mod && !event.altKey && !event.shiftKey && event.key.toLowerCase() === 't') {
      event.preventDefault();
      event.stopPropagation();
      editor.openFontDialog('font', input.getElement());
      return;
    }
    if (mod && event.shiftKey && !event.altKey &&
      (event.code === 'Period' || event.code === 'Comma' || event.key === '>' || event.key === '<')) {
      event.preventDefault();
      event.stopPropagation();
      selection.fontSize(event.code === 'Comma' || event.key === '<' ? -1 : 1);
      return;
    }
    if (mod && event.key.toLowerCase() === 's') commit();
    else if (event.key === 'Escape') { commit(); input.blur(); }
    else if (!mod && !event.shiftKey && !event.altKey &&
      (event.key === 'ArrowLeft' || event.key === 'ArrowRight') &&
      selection.current()?.start.key !== selection.current()?.end.key) {
      // Horizontal arrows collapse a native multi-field selection to its
      // corresponding edge. Clear the logical bridge so the next printable
      // input edits the collapsed caret instead of replacing the old range.
      event.preventDefault();
      event.stopPropagation();
      selection.collapse(event.key === 'ArrowLeft' ? -1 : 1);
    }
    else if (join) {
      event.preventDefault(); event.stopPropagation();
      await replaceOutlineRange('', t('Delete'), { range: join });
    }
    else if ((event.key === 'Backspace' || event.key === 'Delete') && selection.current()?.start.key !== selection.current()?.end.key) {
      event.preventDefault(); event.stopPropagation();
      await replaceOutlineRange('', t('Delete'));
    }
    else if (!mod && !event.altKey && event.key.length === 1 &&
      selection.current()?.start.key !== selection.current()?.end.key) {
      // The browser can only edit the focused contenteditable.  Route a
      // cross-field replacement through the shared transaction instead of
      // letting it insert into whichever field happens to own focus.
      event.preventDefault(); event.stopPropagation();
      await replaceOutlineRange(event.key, t('Edit text'));
    }
    else if (event.shiftKey && !mod && !event.altKey && (event.key === 'ArrowUp' || event.key === 'ArrowDown')) {
      const current = input.getSelection();
      // A programmatic selection (including a selection restored after a
      // draft render) may not have delivered selectionchange yet. Seed the
      // shared anchor from the focused field before extending it.
      if (selection.focusOffset(selectionField) === null ||
        selection.current()?.start.key === selection.current()?.end.key)
        selection.update(selectionField, current.start, current.end);
      const focus = selection.focusOffset(selectionField) ?? (event.key === 'ArrowUp' ? current.start : current.end);
      const crossField = selection.current()?.start.key !== selection.current()?.end.key;
      const boundary = event.key === 'ArrowUp'
        ? focus === 0 || (crossField && focus === value.length)
        : focus === value.length || (crossField && focus === 0);
      const extendField = selection.focusedField() ?? selectionField;
      if (boundary && selection.extend(extendField, event.key === 'ArrowUp' ? -1 : 1)) {
        event.preventDefault(); event.stopPropagation();
      }
    }
    else if (
      !mod &&
      !event.shiftKey &&
      !event.altKey &&
      (event.key === 'ArrowUp' || event.key === 'ArrowDown')
    ) {
      const current = input.getSelection();
      const boundary =
        event.key === 'ArrowUp'
          ? current.start === 0 && current.end === 0
          : current.start === value.length && current.end === value.length;
      if (boundary && moveAcrossTextboxes(event.key === 'ArrowUp' ? -1 : 1)) {
        event.preventDefault();
        event.stopPropagation();
      }
    }
    else if (event.key === 'Tab' && !mod && !event.altKey) {
      event.preventDefault(); event.stopPropagation();
      await changeLevel(event.shiftKey);
    }
    else if (event.key === 'Enter' && !mod && !event.altKey &&
      selection.current()?.start.key !== selection.current()?.end.key) {
      event.preventDefault(); event.stopPropagation();
      const selected = selection.current()!;
      const slots = outlineSlots(doc.pres);
      const start = slots.find(item => item.key === selected.start.key && item.title);
      const end = slots.find(item => item.key === selected.end.key);
      const sameSlideBody = start && end && start.slide === end.slide && !end.title && getSlideLayout(start.slide);
      const adjacentTitle = start && end && end.index === start.index + 1 && end.title;
      if (start && end && (sameSlideBody || adjacentTitle)) {
        const ownerDocument = input.getElement()!.ownerDocument;
        let index: number | null = null;
        doc.transact(t('New slide'), () => {
          for (const field of selection.fields()) {
            if (field.key.startsWith(`${getSlidePartName(start.slide)}:`) ||
              field.key.startsWith(`${getSlidePartName(end.slide)}:`)) field.apply(field.flush());
          }
          index = splitOutlineTitleRange(doc.pres, start.slide,
            { id: start.id, offset: selected.start.offset },
            { id: end.id, offset: selected.end.offset, slide: end.slide });
          if (index !== null) doc.selectSlide(index);
        });
        selection.clear();
        await tick();
        ownerDocument.querySelector<HTMLElement>(`[data-outline-slide="${index}"] [role="textbox"]`)?.focus();
      } else await replaceOutlineRange('\n', t('Edit text'), { split: true });
    }
    // Mac PowerPoint splits outline titles into slides for both Enter and Shift+Enter.
    else if (event.key === 'Enter' && !mod && !event.altKey && title) {
      if (!getSlideLayout(slide)) return;
      const ownerDocument = input.getElement()!.ownerDocument;
      const selected = input.getSelection();
      event.preventDefault(); event.stopPropagation(); commit();
      const source = doc.shapeById(slideIndex, shapeId)!;
      let index: number | null = null;
      doc.transact(t('New slide'), () => {
        index = splitOutlineTitle(doc.pres, slide, source, selected);
        if (index !== null) doc.selectSlide(index);
      });
      await tick();
      ownerDocument.querySelector<HTMLElement>(`[data-outline-slide="${index}"] [role="textbox"]`)?.focus();
    }
  }
  $effect(() => {
    doc.version;
    if (!changes.length && doc.pres === presentation) {
      const shape = doc.shapeById(slideIndex, shapeId);
      if (shape) value = getShapeText(shape);
    }
  });
  onMount(() => {
    selectionField = {
      key: `${getSlidePartName(slide)}:${shapeId}`,
      root: input.getElement()!,
      text: () => value,
      copy: (start, end) => {
        const shape = doc.shapeById(slideIndex, shapeId)!;
        return copyTextRange(projectTextEdits(shape, changes, undefined, doc.pres), start, end);
      },
      flush: flushDraft,
      apply: edits => {
        const shape = doc.shapeById(slideIndex, shapeId);
        if (shape) replayTextEdits(shape, edits);
      },
      formats: (start, end) => {
        const shape = doc.shapeById(slideIndex, shapeId);
        const formats = shape ? textFormatsInRange(projectTextEdits(shape, changes, undefined, doc.pres), { start, end }, undefined, { pres: doc.pres, source: shape }) : [];
        if (start === end && typingFormat) return [mergeTextFormat(typingFormat.reset ? undefined : formats[0], typingFormat.format)];
        return formats;
      },
      applyFormat: (start, end, format, reset) => {
        const shape = doc.shapeById(slideIndex, shapeId);
        if (start === end) {
          // Keep the same pending-format merge semantics as canvas editing:
          // successive toolbar commands (for example Bold then Italic) apply
          // to the same future input, while reset starts a fresh format.
          typingFormat = {
            format: mergeTextFormat(reset ? undefined : typingFormat?.format, format),
            reset: reset || typingFormat?.reset || false,
          };
          draftVersion++;
        } else if (shape) {
          typingFormat = undefined;
          setShapeTextFormat(shape, format, { range: { start, end }, reset });
        }
      },
      changeCase: (start, end, caseValue: TextCase, caret?: number) => {
        const shape = doc.shapeById(slideIndex, shapeId);
        if (!shape) return end;
        typingFormat = undefined;
        const before = getShapeText(shape);
        setShapeText(shape, { case: caseValue }, { range: { start, end } });
        const after = getShapeText(shape);
        value = after;
        const nextRange = textCaseSelection(
          before,
          after,
          caret === undefined ? { start: end, end } : { start: caret, end: caret },
          { start, end },
        );
        range = nextRange;
        draftVersion++;
        return range.end;
      },
      applyFontSize: (start, end, direction) => {
        const source = doc.shapeById(slideIndex, shapeId);
        if (!source) return;
        if (start === end) {
          const current = typingFormat?.format.size
            ?? textFormatsInRange(source, { start, end }, undefined, { pres: doc.pres, source })[0]?.size
            ?? defaultTextMetrics(doc.pres, source).size;
          typingFormat = { format: { ...(typingFormat?.format ?? {}), size: stepFontSize(current, direction) }, reset: typingFormat?.reset ?? false };
          draftVersion++;
          return;
        }
        typingFormat = undefined;
        stepShapeFontSize(doc.pres, source, direction, source, { start, end });
        draftVersion++;
      },
      paragraphs: (start, end) => {
        const shape = doc.shapeById(slideIndex, shapeId);
        if (!shape) return [];
        const lengths = Array.from({ length: getShapeParagraphCount(shape) }, (_, index) =>
          getShapeParagraphElements(shape, index).reduce((total, element) => total + (element.kind === 'br' ? 1 : element.text.length), 0));
        return paragraphsInTextRange(lengths, { start, end }).map(index => getParagraphPropertiesEffective(doc.pres, shape, index));
      },
      editParagraphs: (start, end, edit) => {
        const shape = doc.shapeById(slideIndex, shapeId);
        if (!shape) return;
        const lengths = Array.from({ length: getShapeParagraphCount(shape) }, (_, index) =>
          getShapeParagraphElements(shape, index).reduce((total, element) => total + (element.kind === 'br' ? 1 : element.text.length), 0));
        for (const index of paragraphsInTextRange(lengths, { start, end })) edit(shape, index);
      },
      transact: (label, fn) => doc.transact(label, fn),
      focus: offset => { input.focus(); input.setSelectionRange(offset, offset); },
      setRange: offset => { range = { start: offset, end: offset }; },
      select: (start, end) => { input.focus(); input.setSelectionRange(start, end); range = { start, end }; },
      get title() { return title; },
      changeLevel,
    };
    const root = input.getElement()!;
    const dragOver = (event: DragEvent) => {
      if (!selection.textDrag()) return;
      event.preventDefault();
      event.stopPropagation();
      if (event.dataTransfer) event.dataTransfer.dropEffect = copyModifier(event) ? 'copy' : 'move';
    };
    const drop = (event: DragEvent) => { void dropText(event); };
    root.addEventListener('dragover', dragOver);
    root.addEventListener('drop', drop);
    const unregister = selection.register(selectionField);
    return () => {
      root.removeEventListener('dragover', dragOver);
      root.removeEventListener('drop', drop);
      unregister();
    };
  });
  onDestroy(() => untrack(commit));
</script>

<RichTextInput bind:this={input} {value} {html} layout="outline" formatted={editor.outlineShowFormatting} label={`${t(title ? 'Outline title' : 'Outline text')} ${slideIndex + 1}`} style={editor.outlineShowFormatting ? "line-height: normal; min-height: 0" : ""} textZoom={1}
  onfocus={() => doc.selectShape(slideIndex, shapeId)} onpointerdown={event => { if (event.button === 0) { typingFormat = undefined; selection.clear(); } }} onbeforeinput={(next, event) => { if (titleBodyComposition) return; range = next; selectionField && selection.update(selectionField, next.start, next.end); if (event?.inputType === 'insertText' && event.data && selection.current()?.start.key !== selection.current()?.end.key) { event.preventDefault(); void replaceOutlineRange(event.data, t('Edit text')); } }} onselect={next => { if (titleBodyComposition) return; const element = input.getElement(); if (!element || element.ownerDocument.activeElement !== element || doc.selection.kind !== 'shape' || doc.selection.slideIndex !== slideIndex || !doc.selection.shapeIds.includes(shapeId)) return; if (next.start !== range.start || next.end !== range.end) typingFormat = undefined; range = next; selectionField && selection.update(selectionField, next.start, next.end); }}
  oninput={changed} onblur={commit} onkeydown={keys} oncontextmenu={context}
  oncopy={event => copy(event)} oncut={event => copy(event, true)} onpaste={paste}
  onnewline={() => {
    // Mac PowerPoint outline bodies use paragraphs for both Enter and Shift+Enter.
    if (selection.current()?.start.key !== selection.current()?.end.key) selection.replace('\n', [], t('Edit text'));
    else replaceSelection('\n');
  }}
  oncomposition={(active, text) => {
    composing = active;
    if (active) {
      clearTimeout(timer);
      const selected = selection.current();
      const slots = outlineSlots(doc.pres);
      const start = slots.find(item => item.key === selected?.start.key);
      const end = slots.find(item => item.key === selected?.end.key);
      titleBodyComposition = !!start && !!end && start.key !== end.key
        && ((start.index === end.index && start.title && !end.title) || end.index > start.index);
      // Keep the IME inside one editable root; native cross-root deletion can cancel composition.
      if (titleBodyComposition) input.setSelectionRange(range.start, range.start);
    } else if (titleBodyComposition) {
      titleBodyComposition = false;
      // The IME owns the temporary DOM; commit only its final text against the original range.
      draftVersion++;
      if (text) void replaceOutlineRange(text, t('Edit text'));
    } else timer = setTimeout(commit, 600);
  }}
  onhistory={backward => { commit(); void (backward ? doc.undo() : doc.redo()); }}
/>


<dialog bind:this={deletionDialog} aria-label={t('Delete')} onclose={() => { pendingDeletion = undefined; }}>
  <p>{t('This will delete the slide, its notes page and any graphics or media. Do you want to continue?')}</p>
  <footer><button onclick={() => deletionDialog?.close()}>{t('No')}</button><button class="confirm" onclick={() => { const action = pendingDeletion; deletionDialog?.close(); action?.(); }}>{t('Yes')}</button></footer>
</dialog>
{#if title}
  <dialog bind:this={demotionDialog} aria-label={t('Demote')}>
    <p>{t('This will delete the slide, its notes page and any graphics or media. Do you want to continue?')}</p>
    <footer><button onclick={() => demotionDialog?.close()}>{t('No')}</button><button class="confirm" onclick={confirmDemotion}>{t('Yes')}</button></footer>
  </dialog>
{/if}

<style>
  dialog { width: 260px; box-sizing: border-box; border: 1px solid #777; border-radius: 18px; background: #555; color: #fff; padding: 24px 16px 16px; box-shadow: 0 15px 60px #0008; font: 13px/1.35 Arial, sans-serif; }
  dialog::backdrop { background: #0003; }
  dialog p { margin: 0 0 16px; }
  dialog footer { display: flex; gap: 8px; }
  dialog button { flex: 1; border: 0; border-radius: 16px; padding: 6px 12px; background: #666; color: inherit; font: inherit; }
  dialog button.confirm { background: #b4440c; }

</style>

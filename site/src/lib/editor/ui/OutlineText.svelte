<script lang="ts">
  import { onDestroy, tick, untrack } from 'svelte';
  import { getShapeText, getParagraphLevel, setParagraphLevel, getSlides, getSlideLayout, addSlideAt, setShapeText, setShapeParagraphs, findShapeById, copyShape, removeShape } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { outlineShapes, outlineDemotionNeedsConfirmation, promoteOutlineBody, demoteOutlineTitle, outlineParagraphMove, outlineTitleMove, moveOutlineTitle } from '../core/outline.ts';
  import { textEditDiff } from '../core/text-edit-diff.ts';
  import { projectTextEdits, replayTextEdits, type TextEdit } from '../core/text-edit-preview.ts';
  import { copyTextRange, parseTextClipboard, TEXT_CLIPBOARD_TYPE } from '../core/text-clipboard.ts';
  import { parseHtmlTextClipboard, textClipboardHtml } from '../core/html-text-clipboard.ts';
  import RichTextInput from './RichTextInput.svelte';
  import { outlineTextHtml } from '../core/outline-text-html.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let { slideIndex, shapeId, title }: { slideIndex: number; shapeId: number; title: boolean } = $props();
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
    return source ? outlineTextHtml(doc.pres, projectTextEdits(source, changes, undefined, doc.pres), title, editor.outlineShowFormatting) : '';
  });
  let range = { start: 0, end: 0 };
  let timer: ReturnType<typeof setTimeout> | undefined;
  let composing = false;
  let demotionDialog = $state<HTMLDialogElement>();
  let demotionVersion = 0;

  function commit() {
    clearTimeout(timer);
    if (!changes.length) return;
    const edits = changes;
    changes = [];
    draftVersion++;
    // History restoration and external reloads replace the model; an old draft
    // must never be applied to a replacement that happens to reuse shape IDs.
    if (doc.pres !== presentation) return;
    const index = getSlides(presentation).indexOf(slide);
    const shape = index < 0 ? null : doc.shapeById(index, shapeId);
    if (shape) doc.transact(t('Edit text'), () => replayTextEdits(shape, edits));
  }
  function rememberRange() { range = input.getSelection(); }
  function changed(next: string) {
    const change = textEditDiff(value, next, range, input.getSelection().start);
    if (change) { changes.push(change); draftVersion++; }
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
  function copy(event: ClipboardEvent, cut = false) {
    if (!event.clipboardData || composing) return;
    rememberRange();
    if (range.start === range.end) return;
    const shape = doc.shapeById(slideIndex, shapeId)!;
    const copied = copyTextRange(projectTextEdits(shape, changes), range.start, range.end);
    event.clipboardData.setData('text/plain', copied.text);
    event.clipboardData.setData('text/html', textClipboardHtml(copied));
    event.clipboardData.setData(TEXT_CLIPBOARD_TYPE, JSON.stringify(copied));
    event.preventDefault(); event.stopPropagation();
    if (cut) replaceSelection('');
  }
  function paste(event: ClipboardEvent) {
    if (!event.clipboardData || composing) return;
    const plain = event.clipboardData.getData('text/plain');
    const copied = parseTextClipboard(event.clipboardData.getData(TEXT_CLIPBOARD_TYPE), plain)
      ?? parseHtmlTextClipboard(event.clipboardData.getData('text/html'), plain);
    event.preventDefault(); event.stopPropagation();
    replaceSelection(copied?.text ?? plain, copied?.formats);
  }
  async function menuClipboard(action: 'copy' | 'cut' | 'paste') {
    if (composing) return;
    const target = input.getElement()!;
    rememberRange();
    const selection = { ...range };
    const version = doc.version;
    const original = value;
    // Clipboard permission may resolve after navigation, Undo or another edit.
    const current = () => target.isConnected && doc.pres === presentation && doc.version === version
      && target.ownerDocument.activeElement === target && value === original
      && input.getSelection().start === selection.start && input.getSelection().end === selection.end;
    try {
      if (action === 'paste') {
        const items = await navigator.clipboard.read();
        const item = items.find(item => item.types.includes('text/plain') || item.types.includes('text/html'));
        if (!item) return;
        const [plain, html] = await Promise.all(['text/plain', 'text/html'].map(async (mimeType) =>
          item.types.includes(mimeType) ? (await item.getType(mimeType)).text() : ''));
        if (!current()) return;
        const copied = parseHtmlTextClipboard(html!, plain!);
        replaceSelection(copied?.text ?? plain!, copied?.formats);
        commit();
      } else {
        if (selection.start === selection.end) return;
        const shape = doc.shapeById(slideIndex, shapeId)!;
        const copied = copyTextRange(projectTextEdits(shape, changes), selection.start, selection.end);
        await navigator.clipboard.write([new ClipboardItem({
          'text/plain': new Blob([copied.text], { type: 'text/plain' }),
          'text/html': new Blob([textClipboardHtml(copied)], { type: 'text/html' }),
        })]);
        if (action === 'cut' && current()) { replaceSelection(''); commit(); }
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
    editor.openContextMenu(event.clientX, event.clientY, 'outline', {
      moveUp: () => { void moveParagraph(-1); },
      moveDown: () => { void moveParagraph(1); },
      canMoveUp: title ? outlineTitleMove(doc.pres, slide, -1) !== null : outlineParagraphMove(doc.shapeById(slideIndex, shapeId)!, range, -1) !== null,
      canMoveDown: title ? outlineTitleMove(doc.pres, slide, 1) !== null : outlineParagraphMove(doc.shapeById(slideIndex, shapeId)!, range, 1) !== null,
      promote: () => { void changeLevel(true); },
      demote: () => { void changeLevel(false); },
      copy: () => { void menuClipboard('copy'); },
      cut: () => { void menuClipboard('cut'); },
      paste: () => { void menuClipboard('paste'); },
      hasTextSelection: range.start !== range.end,
      canPromote: !title,
      canDemote: !title || slideIndex > 0,
    });
  }
  async function keys(event: KeyboardEvent) {
    if (event.isComposing) return;
    const mod = event.metaKey || event.ctrlKey;
    if (mod && event.key.toLowerCase() === 's') commit();
    else if (event.key === 'Escape') { commit(); input.blur(); }
    else if (event.key === 'Tab' && !mod && !event.altKey) {
      event.preventDefault(); event.stopPropagation();
      await changeLevel(event.shiftKey);
    }
    else if (event.key === 'Enter' && !event.shiftKey && !mod && !event.altKey && title) {
      const layout = getSlideLayout(slide);
      if (!layout) return;
      const ownerDocument = input.getElement()!.ownerDocument;
      const { start, end } = input.getSelection();
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
  onDestroy(() => untrack(commit));
</script>

<RichTextInput bind:this={input} {value} {html} layout="outline" label={`${t(title ? 'Outline title' : 'Outline text')} ${slideIndex + 1}`} style={editor.outlineShowFormatting ? "line-height: normal; min-height: 0" : ""} textZoom={1}
  onfocus={() => doc.selectShape(slideIndex, shapeId)} onbeforeinput={selection => range = selection} onselect={selection => range = selection}
  oninput={changed} onblur={commit} onkeydown={keys} oncontextmenu={context}
  oncopy={event => copy(event)} oncut={event => copy(event, true)} onpaste={paste}
  onnewline={() => replaceSelection('\n')}
  oncomposition={active => { composing = active; if (active) clearTimeout(timer); else timer = setTimeout(commit, 600); }}
  onhistory={backward => { commit(); void (backward ? doc.undo() : doc.redo()); }}
/>


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

<script lang="ts">
  import { onDestroy, untrack } from 'svelte';
  import { getSlideNotes, getSlides, setSlideNotes } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const slide = untrack(() => doc.currentSlide!);
  const presentation = untrack(() => doc.pres);
  let value = $state(getSlideNotes(slide) ?? '');
  let pending = false;
  let composing = false;
  let timer: ReturnType<typeof setTimeout> | undefined;
  let input: HTMLTextAreaElement;
  let pane: HTMLElement;
  let drag: { id: number; y: number; height: number } | null = null;
  let maxHeight = $state(400);

  function commit() {
    clearTimeout(timer);
    if (!pending) return;
    pending = false;
    // An external replacement or deleted slide must never receive a stale draft.
    if (doc.pres !== presentation || !getSlides(presentation).includes(slide)) return;
    if (value !== (getSlideNotes(slide) ?? '')) {
      try { doc.transact(t('Speaker notes'), () => setSlideNotes(slide, value)); }
      catch (error) { editor.toast('error', String(error)); }
    }
  }
  function changed() {
    pending = true;
    clearTimeout(timer);
    if (!composing) timer = setTimeout(commit, 600);
  }
  function keys(event: KeyboardEvent) {
    if (event.isComposing) return;
    const mod = event.metaKey || event.ctrlKey;
    if (mod && ['z', 'y'].includes(event.key.toLowerCase())) {
      event.preventDefault(); event.stopPropagation(); commit();
      void (event.shiftKey || event.key.toLowerCase() === 'y' ? doc.redo() : doc.undo());
    } else if (mod && event.key.toLowerCase() === 's') commit();
    else if (event.key === 'Escape') { commit(); input.blur(); }
  }
  $effect(() => {
    doc.version;
    if (!pending) value = getSlideNotes(slide) ?? '';
  });
  $effect(() => { if (editor.notesFocusRequest && input) { input.focus(); editor.notesFocusRequest = 0; } });
  onDestroy(() => untrack(commit));

  function resizeStart(event: PointerEvent) {
    if (event.button !== 0) return;
    event.preventDefault();
    maxHeight = Math.max(60, (pane.parentElement?.clientHeight ?? 600) - 100);
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
    maxHeight = Math.max(60, (pane.parentElement?.clientHeight ?? 600) - 100);
    editor.notesHeight = event.key === 'Home' ? 60 : event.key === 'End' ? maxHeight : Math.min(maxHeight, Math.max(60, editor.notesHeight + (event.key === 'ArrowUp' ? 10 : -10)));
  }
</script>

<section class="notes-pane" bind:this={pane} aria-label={t('Notes')} style:height="{editor.notesHeight}px">
  <!-- A focusable separator implements the ARIA window-splitter pattern. -->
  <!-- svelte-ignore a11y_no_noninteractive_tabindex, a11y_no_noninteractive_element_interactions -->
  <div class="resize" role="separator" tabindex="0" aria-label={t('Notes pane height')} aria-orientation="horizontal" aria-valuemin={60} aria-valuemax={maxHeight} aria-valuenow={editor.notesHeight} onpointerdown={resizeStart} onpointermove={resizeMove} onpointerup={() => drag = null} onpointercancel={() => drag = null} onlostpointercapture={() => drag = null} onkeydown={resizeKeys}></div>
  <textarea bind:this={input} bind:value aria-label={t('Notes content')} placeholder={t('Click to add notes')} oninput={changed} onblur={commit} onkeydown={keys} oncompositionstart={() => { composing = true; clearTimeout(timer); }} oncompositionend={() => { composing = false; changed(); }}></textarea>
</section>

<style>
  .notes-pane { position: relative; min-height: 60px; max-height: 50vh; box-sizing: border-box; border-top: 1px solid var(--ok-border); background: var(--ok-panel); padding: 10px 16px 8px; }
  textarea { display: block; box-sizing: border-box; width: 100%; height: 100%; resize: none; border: 0; outline: none; color: var(--ok-text); background: transparent; font: 14px/1.5 Arial, sans-serif; }
  textarea::placeholder { color: var(--ok-text-2); }
  .resize { position: absolute; left: 0; right: 0; top: -3px; height: 6px; cursor: ns-resize; touch-action: none; }
  .resize:focus-visible { outline: 2px solid var(--ok-accent); }
</style>

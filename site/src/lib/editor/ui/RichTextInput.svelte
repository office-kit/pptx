<script lang="ts">
  import { onMount, untrack } from 'svelte';
  import { richTextValue, richTextSelection, selectRichText, type TextSelection } from '../core/rich-text-dom.ts';
  let { value, html, label, style, zoom, busy = false, oninput, onselect, onbeforeinput, onkeydown, onnewline, oncomposition, onhistory, oncopy, oncut, onpaste }: {
    value: string; html: string; label: string; style: string; zoom: number; busy?: boolean;
    oninput: (value: string) => void;
    onselect: (range: TextSelection) => void;
    onbeforeinput: (range: TextSelection) => void;
    onkeydown: (event: KeyboardEvent) => void;
    onnewline: () => void;
    oncomposition: (active: boolean) => void;
    onhistory: (backward: boolean) => void;
    oncopy: (event: ClipboardEvent) => void;
    oncut: (event: ClipboardEvent) => void;
    onpaste: (event: ClipboardEvent) => void;
  } = $props();
  let element = $state<HTMLDivElement>();
  let composing = $state(false);
  let selection = { start: 0, end: 0 };
  let rendered = '';
  export function getSelection() { return element && document.activeElement === element ? richTextSelection(element) ?? selection : selection; }
  export function setSelectionRange(start: number, end: number) {
    selection = { start, end };
    if (element && document.activeElement === element) selectRichText(element, start, end);
  }
  export function focus() { element?.focus(); }
  export function select() { setSelectionRange(0, value.length); }
  function capture() {
    selection = getSelection();
    onbeforeinput(selection);
  }
  function changed() {
    selection = getSelection();
    if (!element) return;
    const next = richTextValue(element);
    oninput(next);
    rendered = '';
  }
  function selectionChanged() {
    if (!element || document.activeElement !== element) return;
    const range = richTextSelection(element);
    if (range) { selection = range; onselect(range); }
  }
  $effect(() => {
    const markup = html;
    value;
    if (!element || composing) return;
    const root = element;
    untrack(() => {
      if (rendered === markup) return;
      const active = document.activeElement === element;
      const range = { ...selection };
      root.innerHTML = markup;
      const wrapper = root.firstElementChild;
      if (wrapper?.tagName === 'DIV') wrapper.replaceWith(...wrapper.childNodes);
      // Clipboard sizes remain in points; only the editing view follows canvas zoom.
      for (const span of root.querySelectorAll('span')) {
        if (span.style.fontSize) span.style.fontSize = `calc(${span.style.fontSize} * var(--text-zoom))`;
        if (span.style.fontFamily) span.style.fontFamily += ', var(--ok-font)';
      }
      if ((!value || value.endsWith('\n')) && !root.querySelector('[data-text-paragraph]')) {
        const end = document.createElement('br');
        end.setAttribute('data-caret-end', '');
        root.append(end);
      }
      rendered = markup;
      if (active) setSelectionRange(Math.min(range.start, value.length), Math.min(range.end, value.length));
    });
  });
  onMount(() => {
    document.addEventListener('selectionchange', selectionChanged);
    focus(); setSelectionRange(value.length, value.length);
    return () => { document.removeEventListener('selectionchange', selectionChanged); };
  });
</script>

<div class="inline-edit" bind:this={element} contenteditable="true" role="textbox" tabindex="0" aria-multiline="true" aria-busy={busy} aria-label={label} style={`${style}; --text-zoom: ${zoom};`}
  onfocus={() => { if (element) selectRichText(element, selection.start, selection.end); }}
  onbeforeinput={event => {
    if (!composing && (event.inputType === 'historyUndo' || event.inputType === 'historyRedo')) {
      event.preventDefault();
      onhistory(event.inputType === 'historyUndo');
      return;
    }
    if (busy) { event.preventDefault(); return; }
    capture();
    if (!composing && (event.inputType === 'insertParagraph' || event.inputType === 'insertLineBreak')) {
      event.preventDefault();
      onnewline();
    }
  }}
  oninput={changed}
  oncompositionstart={() => { capture(); composing = true; oncomposition(true); }}
  oncompositionend={() => { composing = false; changed(); oncomposition(false); }}
  oncopy={oncopy} oncut={event => { if (busy) event.preventDefault(); else oncut(event); }} onpaste={event => { if (busy) event.preventDefault(); else onpaste(event); }}
  onpointerdown={event => event.stopPropagation()} onpointerup={event => event.stopPropagation()} ondblclick={event => event.stopPropagation()}
  onkeydown={event => {
    if (!event.isComposing && (event.ctrlKey || event.metaKey) && ['z', 'y'].includes(event.key.toLowerCase())) {
      event.preventDefault(); event.stopPropagation();
      onhistory(event.key.toLowerCase() === 'z' && !event.shiftKey);
      return;
    }
    if (busy) { event.preventDefault(); event.stopPropagation(); return; }
    onkeydown(event);
  }}
></div>

<style>
  .inline-edit { position: absolute; pointer-events: auto; border: 0; background: #fff; font-family: var(--ok-font); font-size: calc(14px * var(--text-zoom)); padding: calc(4px * var(--text-zoom)); z-index: 7; white-space: pre-wrap; overflow-wrap: break-word; overflow: auto; outline: 1px solid var(--ok-selected-border); }

  .inline-edit :global([data-list-marker]::before) {
    content: attr(data-list-marker);
    font-size: var(--marker-size, inherit);
    font-family: var(--marker-font, var(--ok-font));
    color: var(--marker-color, inherit);
    margin-inline-end: 0.4em;
    user-select: none;
  }
</style>

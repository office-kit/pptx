<script lang="ts">
  import { onMount, untrack } from 'svelte';
  import { richTextValue, richTextSelection, selectRichText, type TextSelection } from '../core/rich-text-dom.ts';
  let { value, html, label, style, oninput, onselect, onbeforeinput, onkeydown, onnewline, oncomposition, onhistory, oncopy, oncut, onpaste }: {
    value: string; html: string; label: string; style: string;
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
      if (!value || value.endsWith('\n')) {
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

<div class="inline-edit" bind:this={element} contenteditable="true" role="textbox" tabindex="0" aria-multiline="true" aria-label={label} {style}
  onfocus={() => { if (element) selectRichText(element, selection.start, selection.end); }}
  onbeforeinput={event => {
    capture();
    if (!composing && (event.inputType === 'insertParagraph' || event.inputType === 'insertLineBreak')) {
      event.preventDefault();
      onnewline();
    } else if (!composing && (event.inputType === 'historyUndo' || event.inputType === 'historyRedo')) {
      event.preventDefault();
      onhistory(event.inputType === 'historyUndo');
    }
  }}
  oninput={changed}
  oncompositionstart={() => { capture(); composing = true; oncomposition(true); }}
  oncompositionend={() => { composing = false; changed(); oncomposition(false); }}
  oncopy={oncopy} oncut={oncut} onpaste={onpaste}
  onpointerdown={event => event.stopPropagation()} onpointerup={event => event.stopPropagation()} ondblclick={event => event.stopPropagation()}
  onkeydown={event => {
    if (!event.isComposing && (event.ctrlKey || event.metaKey) && ['z', 'y'].includes(event.key.toLowerCase())) {
      event.preventDefault(); event.stopPropagation();
      onhistory(event.key.toLowerCase() === 'z' && !event.shiftKey);
      return;
    }
    onkeydown(event);
  }}
></div>

<style>
  .inline-edit { position: absolute; pointer-events: auto; border: 1px solid var(--ok-selected-border); background: #fff; font-family: var(--ok-font); font-size: 14px; padding: 4px; z-index: 7; white-space: pre-wrap; overflow-wrap: break-word; overflow: auto; outline: none; }

</style>

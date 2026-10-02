<script lang="ts">
  import { onMount, untrack } from 'svelte';
  import { layoutEditingTabs } from '../core/editing-tabs.ts';
  import { richTextValue, richTextSelection, selectRichText, type TextSelection } from '../core/rich-text-dom.ts';
  let { value, html, label, style, textZoom, busy = false, layout = 'canvas', onfocus, onblur, oncontextmenu, onpointerdown, oninput, onselect, onbeforeinput, onkeydown, onnewline, oncomposition, onhistory, oncopy, oncut, onpaste }: {
    value: string; html: string; label: string; style: string; textZoom: number; busy?: boolean; layout?: 'canvas' | 'outline';
    onfocus?: () => void; onblur?: () => void; oncontextmenu?: (event: MouseEvent) => void;
    onpointerdown?: (event: PointerEvent) => void;
    oninput: (value: string) => void;
    onselect: (range: TextSelection) => void;
    onbeforeinput: (range: TextSelection, event?: InputEvent) => void;
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
  let renderedZoom = 0;
  const previewFontFallback = 'Calibri, "Helvetica Neue", Arial, sans-serif';
  export function getSelection() { return element && document.activeElement === element ? richTextSelection(element) ?? selection : selection; }
  export function setSelectionRange(start: number, end: number) {
    selection = { start, end };
    if (element && document.activeElement === element) selectRichText(element, start, end);
  }
  export function focus() { element?.focus({ preventScroll: true }); }
  export function blur() { element?.blur(); }
  export function getElement() { return element; }
  export function select() { setSelectionRange(0, value.length); }
  function capture(event?: InputEvent) {
    selection = getSelection();
    onbeforeinput(selection, event);
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
    const zoom = textZoom;
    if (!element || composing) return;
    const root = element;
    untrack(() => {
      if (rendered === markup && renderedZoom === zoom) return;
      const active = document.activeElement === element;
      const range = { ...selection };
      root.innerHTML = markup;
      const wrapper = root.firstElementChild;
      if (wrapper?.tagName === 'DIV') wrapper.replaceWith(...wrapper.childNodes);
      // Clipboard sizes remain in points; only the editing view follows canvas zoom.
      for (const span of root.querySelectorAll('span')) {
        if (span.style.fontSize) span.style.fontSize = `calc(${span.style.fontSize} * var(--text-zoom))`;
        if (span.style.letterSpacing) span.style.letterSpacing = `calc(${span.style.letterSpacing} * var(--text-zoom))`;
        if (span.style.fontFamily) span.style.fontFamily += `, ${previewFontFallback}`;
      }
      if ((!value || value.endsWith('\n')) && !root.querySelector('[data-text-paragraph]')) {
        const end = document.createElement('br');
        end.setAttribute('data-caret-end', '');
        root.append(end);
      }
      layoutEditingTabs(root, zoom);
      rendered = markup;
      renderedZoom = zoom;
      if (active) setSelectionRange(Math.min(range.start, value.length), Math.min(range.end, value.length));
    });
  });
  onMount(() => {
    document.addEventListener('selectionchange', selectionChanged);
    if (layout === 'canvas') { focus(); setSelectionRange(value.length, value.length); }
    return () => { document.removeEventListener('selectionchange', selectionChanged); };
  });
</script>

<div class="inline-edit" class:outline={layout === 'outline'} bind:this={element} contenteditable="true" role="textbox" tabindex="0" aria-multiline="true" aria-busy={busy} aria-label={label} style={`${style}; --text-zoom: ${textZoom};`}
  onfocus={() => { if (element) selectRichText(element, selection.start, selection.end); onfocus?.(); }}
  {onblur} {oncontextmenu}
  onbeforeinput={event => {
    if (!composing && (event.inputType === 'historyUndo' || event.inputType === 'historyRedo')) {
      event.preventDefault();
      onhistory(event.inputType === 'historyUndo');
      return;
    }
    if (busy) { event.preventDefault(); return; }
    capture(event);
    if (!composing && (event.inputType === 'insertParagraph' || event.inputType === 'insertLineBreak')) {
      event.preventDefault();
      onnewline();
    }
  }}
  oninput={changed}
  oncompositionstart={() => { capture(); composing = true; oncomposition(true); }}
  oncompositionend={() => { composing = false; changed(); oncomposition(false); }}
  oncopy={oncopy} oncut={event => { if (busy) event.preventDefault(); else oncut(event); }} onpaste={event => { if (busy) event.preventDefault(); else onpaste(event); }}
  onpointerdown={event => { event.stopPropagation(); onpointerdown?.(event); }} onpointerup={event => event.stopPropagation()} ondblclick={event => event.stopPropagation()}
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
  .inline-edit { position: absolute; pointer-events: auto; border: 0; background: transparent; font-family: var(--ok-font); font-size: calc(14px * var(--text-zoom)); padding: calc(4px * var(--text-zoom)); z-index: 7; white-space: pre-wrap; overflow-wrap: break-word; overflow: auto; outline: 1px solid var(--ok-selected-border); }

  .inline-edit.outline { position: static; width: 100%; box-sizing: border-box; min-height: 24px; padding: 0 4px; background: transparent; color: var(--ok-text); font: 13px/24px Arial, sans-serif; outline: none; overflow: visible; }
  .outline :global([data-outline-paragraph]) { position: relative; display: inline-block; box-sizing: border-box; width: 100%; padding-left: calc(10px + var(--outline-level) * 10px); vertical-align: top; }
  .outline :global([data-outline-paragraph][data-outline-title]) { padding-left: 0; font-weight: bold; }
  .outline :global([data-outline-marker]::before) { content: attr(data-outline-marker); position: absolute; left: calc(var(--outline-level) * 10px); user-select: none; }

  .inline-edit :global([data-list-marker]::before) {
    content: attr(data-list-marker);
    font-size: var(--marker-size, inherit);
    font-family: var(--marker-font, var(--ok-font));
    color: var(--marker-color, inherit);
    margin-inline-end: 0.4em;
    user-select: none;
  }
</style>

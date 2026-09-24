import type { EditorShape } from './editor.ts';
import { textFontFamily } from './text-font.ts';

/** Measure the natural text-box size in its unrotated local coordinates. */
export function measureRichTextBox(field: RichTextField): { width: number; height: number } {
  const copy = field.cloneNode(true) as HTMLElement;
  copy.removeAttribute('id');
  copy.removeAttribute('aria-label');
  copy.contentEditable = 'false';
  copy.setAttribute('aria-hidden', 'true');
  const vertical = field.style.writingMode.startsWith('vertical');
  const wrap = field.style.whiteSpace !== 'pre';
  Object.assign(copy.style, {
    visibility: 'hidden',
    pointerEvents: 'none',
    position: 'absolute',
    left: '0',
    top: '0',
    transform: 'none',
    overflow: 'visible',
    alignContent: 'start',
    border: '0',
    width: !wrap || vertical ? 'max-content' : field.style.width,
    height: !wrap || !vertical ? 'max-content' : field.style.height,
  });
  field.parentElement!.append(copy);
  try {
    const rect = copy.getBoundingClientRect();
    return { width: rect.width, height: rect.height };
  } finally {
    copy.remove();
  }
}

/** Preserve the text anchor while resizing, including rotated text boxes. */
export function resizedTextBounds(
  shape: EditorShape,
  width: number,
  height: number,
): NonNullable<EditorShape['bounds']> {
  const old = shape.bounds!;
  const anchor = shape.textFrame?.anchor;
  const fraction = anchor === 'center' ? 0.5 : anchor === 'bottom' ? 1 : 0;
  const vertical = shape.textDirection !== 'horz';
  const inlineAnchor = shape.anchorCenter ? 0.5 : 0;
  const ax = vertical
    ? shape.textDirection === 'vert270'
      ? fraction
      : 1 - fraction
    : inlineAnchor;
  const ay = vertical ? inlineAnchor : fraction;
  const w = Math.max(1, Math.round(width)),
    h = Math.max(1, Math.round(height));
  const dx = (old.w - w) * (ax - 0.5),
    dy = (old.h - h) * (ay - 0.5);
  const angle = (shape.rotation * Math.PI) / 180;
  return {
    x: Math.round(old.x + (old.w - w) / 2 + dx * Math.cos(angle) - dy * Math.sin(angle)),
    y: Math.round(old.y + (old.h - h) / 2 + dx * Math.sin(angle) + dy * Math.cos(angle)),
    w,
    h,
  };
}

/** Fit the live editing surface without changing its text or selection. */
export function fitRichText(
  field: RichTextField,
  runs: EditorShape['runs'],
  scale: number,
  paragraphs: EditorShape['paragraphs'],
  baseSize: number,
  lnSpcReduction = 0,
): NonNullable<EditorShape['autoFitParams']> {
  const anchor = field.style.alignContent;
  const overflow = field.style.overflow;
  field.style.alignContent = 'start';
  field.style.overflow = 'hidden';
  const render = (fontScale: number) => {
    field.style.fontSize = baseSize * scale * 12700 * fontScale + 'px';
    field.renderRuns(runs, scale, paragraphs, { fontScale, lnSpcReduction });
  };
  const fits = () =>
    field.scrollHeight <= field.clientHeight && field.scrollWidth <= field.clientWidth;
  let fontScale = 1;
  try {
    render(1);
    if (!fits()) {
      let low = 0.25,
        high = 1;
      for (let i = 0; i < 9; i++) {
        const mid = (low + high) / 2;
        render(mid);
        if (fits()) low = mid;
        else high = mid;
      }
      // DrawingML persists hundred-thousandths; round down so it still fits.
      fontScale = Math.floor(low * 100000) / 100000;
      render(fontScale);
    }
  } finally {
    field.style.alignContent = anchor;
    field.style.overflow = overflow;
  }
  return { fontScale, lnSpcReduction };
}

export interface RichTextField extends HTMLDivElement {
  value: string;
  readonly selectionStart: number;
  readonly selectionEnd: number;
  select(): void;
  deleteSelection(): void;
  replaceSelection(text: string): void;
  setSelectionRange(start: number, end: number): void;
  renderRuns(
    runs: EditorShape['runs'],
    scale: number,
    paragraphs: EditorShape['paragraphs'],
    autoFit?: EditorShape['autoFitParams'],
  ): void;
  dispose(): void;
}

/** Keep paragraph boundaries and inherited properties in sync before a save. */
export function replaceParagraphRange(
  paragraphs: EditorShape['paragraphs'],
  start: number,
  end: number,
  text: string,
): EditorShape['paragraphs'] {
  if (!paragraphs.length) return [];
  const shift = text.length - (end - start);
  const breaks = paragraphs
    .slice(0, -1)
    .map((p) => p.end)
    .filter((position) => position < start || position >= end)
    .map((position) => (position < start ? position : position + shift));
  for (let i = 0; i < text.length; i++) if (text[i] === '\n') breaks.push(start + i);
  breaks.sort((a, b) => a - b);
  breaks.push(paragraphs.at(-1)!.end + shift);
  let from = 0;
  return breaks.map((to) => {
    const original = from <= start ? from : from <= start + text.length ? start : from - shift;
    const template =
      paragraphs.find((p) => p.start <= original && original <= p.end) ?? paragraphs.at(-1)!;
    const paragraph = { ...template, start: from, end: to };
    from = to + 1;
    return paragraph;
  });
}

/** A styled editing surface with offsets in the same UTF-16 space as DrawingML text. */
export function createRichTextField(): RichTextField {
  const field = document.createElement('div') as RichTextField;
  field.contentEditable = 'true';
  field.setAttribute('role', 'textbox');
  field.setAttribute('aria-multiline', 'true');
  field.spellcheck = false;
  let saved = { start: 0, end: 0, backward: false };
  const getSelection = () => {
    const root = field.getRootNode() as ShadowRoot & { getSelection?: () => Selection | null };
    return root.getSelection?.() ?? document.getSelection();
  };
  const offset = (node: Node, position: number) => {
    const range = document.createRange();
    range.setStart(field, 0);
    range.setEnd(node, position);
    return range.cloneContents().textContent?.length ?? 0;
  };
  const readSelection = () => {
    const selection = getSelection();
    if (
      selection?.anchorNode &&
      selection.focusNode &&
      field.contains(selection.anchorNode) &&
      field.contains(selection.focusNode)
    ) {
      const anchor = offset(selection.anchorNode, selection.anchorOffset);
      const focus = offset(selection.focusNode, selection.focusOffset);
      saved = {
        start: Math.min(anchor, focus),
        end: Math.max(anchor, focus),
        backward: anchor > focus,
      };
    }
    return saved;
  };
  const point = (position: number): [Node, number] => {
    const walker = document.createTreeWalker(field, NodeFilter.SHOW_TEXT);
    let node: Node | null;
    let remaining = position;
    while ((node = walker.nextNode())) {
      const length = node.textContent?.length ?? 0;
      if (remaining < length || (remaining === length && !node.textContent?.endsWith('\n')))
        return [node, remaining];
      remaining -= length;
    }
    return [field, field.childNodes.length];
  };
  const restore = () => {
    const selection = getSelection();
    if (!selection) return;
    const first = point(saved.start),
      last = point(saved.end);
    const anchor = saved.backward ? last : first,
      focus = saved.backward ? first : last;
    selection.setBaseAndExtent(anchor[0], anchor[1], focus[0], focus[1]);
  };
  Object.defineProperties(field, {
    value: {
      get: () => field.textContent ?? '',
      set: (value: string) => {
        field.textContent = value;
      },
    },
    selectionStart: { get: () => readSelection().start },
    selectionEnd: { get: () => readSelection().end },
  });
  field.setSelectionRange = (start, end) => {
    saved = { start, end, backward: false };
    restore();
  };
  field.select = () => field.setSelectionRange(0, field.value.length);
  field.renderRuns = (runs, scale, paragraphs, autoFit) => {
    const ownsSelection = field.contains(getSelection()?.anchorNode ?? null);
    scale *= autoFit?.fontScale ?? 1;
    const lineScale = 1 - (autoFit?.lnSpcReduction ?? 0);
    readSelection();
    const value = field.value;
    const fragment = document.createDocumentFragment();
    const counters = new Map<number, number>();
    for (const paragraph of paragraphs) {
      const block = document.createElement('div');
      block.className = 'edit-paragraph';
      const props = paragraph.properties;
      const bullet = props.bullet;
      const numbered =
        bullet === 'number' || (typeof bullet === 'object' && bullet && 'autoNum' in bullet);
      for (const level of counters.keys()) if (level > props.level) counters.delete(level);
      const number = (counters.get(props.level) ?? 0) + 1;
      if (numbered) counters.set(props.level, number);
      else counters.delete(props.level);
      const marker = numbered
        ? number + '.'
        : bullet && typeof bullet === 'object' && 'char' in bullet
          ? bullet.char
          : bullet === 'bullet' || (bullet !== 'none' && props.level > 0)
            ? props.level === 0
              ? '•'
              : props.level === 1
                ? '◦'
                : '▪'
            : '';
      if (marker) block.dataset.bullet = marker;
      const left = props.marL ?? (marker ? 342900 : props.level * 304800);
      const firstRun = runs.find((run) => run.start <= paragraph.end && run.end > paragraph.start);
      Object.assign(block.style, {
        fontSize: firstRun?.format.size ? firstRun.format.size * 12700 * scale + 'px' : '',
        margin: '0',
        paddingInlineStart: left * scale + 'px',
        paddingInlineEnd: (props.marR ?? 0) * scale + 'px',
        textIndent: (props.indent ?? (marker ? -left : 0)) * scale + 'px',
        textAlign:
          props.align === 'distribute' ? 'justify' : (props.align ?? field.style.textAlign),
        textAlignLast: props.align === 'distribute' ? 'justify' : 'auto',
        lineHeight:
          props.lineSpacing?.kind === 'pts'
            ? props.lineSpacing.value * 12700 * scale * lineScale + 'px'
            : String((props.lineSpacing?.value ?? 1.05) * lineScale),
        marginBlockStart: (props.spcBefPts ?? 0) * 12700 * scale + 'px',
        marginBlockEnd: (props.spcAftPts ?? 0) * 12700 * scale + 'px',
        direction: props.rtl ? 'rtl' : 'ltr',
      });
      let offset = paragraph.start;
      for (const run of runs) {
        const start = Math.max(run.start, paragraph.start),
          end = Math.min(run.end, paragraph.end);
        if (start >= end) continue;
        if (start > offset) block.append(document.createTextNode(value.slice(offset, start)));
        const span = document.createElement('span');
        span.textContent = value.slice(start, end);
        const format = run.format;
        const decoration = [];
        if (format.underline && format.underline !== 'none') decoration.push('underline');
        if (format.strike && format.strike !== 'noStrike') decoration.push('line-through');
        Object.assign(span.style, {
          fontFamily: textFontFamily(format.font),
          fontSize: (format.size ?? 18) * 12700 * scale * (format.baseline ? 0.65 : 1) + 'px',
          fontWeight: format.bold ? 'bold' : 'normal',
          fontStyle: format.italic ? 'italic' : 'normal',
          fontKerning:
            format.kern === undefined
              ? 'auto'
              : format.kern > 0 &&
                  (format.size ?? 18) * (autoFit?.fontScale ?? 1) * 100 >= format.kern
                ? 'normal'
                : 'none',
          color: format.color ?? '#000000',
          backgroundColor: format.highlight ?? 'transparent',
          textDecoration: decoration.join(' ') || 'none',
          letterSpacing: ((format.spc ?? 0) / 100) * 12700 * scale + 'px',
          verticalAlign: (format.baseline ?? 0) * (format.size ?? 18) * 12700 * scale + 'px',
        });
        block.append(span);
        offset = end;
      }
      if (offset < paragraph.end)
        block.append(document.createTextNode(value.slice(offset, paragraph.end)));
      if (paragraph.end < value.length) block.append(document.createTextNode('\n'));
      if (paragraph.start === paragraph.end && paragraph.end === value.length)
        block.append(document.createTextNode(''), document.createElement('br'));
      fragment.append(block);
    }
    field.replaceChildren(fragment);
    if (ownsSelection) restore();
  };
  const onSelection = () => {
    if (field.isConnected) {
      readSelection();
      field.dispatchEvent(new Event('select'));
    }
  };
  document.addEventListener('selectionchange', onSelection);
  field.dispose = () => document.removeEventListener('selectionchange', onSelection);
  const insert = (
    text: string,
    inputType: string,
    replacement?: { start: number; end: number },
  ) => {
    const { start, end } = replacement ?? readSelection();
    field.dispatchEvent(new InputEvent('beforeinput', { inputType, data: text, bubbles: true }));
    const range = document.createRange();
    const first = point(start),
      last = point(end);
    range.setStart(first[0], first[1]);
    range.setEnd(last[0], last[1]);
    range.deleteContents();
    range.insertNode(document.createTextNode(text));
    field.setSelectionRange(start + text.length, start + text.length);
    field.dispatchEvent(new InputEvent('input', { inputType, data: text, bubbles: true }));
  };
  field.replaceSelection = (text) => {
    const { start, end } = readSelection();
    field.focus();
    field.setSelectionRange(start, end);
    insert(text.replace(/\r\n?/g, '\n'), 'insertReplacementText');
  };
  field.deleteSelection = () => {
    const { start, end } = readSelection();
    if (start === end) return;
    field.focus();
    field.setSelectionRange(start, end);
    insert('', 'deleteContentBackward');
  };
  field.addEventListener('beforeinput', (event) => {
    if (event.isTrusted && event.inputType.startsWith('delete')) {
      let { start, end } = readSelection();
      if (start === end) {
        if (event.inputType === 'deleteContentBackward') start = Math.max(0, start - 1);
        else if (event.inputType === 'deleteContentForward')
          end = Math.min(field.value.length, end + 1);
        else {
          const target = event.getTargetRanges()[0];
          if (
            target &&
            field.contains(target.startContainer) &&
            field.contains(target.endContainer)
          ) {
            start = offset(target.startContainer, target.startOffset);
            end = offset(target.endContainer, target.endOffset);
          }
        }
      }
      if (field.value.slice(start, end).includes('\n')) {
        event.preventDefault();
        insert('', event.inputType, { start, end });
        return;
      }
    }
    const newLine = ['insertParagraph', 'insertLineBreak'].includes(event.inputType);
    const multilineText = event.inputType === 'insertText' && /[\r\n]/.test(event.data ?? '');
    if (event.isTrusted && (newLine || multilineText)) {
      event.preventDefault();
      insert(newLine ? '\n' : event.data!.replace(/\r\n?/g, '\n'), event.inputType);
    }
  });
  field.addEventListener('paste', (event) => {
    event.preventDefault();
    insert(
      event.clipboardData?.getData('text/plain').replace(/\r\n?/g, '\n') ?? '',
      'insertFromPaste',
    );
  });
  return field;
}

/** Moves the text bounds as a unit without changing paragraph alignment or wrapping. */
export function centerRichText(field: RichTextField, centered: boolean): void {
  for (const child of field.children) (child as HTMLElement).style.translate = '';
  if (!centered || !field.parentElement || !field.textContent?.trim()) return;
  const copy = field.cloneNode(true) as HTMLElement;
  copy.removeAttribute('id');
  copy.contentEditable = 'false';
  copy.setAttribute('aria-hidden', 'true');
  Object.assign(copy.style, {
    visibility: 'hidden',
    pointerEvents: 'none',
    transform: 'none',
    position: 'absolute',
    left: '0',
    top: '0',
  });
  field.parentElement.append(copy);
  try {
    for (const block of copy.querySelectorAll<HTMLElement>('[data-bullet]')) {
      const markerStyle = getComputedStyle(block, '::before');
      const marker = document.createElement('span');
      marker.textContent = block.dataset.bullet!;
      Object.assign(marker.style, { font: markerStyle.font, marginRight: markerStyle.marginRight });
      delete block.dataset.bullet;
      block.prepend(marker);
    }
    const vertical = field.style.writingMode.startsWith('vertical');
    const style = getComputedStyle(copy);
    const box = copy.getBoundingClientRect();
    const start = vertical
      ? box.top + parseFloat(style.borderTopWidth || '0') + parseFloat(style.paddingTop || '0')
      : box.left + parseFloat(style.borderLeftWidth || '0') + parseFloat(style.paddingLeft || '0');
    const end = vertical
      ? box.bottom -
        parseFloat(style.borderBottomWidth || '0') -
        parseFloat(style.paddingBottom || '0')
      : box.right -
        parseFloat(style.borderRightWidth || '0') -
        parseFloat(style.paddingRight || '0');
    let low = Infinity,
      high = -Infinity;
    const walker = document.createTreeWalker(copy, NodeFilter.SHOW_TEXT);
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      const length = node.textContent?.trimEnd().length ?? 0;
      if (!length) continue;
      const range = document.createRange();
      range.setStart(node, 0);
      range.setEnd(node, length);
      for (const rect of range.getClientRects()) {
        low = Math.min(low, vertical ? rect.top : rect.left);
        high = Math.max(high, vertical ? rect.bottom : rect.right);
      }
    }
    if (!Number.isFinite(low)) return;
    const shift = (start + end - low - high) / 2;
    for (const child of field.children)
      (child as HTMLElement).style.translate = vertical ? `0 ${shift}px` : `${shift}px 0`;
  } finally {
    copy.remove();
  }
}

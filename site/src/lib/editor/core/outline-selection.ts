import type { TextEdit } from './text-edit-preview.ts';
import type { TextFormat } from '@office-kit/pptx';
import { richTextPoint } from './rich-text-dom.ts';

export type OutlinePoint = { key: string; offset: number };
export type OutlineRange = { start: OutlinePoint; end: OutlinePoint };
export type OutlineClipboard = {
  text: string;
  formats: Array<{ start: number; end: number; format: TextFormat }>;
};

export type OutlineSelectionField = {
  key: string;
  root: HTMLElement;
  text: () => string;
  copy: (start: number, end: number) => OutlineClipboard;
  flush: () => readonly TextEdit[];
  apply: (edits: readonly TextEdit[]) => void;
  transact: (label: string, fn: () => void) => void;
  focus: (offset: number) => void;
  setRange: (offset: number) => void;
};

function before(a: OutlinePoint, b: OutlinePoint, fields: OutlineSelectionField[]): boolean {
  const ai = fields.findIndex((field) => field.key === a.key);
  const bi = fields.findIndex((field) => field.key === b.key);
  return ai < bi || (ai === bi && a.offset <= b.offset);
}

/** Shared ordered selection state for the independent contenteditables in Outline View. */
export class OutlineSelectionModel {
  #fields = new Map<string, OutlineSelectionField>();
  #anchor: OutlinePoint | null = null;
  #focus: OutlinePoint | null = null;
  #suppress = false;
  #preserveAnchor = false;
  #transitioning = false;
  #caretGeneration = 0;

  register(field: OutlineSelectionField): () => void {
    const previous = this.#fields.get(field.key);
    this.#fields.set(field.key, field);
    if (previous && previous.root !== field.root) this.clear();
    return () => {
      if (this.#fields.get(field.key) !== field) return;
      this.#fields.delete(field.key);
      if (this.#anchor?.key === field.key || this.#focus?.key === field.key) this.clear();
    };
  }

  fields(): OutlineSelectionField[] {
    return [...this.#fields.values()].sort((a, b) =>
      a.root.compareDocumentPosition(b.root) & 4 ? -1 : 1,
    );
  }

  current(): OutlineRange | null {
    if (!this.#anchor || !this.#focus) return null;
    const fields = this.fields();
    if (
      !fields.some((field) => field.key === this.#anchor!.key) ||
      !fields.some((field) => field.key === this.#focus!.key)
    )
      return null;
    return before(this.#anchor, this.#focus, fields)
      ? { start: this.#anchor, end: this.#focus }
      : { start: this.#focus, end: this.#anchor };
  }

  clear(): void {
    this.#caretGeneration++;
    this.#anchor = null;
    this.#focus = null;
    this.#preserveAnchor = false;
  }

  update(field: OutlineSelectionField, start: number, end: number, extend = false): void {
    if (this.#suppress) return;
    // Chromium reports a cross-contenteditable selection on the anchor root
    // after focus moves to another root. Preserve the logical range until a
    // subsequent keyboard or pointer action explicitly establishes a caret.
    if (
      this.#anchor &&
      this.#focus &&
      this.#anchor.key !== this.#focus.key &&
      field.key === this.#anchor.key
    )
      return;
    const point = { key: field.key, offset: end };
    if (this.#preserveAnchor) {
      this.#focus = point;
      this.#preserveAnchor = false;
      return;
    }
    if (!extend || !this.#anchor) this.#anchor = { key: field.key, offset: start };
    this.#focus = point;
  }

  setCaret(field: OutlineSelectionField, offset: number): void {
    this.#anchor = { key: field.key, offset };
    this.#focus = { key: field.key, offset };
  }

  transitioning(): boolean {
    return this.#transitioning;
  }

  /** Collapse a cross-field range to its logical start or end. */
  collapse(direction: -1 | 1): boolean {
    const range = this.current();
    if (!range) return false;
    const point = direction < 0 ? range.start : range.end;
    const field = this.#fields.get(point.key);
    if (!field) return false;
    this.#suppress = true;
    try {
      field.focus(point.offset);
      field.setRange(point.offset);
    } finally {
      this.#suppress = false;
    }
    this.#caretGeneration++;
    this.#anchor = { ...point };
    this.#focus = { ...point };
    this.#preserveAnchor = false;
    return true;
  }

  /** The native selection focus offset for a registered field, if it owns focus. */
  focusOffset(field: OutlineSelectionField): number | null {
    return this.#focus?.key === field.key ? this.#focus.offset : null;
  }

  focusedField(): OutlineSelectionField | null {
    return this.#focus ? (this.#fields.get(this.#focus.key) ?? null) : null;
  }

  extend(field: OutlineSelectionField, direction: -1 | 1): boolean {
    const fields = this.fields();
    const index = fields.indexOf(field);
    if (index < 0) return false;
    const current =
      this.#focus?.key === field.key ? this.#focus.offset : direction < 0 ? 0 : field.text().length;
    const atBoundary = direction < 0 ? current === 0 : current === field.text().length;
    const bridge =
      this.#anchor &&
      this.#anchor.key !== field.key &&
      ((direction > 0 && current === 0) || (direction < 0 && current === field.text().length));
    if (!atBoundary && !bridge) return false;
    if (
      this.#anchor &&
      ((direction > 0 && current === 0 && this.#anchor.key !== field.key) ||
        (direction < 0 && current === field.text().length && this.#anchor.key !== field.key))
    ) {
      this.#focus = { key: field.key, offset: direction > 0 ? field.text().length : 0 };
      this.#preserveAnchor = true;
      this.#selectNative();
      return true;
    }
    const next = fields[index + direction];
    if (!next) return false;
    if (!this.#anchor) this.#anchor = { key: field.key, offset: current };
    this.#focus = { key: next.key, offset: direction < 0 ? next.text().length : 0 };
    this.#preserveAnchor = true;
    this.#suppress = true;
    this.#transitioning = true;
    try {
      next.focus(this.#focus.offset);
      this.#selectNative();
    } finally {
      this.#transitioning = false;
      this.#suppress = false;
    }
    return true;
  }

  #selectNative(): void {
    if (!this.#anchor || !this.#focus) return;
    const fields = this.fields();
    const start = fields.find((field) => field.key === this.#anchor!.key);
    const end = fields.find((field) => field.key === this.#focus!.key);
    if (!start || !end || !start.root.ownerDocument) return;
    const a = richTextPoint(start.root, this.#anchor.offset);
    const b = richTextPoint(end.root, this.#focus.offset);
    const selection = start.root.ownerDocument.getSelection();
    selection?.removeAllRanges();
    selection?.setBaseAndExtent(a.node, a.offset, b.node, b.offset);
  }

  copy(): OutlineClipboard | null {
    const range = this.current();
    if (!range) return null;
    const fields = this.fields();
    const from = fields.findIndex((field) => field.key === range.start.key);
    const to = fields.findIndex((field) => field.key === range.end.key);
    if (from < 0 || to < from) return null;
    let text = '';
    const formats: OutlineClipboard['formats'] = [];
    for (let index = from; index <= to; index++) {
      const field = fields[index]!;
      const start = index === from ? range.start.offset : 0;
      const end = index === to ? range.end.offset : field.text().length;
      const copied = field.copy(start, end);
      if (index > from) text += '\n';
      const shift = text.length;
      text += copied.text;
      formats.push(
        ...copied.formats.map((format) => ({
          ...format,
          start: shift + format.start,
          end: shift + format.end,
        })),
      );
    }
    return { text, formats };
  }

  replace(text: string, formats: OutlineClipboard['formats'] = [], label = 'Edit text'): boolean {
    const range = this.current();
    if (!range) return false;
    const fields = this.fields();
    const from = fields.findIndex((field) => field.key === range.start.key);
    const to = fields.findIndex((field) => field.key === range.end.key);
    if (from < 0 || to < from) return false;
    const edits = fields.slice(from, to + 1).map((field, index) => ({
      field,
      start: index === 0 ? range.start.offset : 0,
      end: index === to - from ? range.end.offset : field.text().length,
    }));
    const first = edits[0]!;
    const destination = { start: first.start, end: first.start + text.length };
    // Each field owns its pending edits; flush them into the same document transaction.
    first.field.transact(label, () => {
      for (const edit of edits) edit.field.apply(edit.field.flush());
      for (const edit of edits.slice().reverse()) {
        const replacement =
          edit === first
            ? { start: edit.start, end: edit.end, text, formats }
            : { start: edit.start, end: edit.end, text: '', formats: [] };
        edit.field.apply([replacement]);
      }
    });
    this.#anchor = { key: first.field.key, offset: destination.end };
    this.#focus = { key: first.field.key, offset: destination.end };
    this.#preserveAnchor = false;
    // Keep the native caret in the replacement field.  Without this, the
    // next keystroke can still target the field that owned the old focus,
    // even though the logical range has already collapsed.
    // Svelte applies the transaction-driven DOM update on the next turn; a
    // synchronous focus targets the old editor node and is lost during that
    // update. Restore the caret after the replacement has rendered.
    const view = first.field.root.ownerDocument?.defaultView;
    const generation = ++this.#caretGeneration;
    const restoreCaret = () => {
      if (
        generation !== this.#caretGeneration ||
        !this.#fields.has(first.field.key) ||
        !first.field.root.isConnected
      )
        return;
      first.field.setRange(destination.end);
      first.field.focus(destination.end);
    };
    if (view?.requestAnimationFrame) view.requestAnimationFrame(restoreCaret);
    else setTimeout(restoreCaret, 0);
    return true;
  }
}

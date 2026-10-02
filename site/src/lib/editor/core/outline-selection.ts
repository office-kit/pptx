import type { TextEdit } from './text-edit-preview.ts';
import type { ParagraphProperties, SlideShapeData, TextCase, TextFormat } from '@office-kit/pptx';
import { richTextPoint } from './rich-text-dom.ts';
import { textCaseRange } from './text-case.ts';

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
  formats: (start: number, end: number) => TextFormat[];
  applyFormat: (start: number, end: number, format: TextFormat, reset: boolean) => void;
  applyFontSize?: (start: number, end: number, direction: 1 | -1) => void;
  changeCase: (start: number, end: number, value: TextCase, caret?: number) => number;
  paragraphs: (start: number, end: number) => ParagraphProperties[];
  editParagraphs: (
    start: number,
    end: number,
    edit: (shape: SlideShapeData, index: number) => void,
  ) => void;
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
  #listeners = new Set<() => void>();

  subscribe(listener: () => void): () => void {
    this.#listeners.add(listener);
    return () => this.#listeners.delete(listener);
  }

  #changed(): void {
    for (const listener of this.#listeners) listener();
  }

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
    this.#changed();
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
      this.#changed();
      return;
    }
    if (!extend || !this.#anchor) this.#anchor = { key: field.key, offset: start };
    this.#focus = point;
    this.#changed();
  }

  setCaret(field: OutlineSelectionField, offset: number): void {
    this.#anchor = { key: field.key, offset };
    this.#focus = { key: field.key, offset };
    this.#changed();
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
    this.#changed();
    return true;
  }

  /** The native selection focus offset for a registered field, if it owns focus. */
  focusOffset(field: OutlineSelectionField): number | null {
    return this.#focus?.key === field.key ? this.#focus.offset : null;
  }

  focusedField(): OutlineSelectionField | null {
    return this.#focus ? (this.#fields.get(this.#focus.key) ?? null) : null;
  }

  /** Formats intersecting the current ordered range, preserving field order. */
  formats(): TextFormat[] {
    const range = this.current();
    if (!range) return [];
    const fields = this.fields();
    const from = fields.findIndex((field) => field.key === range.start.key);
    const to = fields.findIndex((field) => field.key === range.end.key);
    if (from < 0 || to < from) return [];
    const result: TextFormat[] = [];
    for (let index = from; index <= to; index++) {
      const field = fields[index]!;
      const start = index === from ? range.start.offset : 0;
      const end = index === to ? range.end.offset : field.text().length;
      result.push(...field.formats(start, end));
    }
    return result;
  }

  paragraphs(): ParagraphProperties[] {
    const range = this.current();
    if (!range) return [];
    const fields = this.fields();
    const from = fields.findIndex((field) => field.key === range.start.key);
    const to = fields.findIndex((field) => field.key === range.end.key);
    if (from < 0 || to < from) return [];
    const result: ParagraphProperties[] = [];
    for (let index = from; index <= to; index++) {
      const field = fields[index]!;
      result.push(
        ...field.paragraphs(
          index === from ? range.start.offset : 0,
          index === to ? range.end.offset : field.text().length,
        ),
      );
    }
    return result;
  }

  editParagraphs(
    edit: (shape: SlideShapeData, index: number) => void,
    label = 'Format paragraphs',
  ): boolean {
    const range = this.current();
    if (!range) return false;
    const fields = this.fields();
    const from = fields.findIndex((field) => field.key === range.start.key);
    const to = fields.findIndex((field) => field.key === range.end.key);
    if (from < 0 || to < from) return false;
    const selected = fields.slice(from, to + 1).map((field, index) => ({
      field,
      start: index === 0 ? range.start.offset : 0,
      end: index === to - from ? range.end.offset : field.text().length,
    }));
    const pending = selected.map((item) => ({ item, changes: item.field.flush() }));
    selected[0]!.field.transact(label, () => {
      for (const { item, changes } of pending) item.field.apply(changes);
      for (const item of selected) item.field.editParagraphs(item.start, item.end, edit);
    });
    this.#changed();
    return true;
  }

  /** Apply one character format transaction to every field touched by the range. */
  format(
    format: TextFormat | ((formats: TextFormat[]) => TextFormat),
    reset = false,
    label = 'Format selected text',
  ): boolean {
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
    // Drafts must be part of the same history entry as formatting. This also
    // ensures the format reader sees the committed shape runs below.
    const pending = edits.map((edit) => ({ edit, changes: edit.field.flush() }));
    edits[0]!.field.transact(label, () => {
      for (const { edit, changes } of pending) edit.field.apply(changes);
      const resolved = typeof format === 'function' ? format(this.formats()) : format;
      for (const edit of edits) edit.field.applyFormat(edit.start, edit.end, resolved, reset);
    });
    this.#changed();
    return true;
  }

  /** Apply PowerPoint's relative font-size step as one history entry. */
  fontSize(
    direction: 1 | -1,
    label = direction > 0 ? 'Increase Font Size' : 'Decrease Font Size',
  ): boolean {
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
    if (!edits.every((edit) => edit.field.applyFontSize)) return false;
    const pending = edits.map((edit) => ({ edit, changes: edit.field.flush() }));
    edits[0]!.field.transact(label, () => {
      for (const { edit, changes } of pending) edit.field.apply(changes);
      for (const edit of edits) edit.field.applyFontSize!(edit.start, edit.end, direction);
    });
    this.#changed();
    return true;
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
      this.#changed();
      this.#selectNative();
      return true;
    }
    const next = fields[index + direction];
    if (!next) return false;
    if (!this.#anchor) this.#anchor = { key: field.key, offset: current };
    this.#focus = { key: next.key, offset: direction < 0 ? next.text().length : 0 };
    this.#preserveAnchor = true;
    this.#changed();
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
    this.#changed();
    // Keep the native caret in the replacement field.  Without this, the
    // next keystroke can still target the field that owned the old focus,
    // even though the logical range has already collapsed.
    // Svelte applies the transaction-driven DOM update on the next turn; a
    // synchronous focus targets the old editor node and is lost during that
    // update. Restore the caret after the replacement has rendered.
    const ownerDocument = first.field.root.ownerDocument;
    const activeElementAtSchedule = ownerDocument?.activeElement;
    const view = ownerDocument?.defaultView;
    const generation = ++this.#caretGeneration;
    const restoreCaret = () => {
      if (
        generation !== this.#caretGeneration ||
        !this.#fields.has(first.field.key) ||
        !first.field.root.isConnected
      )
        return;
      // A ribbon command can take focus between the replacement transaction
      // and this frame. Do not steal that focus back after the user has
      // deliberately moved out of the outline editor.
      const activeElement = ownerDocument?.activeElement;
      if (activeElement !== activeElementAtSchedule) return;
      first.field.setRange(destination.end);
      first.field.focus(destination.end);
    };
    if (view?.requestAnimationFrame) view.requestAnimationFrame(restoreCaret);
    else setTimeout(restoreCaret, 0);
    return true;
  }

  changeCase(value: TextCase, label = 'Change Case'): boolean {
    const current = this.current();
    if (!current) return false;
    const fields = this.fields();
    const from = fields.findIndex((field) => field.key === current.start.key);
    const to = fields.findIndex((field) => field.key === current.end.key);
    if (from < 0 || to < from) return false;
    const edits = fields.slice(from, to + 1).map((field, index) => {
      let start = index === 0 ? current.start.offset : 0;
      let end = index === to - from ? current.end.offset : field.text().length;
      const caret =
        from === to && current.start.offset === current.end.offset
          ? current.start.offset
          : undefined;
      if (start === end && from === to) {
        const expanded = textCaseRange(field.text(), { start, end });
        start = expanded.start;
        end = expanded.end;
      }
      return { field, start, end, caret };
    });
    const first = edits[0]!;
    const last = edits.at(-1)!;
    const collapsed = current.start.offset === current.end.offset && from === to;
    let destination = first.start;
    let finalEnd = last.end;
    first.field.transact(label, () => {
      for (const edit of edits) edit.field.apply(edit.field.flush());
      for (const edit of edits) {
        const nextEnd =
          edit.caret === undefined
            ? edit.field.changeCase(edit.start, edit.end, value)
            : edit.field.changeCase(edit.start, edit.end, value, edit.caret);
        if (edit === first) destination = nextEnd;
        if (edit === last) finalEnd = nextEnd;
      }
    });
    this.#anchor = collapsed
      ? { key: first.field.key, offset: destination }
      : { key: first.field.key, offset: first.start };
    this.#focus = collapsed
      ? { key: first.field.key, offset: destination }
      : { key: last.field.key, offset: finalEnd };
    this.#preserveAnchor = false;
    this.#changed();
    const ownerDocument = first.field.root.ownerDocument;
    const activeElementAtSchedule = ownerDocument?.activeElement;
    const view = ownerDocument?.defaultView;
    const generation = ++this.#caretGeneration;
    const restoreCaret = () => {
      if (
        generation !== this.#caretGeneration ||
        !this.#fields.has(first.field.key) ||
        !first.field.root.isConnected
      )
        return;
      if (ownerDocument?.activeElement !== activeElementAtSchedule) return;
      if (collapsed) {
        first.field.setRange(destination);
        first.field.focus(destination);
      } else {
        first.field.focus(first.start);
        this.#selectNative();
      }
    };
    if (view?.requestAnimationFrame) view.requestAnimationFrame(restoreCaret);
    else setTimeout(restoreCaret, 0);
    return true;
  }
}

import type { TextEdit } from './text-edit-preview.ts';

/** Find the changed text without moving repeated characters across the caret. */
export function textEditDiff(
  before: string,
  value: string,
  selection: { start: number; end: number },
  caretAfter: number,
): TextEdit | null {
  if (before === value) return null;
  let start = 0;
  const limit = Math.min(selection.start, caretAfter);
  while (
    start < limit &&
    start < before.length &&
    start < value.length &&
    before[start] === value[start]
  )
    start++;
  const splitsPair = (text: string, at: number) =>
    at > 0 &&
    at < text.length &&
    /[\uD800-\uDBFF]/.test(text[at - 1]!) &&
    /[\uDC00-\uDFFF]/.test(text[at]!);
  if (splitsPair(before, start) || splitsPair(value, start)) start--;
  let end = before.length;
  let newEnd = value.length;
  while (
    end > Math.max(start, selection.end) &&
    newEnd > start &&
    before[end - 1] === value[newEnd - 1]
  ) {
    end--;
    newEnd--;
  }
  if (splitsPair(before, end) || splitsPair(value, newEnd)) {
    end++;
    newEnd++;
  }
  return { start, end, text: value.slice(start, newEnd) };
}

export type TextRange = { start: number; end: number };

/** The reference desktop app applies Change Case to the current word when only a caret is present. */
export function textCaseRange(text: string, range: TextRange): TextRange {
  if (range.start !== range.end) return range;
  for (const part of new Intl.Segmenter(undefined, { granularity: 'word' }).segment(text)) {
    if (
      part.isWordLike &&
      part.index <= range.start &&
      range.start <= part.index + part.segment.length
    ) {
      return { start: part.index, end: part.index + part.segment.length };
    }
  }
  return range;
}

/** Keep a caret at its original character boundary even when casing expands a character. */
export function textCaseSelection(
  before: string,
  after: string,
  selection: TextRange,
  target: TextRange,
): TextRange {
  const delta = after.length - before.length;
  if (selection.start !== selection.end)
    return { start: selection.start, end: selection.end + delta };
  let outputOffset = target.start;
  for (const character of before.slice(target.start, selection.start)) {
    const upper = character.toUpperCase();
    const lower = character.toLowerCase();
    outputOffset += after.startsWith(upper, outputOffset)
      ? upper.length
      : after.startsWith(lower, outputOffset)
        ? lower.length
        : character.length;
  }
  return { start: outputOffset, end: outputOffset };
}

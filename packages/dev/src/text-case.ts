export type TextCase = 'sentence' | 'lower' | 'upper' | 'title' | 'toggle';

/** Reverse-ordered replacements retain each source character's run formatting. */
export function textCaseEdits(
  text: string,
  mode: TextCase,
  range = { start: 0, end: text.length },
) {
  if (!['sentence', 'lower', 'upper', 'title', 'toggle'].includes(mode))
    throw new Error('Invalid text case.');
  const { start, end } = range;
  if (
    !Number.isInteger(start) ||
    !Number.isInteger(end) ||
    start < 0 ||
    end < start ||
    end > text.length
  )
    throw new Error('Invalid text range.');
  const source = text.slice(start, end);
  // Convert the whole selection so contextual casing (such as final sigma) is retained.
  const lower = source.toLowerCase();
  const capitals = new Set<number>();
  if (mode === 'sentence' || mode === 'title') {
    const segments = new Intl.Segmenter(undefined, {
      granularity: mode === 'sentence' ? 'sentence' : 'word',
    }).segment(source);
    for (const segment of segments) {
      if (mode === 'title' && !segment.isWordLike) continue;
      const first = segment.segment.search(/\p{Cased_Letter}/u);
      if (first !== -1) capitals.add(segment.index + first);
    }
  }
  const edits: Array<{ start: number; end: number; text: string }> = [];
  let position = start;
  let output = 0;
  for (const character of source) {
    const upper = character.toUpperCase();
    const length = character.toLowerCase().length;
    const makeUpper =
      mode === 'upper' ||
      capitals.has(position - start) ||
      (mode === 'toggle' && character !== upper);
    const replacement = makeUpper ? upper : lower.slice(output, output + length);
    if (replacement !== character)
      edits.push({ start: position, end: position + character.length, text: replacement });
    position += character.length;
    output += length;
  }
  return edits.reverse();
}

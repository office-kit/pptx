import type { EditorShape } from './editor.ts';

export interface TextMatch {
  slide: number;
  shape: number;
  start: number;
  end: number;
}

export interface TextSearchOptions {
  matchCase?: boolean;
  wholeWords?: boolean;
}

/** Literal matches with offsets in the editor's UTF-16 text. */
export function findText(
  slides: Array<{ shapes: Array<Pick<EditorShape, 'id' | 'text' | 'textable' | 'bounds'>> }>,
  query: string,
  options: TextSearchOptions = {},
): TextMatch[] {
  if (!query) return [];
  const expression = new RegExp(
    query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'),
    options.matchCase ? 'gu' : 'giu',
  );
  const matches: TextMatch[] = [];
  for (const [slide, model] of slides.entries()) {
    for (const shape of model.shapes) {
      if (!shape.textable || !shape.bounds) continue;
      for (const match of shape.text.matchAll(expression)) {
        if (options.wholeWords) {
          const before =
            Array.from(shape.text.slice(Math.max(0, match.index - 2), match.index)).at(-1) ?? '';
          const after =
            Array.from(
              shape.text.slice(match.index + match[0].length, match.index + match[0].length + 2),
            )[0] ?? '';
          if (/[\p{L}\p{N}\p{M}\p{Pc}]/u.test(before) || /[\p{L}\p{N}\p{M}\p{Pc}]/u.test(after))
            continue;
        }
        matches.push({
          slide,
          shape: shape.id,
          start: match.index,
          end: match.index + match[0].length,
        });
      }
    }
  }
  return matches;
}

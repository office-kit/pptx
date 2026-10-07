// Mac PowerPoint sets a large-button caption of more than one word on two
// lines, broken where the wider line is narrowest ("New / Slide",
// "Insert Row / Above", "Insert / Column Right"); a tie breaks after the later
// word. Width is the rendered width at the ribbon's 11 pt, not the character
// count: "Insert Row" has more characters than "Row Above" but is narrower.
// A line never starts with "&" ("Header & / Footer", "Date & / Time").
// Japanese captions have no spaces, so they break after a particle
// (の, と, を, に, へ) or where the script changes ("新しい / スライド",
// "インクを / テキストに変換", "最前面へ / 移動"); a caption with neither
// ("スクリーンショット") stays on one line. Rendered with
// `white-space: pre-line`.
const MIN_UNSPACED_BREAK = 6;
const PARTICLES = new Set(['の', 'と', 'を', 'に', 'へ']);
const FONT = '11px -apple-system, BlinkMacSystemFont, "Segoe UI", system-ui, sans-serif';

type Script = 'hiragana' | 'katakana' | 'han' | 'other';
const scriptOf = (char: string): Script =>
  /\p{Script=Hiragana}/u.test(char)
    ? 'hiragana'
    : /[\p{Script=Katakana}ー]/u.test(char)
      ? 'katakana'
      : /\p{Script=Han}/u.test(char)
        ? 'han'
        : 'other';

let context: CanvasRenderingContext2D | null | undefined;
// Without a DOM (unit tests, server rendering) the character count stands in.
function textWidth(text: string): number {
  if (context === undefined)
    context =
      typeof document === 'undefined' ? null : document.createElement('canvas').getContext('2d');
  if (!context) return text.length;
  context.font = FONT;
  return context.measureText(text).width;
}

function breakPoints(label: string): number[] {
  if (label.includes(' ')) {
    const points: number[] = [];
    for (let index = 0; index < label.length; index += 1)
      if (label[index] === ' ' && label[index + 1] !== '&') points.push(index);
    return points;
  }
  if (label.length < MIN_UNSPACED_BREAK) return [];
  const chars = [...label];
  const points: number[] = [];
  for (let index = 1; index < chars.length; index += 1) {
    const before = chars[index - 1]!;
    const after = chars[index]!;
    const scriptChange =
      scriptOf(before) !== scriptOf(after) &&
      scriptOf(before) !== 'other' &&
      scriptOf(after) !== 'other';
    if (PARTICLES.has(before) || (scriptChange && !PARTICLES.has(after))) points.push(index);
  }
  return points;
}

export function captionLines(label: string, measure = textWidth): string {
  const points = breakPoints(label);
  if (!points.length) return label;
  const spaced = label.includes(' ');
  const chars = [...label];
  let split = -1;
  let longest = Number.POSITIVE_INFINITY;
  for (const point of points) {
    const first = spaced ? label.slice(0, point) : chars.slice(0, point).join('');
    const second = spaced ? label.slice(point + 1) : chars.slice(point).join('');
    const width = Math.max(measure(first), measure(second));
    if (width <= longest) {
      split = point;
      longest = width;
    }
  }
  return spaced
    ? `${label.slice(0, split)}\n${label.slice(split + 1)}`
    : `${chars.slice(0, split).join('')}\n${chars.slice(split).join('')}`;
}

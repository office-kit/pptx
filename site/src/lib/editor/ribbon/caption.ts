// Mac PowerPoint sets a large-button caption of more than one word on two
// lines, broken where the lines balance ("New / Slide", "Header & / Footer",
// "Ink to / Text"); a tie breaks after the later word. Japanese captions have
// no spaces, so they break after a particle (の, と, を, に) or where the
// script changes ("新しい / スライド", "インクを / テキストに変換"); a
// caption with neither ("スクリーンショット") stays on one line. Rendered
// with `white-space: pre-line`.
const MIN_UNSPACED_BREAK = 6;
const PARTICLES = new Set(['の', 'と', 'を', 'に']);

type Script = 'hiragana' | 'katakana' | 'han' | 'other';
const scriptOf = (char: string): Script =>
  /\p{Script=Hiragana}/u.test(char)
    ? 'hiragana'
    : /[\p{Script=Katakana}ー]/u.test(char)
      ? 'katakana'
      : /\p{Script=Han}/u.test(char)
        ? 'han'
        : 'other';

function breakPoints(label: string): number[] {
  if (label.includes(' ')) {
    const points: number[] = [];
    for (let index = 0; index < label.length; index += 1)
      if (label[index] === ' ') points.push(index);
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

export function captionLines(label: string): string {
  const points = breakPoints(label);
  if (!points.length) return label;
  const spaced = label.includes(' ');
  const chars = [...label];
  let split = -1;
  let longest = Number.POSITIVE_INFINITY;
  for (const point of points) {
    const first = spaced ? label.slice(0, point) : chars.slice(0, point).join('');
    const second = spaced ? label.slice(point + 1) : chars.slice(point).join('');
    const width = Math.max(first.length, second.length);
    if (width <= longest) {
      split = point;
      longest = width;
    }
  }
  return spaced
    ? `${label.slice(0, split)}\n${label.slice(split + 1)}`
    : `${chars.slice(0, split).join('')}\n${chars.slice(split).join('')}`;
}

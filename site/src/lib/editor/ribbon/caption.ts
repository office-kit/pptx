// Mac PowerPoint wraps a large ribbon button's caption onto at most two lines,
// breaking where the wider line is narrowest ("Insert Row" / "Above",
// "Animate as" / "Background"), so the button is no wider than that line.
// English breaks at spaces; Japanese, which has none, breaks between
// characters once the caption is wider than a button ("図形の" / "塗りつぶし").
let context: CanvasRenderingContext2D | null | undefined;
const FONT = '11px -apple-system, BlinkMacSystemFont, "Segoe UI", system-ui, sans-serif';
// Treat lines within a pixel as equal; PowerPoint then keeps the longer first.
const TIE_PX = 1;
// A caption narrower than a large button's icon area stays on one line.
const ONE_LINE_PX = 48;
// Japanese line-breaking rules: no line starts with these.
const NO_LINE_START = new Set('ーァィゥェォャュョッぁぃぅぇぉゃゅょっ、。)）」』・:：');
const PARTICLES = new Set('のへをとて');
const KATAKANA = /[゠-ヿ]/;

function measure(text: string): number {
  if (context === undefined)
    context =
      typeof document === 'undefined' ? null : document.createElement('canvas').getContext('2d');
  if (!context) return text.length * 7;
  context.font = FONT;
  return context.measureText(text).width;
}

function balanced(candidates: readonly [string, string][]): string | null {
  let best: string | null = null;
  let bestWidth = Number.POSITIVE_INFINITY;
  // Later breaks first, so a tie keeps the longer first line.
  for (const [first, second] of [...candidates].reverse()) {
    const width = Math.max(measure(first), measure(second));
    if (width < bestWidth - TIE_PX) {
      best = `${first}\n${second}`;
      bestWidth = width;
    }
  }
  return best;
}

export function caption(label: string): string {
  const words = label.split(' ');
  if (words.length > 1) {
    return (
      balanced(
        words
          .slice(1)
          .map((_, index) => [
            words.slice(0, index + 1).join(' '),
            words.slice(index + 1).join(' '),
          ]),
      ) ?? label
    );
  }
  // A single Latin word ("Corrections") never breaks.
  if (/^[\x20-\x7E]*$/.test(label) || measure(label) <= ONE_LINE_PX) return label;
  const chars = [...label];
  const breaks: [string, string][] = [];
  for (let at = 1; at < chars.length; at++) {
    const before = chars[at - 1]!;
    const after = chars[at]!;
    if (NO_LINE_START.has(after)) continue;
    // Break after a particle or where a katakana word starts or ends, so
    // words stay whole ("代替" / "テキスト", "前面へ" / "移動").
    if (PARTICLES.has(before) || KATAKANA.test(before) !== KATAKANA.test(after))
      breaks.push([chars.slice(0, at).join(''), chars.slice(at).join('')]);
  }
  return balanced(breaks) ?? label;
}

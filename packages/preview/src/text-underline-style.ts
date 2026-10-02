import type { TextFormat } from '@office-kit/pptx';

const escapeXml = (value: string): string =>
  value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');

// CSS has no long-dash, dash-dot, or double-wave decoration, and Chromium
// ignores thickness on dashed decorations. Repeated SVG backgrounds preserve
// these patterns without adding boxes that change wrapping.
/**
 * DrawingML underline as CSS for an inline decoration box. Pass a resolved
 * CSS color for image-based patterns. For `words`, apply only to non-space
 * text segments; keep strikethrough on a separate outer box.
 */
export function textUnderlineStyle(underline: TextFormat['underline'], color: string): string {
  if (underline === undefined || underline === false || underline === 'none') return '';
  const heavy = underline === 'heavy' || String(underline).endsWith('Heavy');
  const thickness = heavy ? ';text-decoration-thickness:0.1em' : '';
  if (underline === 'dbl') return 'text-decoration:underline;text-decoration-style:double';
  if (underline === 'wavy' || underline === 'wavyHeavy')
    return `text-decoration:underline;text-decoration-style:wavy${thickness}`;
  const custom =
    typeof underline === 'string' &&
    (underline.startsWith('dash') ||
      underline.startsWith('dotted') ||
      underline.startsWith('dotDash') ||
      underline.startsWith('dotDotDash') ||
      underline === 'wavyDbl');
  if (!custom) return `text-decoration:underline${thickness}`;
  const stroke = heavy ? 2 : 1;
  const dash = underline.startsWith('dotted')
    ? '1 2'
    : underline === 'dash' || underline === 'dashHeavy'
      ? '4 2'
      : underline.startsWith('dashLong')
        ? '8 3'
        : underline.startsWith('dotDotDash')
          ? '5 2 1 2 1 2'
          : '5 2 1 2';
  const width =
    underline === 'wavyDbl'
      ? 8
      : underline.startsWith('dotted')
        ? 3
        : underline === 'dash' || underline === 'dashHeavy'
          ? 6
          : underline.startsWith('dashLong')
            ? 11
            : underline.startsWith('dotDotDash')
              ? 13
              : 10;
  const path =
    underline === 'wavyDbl'
      ? '<path d="M0 1 Q2 -1 4 1 T8 1 M0 4 Q2 2 4 4 T8 4"/>'
      : `<path d="M0 3 H${width}" stroke-dasharray="${dash}"/>`;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="6" viewBox="0 0 ${width} 6"><g fill="none" stroke="${escapeXml(color)}" stroke-width="${stroke}">${path}</g></svg>`;
  const url = encodeURIComponent(svg).replaceAll("'", '%27');
  return `text-decoration:none;background-image:url('data:image/svg+xml,${url}');background-repeat:repeat-x;background-size:${width / 20}em 0.3em;background-position:0 100%;box-decoration-break:clone;-webkit-box-decoration-break:clone`;
}

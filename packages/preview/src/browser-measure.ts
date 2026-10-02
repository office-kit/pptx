import type { TextMeasurer } from './text-layout.ts';

/** Custom tab alignment needs the same font advances as the browser paints. */
export function browserTextMeasurer(): TextMeasurer | undefined {
  if (typeof document === 'undefined') return undefined;
  const context = document.createElement('canvas').getContext('2d');
  if (!context) return undefined;
  return (text, spec) => {
    context.font = `${spec.italic ? 'italic' : 'normal'} ${spec.bold ? 'bold' : 'normal'} ${spec.sizePx}px ${spec.family}`;
    context.fontKerning = spec.kerning === false ? 'none' : 'normal';
    // Match SVG tracking after glyph shaping, including trailing letter spacing.
    context.letterSpacing = `${spec.letterSpacingPx}px`;
    const metrics = context.measureText(text);
    return {
      widthPx: metrics.width,
      ascentPx: metrics.fontBoundingBoxAscent,
      descentPx: metrics.fontBoundingBoxDescent,
      lineGapPx: 0,
    };
  };
}

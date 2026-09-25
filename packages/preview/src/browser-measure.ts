import type { TextMeasurer } from './text-layout.ts';

/** Custom tab alignment needs the same font advances as the browser paints. */
export function browserTextMeasurer(): TextMeasurer | undefined {
  if (typeof document === 'undefined') return undefined;
  const context = document.createElement('canvas').getContext('2d');
  if (!context) return undefined;
  return (text, spec) => {
    context.font = `${spec.italic ? 'italic' : 'normal'} ${spec.bold ? 'bold' : 'normal'} ${spec.sizePx}px ${spec.family}`;
    const metrics = context.measureText(text);
    return {
      widthPx: metrics.width + Math.max(0, [...text].length - 1) * spec.letterSpacingPx,
      ascentPx: metrics.fontBoundingBoxAscent,
      descentPx: metrics.fontBoundingBoxDescent,
      lineGapPx: 0,
    };
  };
}

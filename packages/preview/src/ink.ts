import { getInkBounds, type InkStroke } from '@office-kit/pptx';

/** Transparent, tightly cropped fallback artwork for constant-width native ink. */
export const renderInkToSvg = (strokes: ReadonlyArray<InkStroke>): string => {
  const bounds = getInkBounds(strokes);
  const paths = strokes
    .map((stroke) => {
      const points = stroke.points.map((point) => ({
        x: Math.round(point.x) - bounds.x,
        y: Math.round(point.y) - bounds.y,
      }));
      const width = Math.round(stroke.widthEmu);
      // getInkBounds validates the color, widths and every point before emission.
      const color = stroke.color.toUpperCase();
      if (points.length === 1)
        return `<circle cx="${points[0]!.x}" cy="${points[0]!.y}" r="${width / 2}" fill="${color}"/>`;
      return `<path d="${points.map((point, index) => `${index ? 'L' : 'M'}${point.x} ${point.y}`).join(' ')}" fill="none" stroke="${color}" stroke-width="${width}" stroke-linecap="round" stroke-linejoin="round"/>`;
    })
    .join('');
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${bounds.cx / 9525}" height="${bounds.cy / 9525}" viewBox="0 0 ${bounds.cx} ${bounds.cy}">${paths}</svg>`;
};

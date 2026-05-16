// Shape-level SVG renderer for the playground.
//
// pptx-kit doesn't model the full DrawingML rendering pipeline — that
// would be a renderer in its own right (see python-pptx-renderer,
// pptxgenjs-renderer, etc). What we do here is build a reasonable
// approximation: geometry by preset name, fill / stroke / rotation /
// flip applied, embedded pictures shown via data URL, and text laid
// out as centered SVG <text>.
//
// Good enough to confirm "the shapes are where I expect" without
// pulling in a real OOXML renderer.

import {
  getShapeBounds,
  getShapeFill,
  getShapeFlip,
  getShapeImageBytes,
  getShapeImageFormat,
  getShapeKind,
  getShapePreset,
  getShapeRotation,
  getShapeStroke,
  getShapeText,
  getSlideBackground,
  getSlideShapes,
  getSlideSize,
  type PresentationData,
  type ShapeFill,
  type ShapeStroke,
  type SlideData,
  type SlideShapeData,
} from 'pptx-kit';

// 16:9 fallback in EMU (10" × 5.625" doesn't exist as a real preset;
// 13.333" × 7.5" widescreen does — see ECMA-376 §19.3.1.39).
const DEFAULT_SIZE = { width: 12_192_000, height: 6_858_000 };

const u8ToBase64 = (data: Uint8Array): string => {
  let s = '';
  const chunk = 0x8000;
  for (let i = 0; i < data.length; i += chunk) {
    s += String.fromCharCode(...data.subarray(i, i + chunk));
  }
  return btoa(s);
};

const imageMime: Record<string, string> = {
  png: 'image/png',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  tiff: 'image/tiff',
  webp: 'image/webp',
  svg: 'image/svg+xml',
};

const escapeXml = (s: string): string =>
  s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');

const safeColor = (c: string | null | undefined): string | null => {
  if (!c) return null;
  // pptx-kit normalizes scheme tokens as "scheme:<token>". The playground
  // is not theme-aware, so substitute a recognisable placeholder gray
  // when the resolved color isn't an #RRGGBB literal.
  if (c.startsWith('scheme:')) return '#9CA3AF';
  if (c.startsWith('#')) return c;
  if (/^[0-9A-Fa-f]{6}$/.test(c)) return `#${c}`;
  return c;
};

// Returns the SVG attributes (fill, stroke, stroke-width) for a shape's
// solid / inherited / none paints. The renderer falls back to a tinted
// gray when a shape inherits its fill from the layout — pptx-kit doesn't
// resolve the inheritance chain at read time, and the playground stays
// readable either way.
const paint = (fill: ShapeFill, stroke: ShapeStroke): {
  fill: string;
  stroke: string;
  strokeWidth: number;
} => {
  let fillColor = '#E5E7EB'; // inherit fallback
  if (fill.kind === 'solid') fillColor = safeColor(fill.color) ?? fillColor;
  else if (fill.kind === 'none') fillColor = 'none';
  else if (fill.kind === 'gradient') fillColor = '#FDBA74';
  else if (fill.kind === 'pattern') fillColor = '#BFDBFE';
  else if (fill.kind === 'image') fillColor = '#DDD6FE';

  let strokeColor = '#9CA3AF';
  let strokeWidth = 9525; // 1pt default
  if (stroke.kind === 'solid') {
    strokeColor = safeColor(stroke.color) ?? strokeColor;
    if (stroke.widthEmu !== undefined) strokeWidth = stroke.widthEmu;
  } else if (stroke.kind === 'none') {
    strokeColor = 'none';
    strokeWidth = 0;
  }
  return { fill: fillColor, stroke: strokeColor, strokeWidth };
};

// ---------------------------------------------------------------------------
// Preset geometry → normalized [0,1] points.

const polygon = (n: number, rotation = -Math.PI / 2): Array<[number, number]> => {
  const out: Array<[number, number]> = [];
  for (let i = 0; i < n; i++) {
    const a = rotation + (i * 2 * Math.PI) / n;
    out.push([0.5 + 0.5 * Math.cos(a), 0.5 + 0.5 * Math.sin(a)]);
  }
  return out;
};

const star = (points: number, innerRatio = 0.42): Array<[number, number]> => {
  const out: Array<[number, number]> = [];
  const rotation = -Math.PI / 2;
  for (let i = 0; i < points * 2; i++) {
    const a = rotation + (i * Math.PI) / points;
    const r = i % 2 === 0 ? 0.5 : 0.5 * innerRatio;
    out.push([0.5 + r * Math.cos(a), 0.5 + r * Math.sin(a)]);
  }
  return out;
};

// Each entry returns normalized polygon points in [0,1] for a 1×1 bounding
// box; the caller scales by the actual shape bounds.
const PRESET_POINTS: Record<string, () => Array<[number, number]>> = {
  triangle: () => [[0.5, 0], [1, 1], [0, 1]],
  rtTriangle: () => [[0, 0], [1, 1], [0, 1]],
  diamond: () => [[0.5, 0], [1, 0.5], [0.5, 1], [0, 0.5]],
  parallelogram: () => [[0.25, 0], [1, 0], [0.75, 1], [0, 1]],
  trapezoid: () => [[0.25, 0], [0.75, 0], [1, 1], [0, 1]],
  pentagon: () => polygon(5),
  hexagon: () => polygon(6),
  heptagon: () => polygon(7),
  octagon: () => polygon(8),
  decagon: () => polygon(10),
  dodecagon: () => polygon(12),
  star4: () => star(4),
  star5: () => star(5),
  star6: () => star(6),
  star7: () => star(7),
  star8: () => star(8),
  star10: () => star(10),
  star12: () => star(12),
  star16: () => star(16),
  star24: () => star(24),
  star32: () => star(32),
  rightArrow: () => [
    [0, 0.3],
    [0.65, 0.3],
    [0.65, 0],
    [1, 0.5],
    [0.65, 1],
    [0.65, 0.7],
    [0, 0.7],
  ],
  leftArrow: () => [
    [1, 0.3],
    [0.35, 0.3],
    [0.35, 0],
    [0, 0.5],
    [0.35, 1],
    [0.35, 0.7],
    [1, 0.7],
  ],
  upArrow: () => [
    [0.3, 1],
    [0.3, 0.35],
    [0, 0.35],
    [0.5, 0],
    [1, 0.35],
    [0.7, 0.35],
    [0.7, 1],
  ],
  downArrow: () => [
    [0.3, 0],
    [0.3, 0.65],
    [0, 0.65],
    [0.5, 1],
    [1, 0.65],
    [0.7, 0.65],
    [0.7, 0],
  ],
  leftRightArrow: () => [
    [0, 0.5],
    [0.18, 0.2],
    [0.18, 0.35],
    [0.82, 0.35],
    [0.82, 0.2],
    [1, 0.5],
    [0.82, 0.8],
    [0.82, 0.65],
    [0.18, 0.65],
    [0.18, 0.8],
  ],
  upDownArrow: () => [
    [0.5, 0],
    [0.2, 0.18],
    [0.35, 0.18],
    [0.35, 0.82],
    [0.2, 0.82],
    [0.5, 1],
    [0.8, 0.82],
    [0.65, 0.82],
    [0.65, 0.18],
    [0.8, 0.18],
  ],
  chevron: () => [
    [0, 0],
    [0.7, 0],
    [1, 0.5],
    [0.7, 1],
    [0, 1],
    [0.3, 0.5],
  ],
};

const renderTextNode = (
  text: string,
  cx: number,
  cy: number,
  fontPx: number,
  color: string,
): string => {
  if (!text.trim()) return '';
  const lines = text.split('\n').slice(0, 6);
  const lineHeight = fontPx * 1.15;
  const totalHeight = lineHeight * (lines.length - 1);
  const startY = cy - totalHeight / 2 + fontPx * 0.35;
  const tspans = lines
    .map((line, i) => {
      const trimmed = line.length > 32 ? `${line.slice(0, 30)}…` : line;
      return `<tspan x="${cx}" y="${startY + i * lineHeight}">${escapeXml(trimmed)}</tspan>`;
    })
    .join('');
  return `<text text-anchor="middle" font-size="${fontPx}" font-family="sans-serif" fill="${color}">${tspans}</text>`;
};

// EMU helper.
const E = (n: number): string => n.toFixed(0);

const renderShape = (shape: SlideShapeData): string => {
  const bounds = getShapeBounds(shape);
  if (!bounds) return '';
  const { x, y, w, h } = bounds;
  if (w <= 0 || h <= 0) return '';

  const kind = getShapeKind(shape);
  const fill = getShapeFill(shape);
  const stroke = getShapeStroke(shape);
  const rotation = getShapeRotation(shape);
  const flip = getShapeFlip(shape) ?? { horizontal: false, vertical: false };
  const cx = x + w / 2;
  const cy = y + h / 2;

  // Build the transform string. Rotation around the shape's center,
  // then flips around the same point if set.
  const transforms: string[] = [];
  if (rotation !== 0) transforms.push(`rotate(${rotation} ${E(cx)} ${E(cy)})`);
  if (flip.horizontal) transforms.push(`translate(${E(2 * cx)} 0) scale(-1 1)`);
  if (flip.vertical) transforms.push(`translate(0 ${E(2 * cy)}) scale(1 -1)`);
  const transform = transforms.length > 0 ? ` transform="${transforms.join(' ')}"` : '';

  const text = getShapeText(shape);
  const fontPx = Math.max(80_000, Math.min(180_000, h * 0.12));
  const textColor = '#1f2937';
  const textNode = text ? renderTextNode(text, cx, cy, fontPx, textColor) : '';

  if (kind === 'picture') {
    const bytes = getShapeImageBytes(shape);
    const format = getShapeImageFormat(shape);
    if (bytes && format) {
      const mime = imageMime[format] ?? 'application/octet-stream';
      const dataUrl = `data:${mime};base64,${u8ToBase64(bytes)}`;
      return `<g${transform}><image x="${E(x)}" y="${E(y)}" width="${E(w)}" height="${E(h)}" href="${dataUrl}" preserveAspectRatio="none"/></g>`;
    }
    // Image bytes missing → render a placeholder.
    return `<g${transform}><rect x="${E(x)}" y="${E(y)}" width="${E(w)}" height="${E(h)}" fill="#F3F4F6" stroke="#9CA3AF" stroke-width="${E(9525)}" stroke-dasharray="${E(50000)},${E(30000)}"/>${renderTextNode('Image', cx, cy, fontPx, '#6B7280')}</g>`;
  }

  if (kind === 'connector') {
    const p = paint(fill, stroke);
    const sw = p.strokeWidth || 19_050; // 2pt fallback
    // Connectors are drawn as a single line. Flips flip the diagonal.
    let x1 = x as number;
    let y1 = y as number;
    let x2 = (x as number) + (w as number);
    let y2 = (y as number) + (h as number);
    if (flip.horizontal) {
      x1 = (x as number) + (w as number);
      x2 = x as number;
    }
    if (flip.vertical) {
      y1 = (y as number) + (h as number);
      y2 = y as number;
    }
    return `<line x1="${E(x1)}" y1="${E(y1)}" x2="${E(x2)}" y2="${E(y2)}" stroke="${p.stroke}" stroke-width="${E(sw)}" stroke-linecap="round"${transform}/>`;
  }

  if (kind === 'graphicFrame' || kind === 'group') {
    const label = kind === 'graphicFrame' ? 'Chart / Table' : 'Group';
    return `<g${transform}><rect x="${E(x)}" y="${E(y)}" width="${E(w)}" height="${E(h)}" fill="#F3F4F6" stroke="#9CA3AF" stroke-width="${E(9525)}" stroke-dasharray="${E(50000)},${E(30000)}"/>${renderTextNode(text || label, cx, cy, fontPx, '#6B7280')}</g>`;
  }

  // kind === 'shape'
  const preset = getShapePreset(shape) ?? 'rect';
  const p = paint(fill, stroke);

  if (preset === 'rect') {
    return `<g${transform}><rect x="${E(x)}" y="${E(y)}" width="${E(w)}" height="${E(h)}" fill="${p.fill}" stroke="${p.stroke}" stroke-width="${E(p.strokeWidth)}"/>${textNode}</g>`;
  }
  if (preset === 'roundRect') {
    const r = E(Math.min(w, h) * 0.18);
    return `<g${transform}><rect x="${E(x)}" y="${E(y)}" width="${E(w)}" height="${E(h)}" rx="${r}" ry="${r}" fill="${p.fill}" stroke="${p.stroke}" stroke-width="${E(p.strokeWidth)}"/>${textNode}</g>`;
  }
  if (preset === 'ellipse' || preset === 'oval') {
    return `<g${transform}><ellipse cx="${E(cx)}" cy="${E(cy)}" rx="${E(w / 2)}" ry="${E(h / 2)}" fill="${p.fill}" stroke="${p.stroke}" stroke-width="${E(p.strokeWidth)}"/>${textNode}</g>`;
  }

  const pointsFn = PRESET_POINTS[preset];
  if (pointsFn) {
    const points = pointsFn()
      .map(([nx, ny]) => `${E(x + nx * w)},${E(y + ny * h)}`)
      .join(' ');
    return `<g${transform}><polygon points="${points}" fill="${p.fill}" stroke="${p.stroke}" stroke-width="${E(p.strokeWidth)}"/>${textNode}</g>`;
  }

  // Unrecognised preset → fall back to a labelled rectangle.
  return `<g${transform}><rect x="${E(x)}" y="${E(y)}" width="${E(w)}" height="${E(h)}" fill="${p.fill}" stroke="${p.stroke}" stroke-width="${E(p.strokeWidth)}"/>${textNode || renderTextNode(preset, cx, cy, fontPx, textColor)}</g>`;
};

export const renderSlideSvg = (pres: PresentationData, slide: SlideData): string => {
  const size = getSlideSize(pres) ?? DEFAULT_SIZE;
  const W = size.width;
  const H = size.height;
  const bg = getSlideBackground(slide);
  const bgColor =
    bg.kind === 'solid' ? safeColor(bg.color) ?? '#FFFFFF' : '#FFFFFF';

  const shapesSvg = getSlideShapes(slide).map(renderShape).join('');

  return [
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" preserveAspectRatio="xMidYMid meet">`,
    `<rect width="${W}" height="${H}" fill="${bgColor}"/>`,
    shapesSvg,
    '</svg>',
  ].join('');
};

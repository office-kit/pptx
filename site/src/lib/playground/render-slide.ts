// Per-slide SVG renderer for the playground.
//
// pptx-kit does not ship a full DrawingML renderer — that would be a
// project in its own right (see python-pptx-renderer, pptxgenjs viewer,
// or the LibreOffice headless pipeline). What we do here is build a
// reasonable approximation by:
//
//   1. Painting each shape's preset geometry on the slide canvas at its
//      actual EMU bounds, with the rotation / flip / solid fill / stroke
//      pptx-kit reports.
//   2. Laying the shape's text on top via an SVG `<foreignObject>` that
//      hosts real HTML — that's how we get proper word wrap, per-run
//      bold / italic / underline / font, paragraph alignment, vertical
//      anchor, and bullets without re-implementing line-break logic.
//   3. Resolving theme colors (`scheme:accent1`, `scheme:tx1`, ...) via
//      `getPresentationTheme` so brand colors light up correctly.
//
// Embedded charts, tables, SmartArt, gradient / pattern / picture fills,
// the full text-property inheritance cascade (rPr → defRPr → lstStyle →
// placeholder → master → theme), and custom geometry stay as labelled
// placeholders — proper handling needs a real renderer.

import {
  getParagraphAlignment,
  getParagraphBullet,
  getParagraphLevel,
  getPresentationTheme,
  getShapeBoundsResolved,
  getShapeFill,
  getShapeFlip,
  getShapeImageBytes,
  getShapeImageFormat,
  getShapeKind,
  getShapeParagraphCount,
  getShapePlaceholderType,
  getShapePreset,
  getShapeRotation,
  getShapeRunCount,
  getShapeRunFormat,
  getShapeRunText,
  getShapeStroke,
  getShapeTextAnchor,
  getShapeTextMargins,
  getSlideBackground,
  getSlideShapes,
  getSlideSize,
  type PresentationData,
  type PresentationTheme,
  type ShapeBounds,
  type ShapeFill,
  type ShapeStroke,
  type SlideData,
  type SlideShapeData,
  type TextFormat,
} from 'pptx-kit';

// Widescreen 16:9 fallback in EMU (13.333" × 7.5"), the PowerPoint
// default since 2013. See ECMA-376 §19.3.1.39 `SlideSizeType`.
const DEFAULT_SIZE = { width: 12_192_000, height: 6_858_000 };

// The renderer projects EMU coordinates onto a CSS-pixel-at-96-DPI grid
// before emitting them. 1 CSS px = 9525 EMU at 96 DPI, so this lets us
// use natural-looking numbers in the SVG (viewBox ≈ 1280×720 for a
// 16:9 slide) and matches what HTML inside `<foreignObject>` expects.
// Real browsers refuse to render text when CSS font-size grows into
// the hundreds of thousands of pixels (which is what happens if you
// keep EMU as the SVG user unit).
const EMU_PER_PX = 9525;

// CSS px per typographic point.
const PX_PER_PT = 96 / 72;

const DEFAULT_BODY_PT = 18;
const DEFAULT_TITLE_PT = 36;
const DEFAULT_FONT = 'Calibri, "Helvetica Neue", Arial, sans-serif';

// Default body inset (PowerPoint default), per ECMA-376: 91440 EMU
// horizontal × 45720 EMU vertical.
const DEFAULT_INSET_X = 91_440;
const DEFAULT_INSET_Y = 45_720;

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

// ---------------------------------------------------------------------------
// Color resolution.

const SCHEME_TO_THEME: Record<string, keyof Omit<PresentationTheme, 'name'>> = {
  tx1: 'dark1',
  bg1: 'light1',
  tx2: 'dark2',
  bg2: 'light2',
  dk1: 'dark1',
  lt1: 'light1',
  dk2: 'dark2',
  lt2: 'light2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hlink: 'hyperlink',
  folHlink: 'followedHyperlink',
};

const normalizeHex = (s: string): string => {
  if (s.startsWith('#')) return s;
  if (/^[0-9A-Fa-f]{6}$/.test(s)) return `#${s}`;
  if (/^[0-9A-Fa-f]{8}$/.test(s)) return `#${s.slice(2)}`; // alpha-prefixed
  return s;
};

const resolveColor = (
  c: string | null | undefined,
  theme: PresentationTheme | null,
  fallback = '#1F2937',
): string => {
  if (!c) return fallback;
  if (c.startsWith('scheme:')) {
    const token = c.slice('scheme:'.length);
    if (theme) {
      const key = SCHEME_TO_THEME[token];
      if (key) return normalizeHex(theme[key]);
    }
    // Sensible per-token fallbacks when the theme is missing.
    if (token === 'tx1' || token === 'dk1') return '#000000';
    if (token === 'bg1' || token === 'lt1') return '#FFFFFF';
    if (token === 'tx2' || token === 'dk2') return '#1F2937';
    if (token === 'bg2' || token === 'lt2') return '#E5E7EB';
    return fallback;
  }
  return normalizeHex(c);
};

// ---------------------------------------------------------------------------
// Fill / stroke paint with theme resolution.

const paint = (
  fill: ShapeFill,
  stroke: ShapeStroke,
  theme: PresentationTheme | null,
  isPlaceholder: boolean,
): { fill: string; stroke: string; strokeWidth: number } => {
  let fillColor: string;
  if (fill.kind === 'solid') {
    fillColor = resolveColor(fill.color, theme, '#E5E7EB');
  } else if (fill.kind === 'none') {
    fillColor = 'none';
  } else if (fill.kind === 'gradient') {
    fillColor = '#FDBA74';
  } else if (fill.kind === 'pattern') {
    fillColor = '#BFDBFE';
  } else if (fill.kind === 'image') {
    fillColor = '#DDD6FE';
  } else {
    // `inherit`: placeholders inherit from layout/master, which is usually
    // transparent for text placeholders. Drawing them as a filled tile
    // would obscure real shapes — leave them transparent.
    fillColor = isPlaceholder ? 'none' : '#F3F4F6';
  }

  let strokeColor = 'none';
  let strokeWidth = 0;
  if (stroke.kind === 'solid') {
    strokeColor = resolveColor(stroke.color, theme, '#9CA3AF');
    strokeWidth = stroke.widthEmu ?? 9_525; // 1pt
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

// ---------------------------------------------------------------------------
// Text body rendering via foreignObject + XHTML.

const ALIGNMENT_TO_CSS: Record<string, string> = {
  left: 'left',
  center: 'center',
  right: 'right',
  justify: 'justify',
};

const ANCHOR_TO_CSS: Record<string, string> = {
  top: 'flex-start',
  center: 'center',
  bottom: 'flex-end',
};

const placeholderDefaultPt = (phType: string | null): number => {
  if (phType === 'title' || phType === 'ctrTitle') return DEFAULT_TITLE_PT;
  if (phType === 'subTitle') return 24;
  return DEFAULT_BODY_PT;
};

const bulletChar = (level: number): string =>
  level <= 0 ? '•' : level === 1 ? '◦' : '▪';

const renderRun = (
  text: string,
  format: TextFormat | null,
  theme: PresentationTheme | null,
  defaultPt: number,
): string => {
  if (text === '') return '';
  const styles: string[] = [];
  const sizePt = format?.size ?? defaultPt;
  styles.push(`font-size:${(sizePt * PX_PER_PT).toFixed(2)}px`);
  styles.push(`line-height:1.2`);
  if (format?.font) styles.push(`font-family:${escapeXml(format.font)}, ${DEFAULT_FONT}`);
  if (format?.bold) styles.push('font-weight:700');
  if (format?.italic) styles.push('font-style:italic');
  const underline = format?.underline;
  if (underline !== undefined && underline !== false && underline !== 'none') {
    styles.push('text-decoration:underline');
  }
  if (format?.color !== undefined && format.color !== null) {
    styles.push(`color:${resolveColor(format.color, theme, '#000000')}`);
  }
  return `<span style="${styles.join(';')}">${escapeXml(text)}</span>`;
};

const renderTextBody = (
  shape: SlideShapeData,
  bounds: { x: number; y: number; w: number; h: number },
  theme: PresentationTheme | null,
  phType: string | null,
): string => {
  let paragraphCount: number;
  try {
    paragraphCount = getShapeParagraphCount(shape);
  } catch {
    return '';
  }
  if (paragraphCount === 0) return '';

  const defaultPt = placeholderDefaultPt(phType);
  const anchor = getShapeTextAnchor(shape) ?? 'top';
  const margins = getShapeTextMargins(shape);
  const lIns = margins?.left ?? DEFAULT_INSET_X;
  const tIns = margins?.top ?? DEFAULT_INSET_Y;
  const rIns = margins?.right ?? DEFAULT_INSET_X;
  const bIns = margins?.bottom ?? DEFAULT_INSET_Y;

  const innerX = bounds.x + lIns;
  const innerY = bounds.y + tIns;
  const innerW = Math.max(0, bounds.w - lIns - rIns);
  const innerH = Math.max(0, bounds.h - tIns - bIns);
  if (innerW <= 0 || innerH <= 0) return '';

  const paragraphs: string[] = [];
  let hasAnyText = false;
  for (let p = 0; p < paragraphCount; p++) {
    let runCount: number;
    try {
      runCount = getShapeRunCount(shape, p);
    } catch {
      runCount = 0;
    }
    const align = getParagraphAlignment(shape, p) ?? 'left';
    const level = getParagraphLevel(shape, p);
    const bulletStyle = getParagraphBullet(shape, p);
    const runs: string[] = [];
    for (let r = 0; r < runCount; r++) {
      let txt = '';
      try {
        txt = getShapeRunText(shape, p, r);
      } catch {
        continue;
      }
      const fmt = getShapeRunFormat(shape, p, r);
      if (txt) hasAnyText = true;
      runs.push(renderRun(txt, fmt, theme, defaultPt));
    }
    const pStyles: string[] = [
      `margin:0`,
      `padding:0`,
      `text-align:${ALIGNMENT_TO_CSS[align] ?? 'left'}`,
      // Leading indent for nested bullet levels.
      level > 0 ? `padding-left:${(level * 24 * PX_PER_PT).toFixed(2)}px` : '',
    ].filter(Boolean);
    let prefix = '';
    const explicitChar =
      bulletStyle !== null &&
      typeof bulletStyle === 'object' &&
      'char' in bulletStyle
        ? bulletStyle.char
        : null;
    const showBullet =
      bulletStyle === 'bullet' || explicitChar !== null || (bulletStyle !== 'none' && level > 0);
    if (showBullet) {
      const char = explicitChar ?? bulletChar(level);
      prefix = `<span style="margin-right:${(0.4 * defaultPt * PX_PER_PT).toFixed(2)}px">${escapeXml(char)}</span>`;
    }
    paragraphs.push(
      `<p style="${pStyles.join(';')}">${prefix}${runs.join('') || '&#8203;'}</p>`,
    );
  }
  if (!hasAnyText) return '';

  const justify = ANCHOR_TO_CSS[anchor] ?? 'flex-start';
  // The default text color matters: the site is dark-mode, so without
  // an explicit color CSS inherits white-on-white and the text
  // disappears. Resolve to the theme's `tx1` (usually near-black on a
  // light slide) and let per-run colors override.
  const defaultColor = resolveColor('scheme:tx1', theme, '#000000');
  // Use an inset div for the body, full-bleed foreignObject so the SVG
  // transform / clipping behaves cleanly. Word-break:break-word keeps
  // long URLs / words from overflowing the shape.
  const body = `<div xmlns="http://www.w3.org/1999/xhtml" style="display:flex;flex-direction:column;justify-content:${justify};width:100%;height:100%;box-sizing:border-box;overflow:hidden;font-family:${DEFAULT_FONT};color:${defaultColor};word-break:break-word">${paragraphs.join('')}</div>`;
  return `<foreignObject x="${E(innerX)}" y="${E(innerY)}" width="${E(innerW)}" height="${E(innerH)}">${body}</foreignObject>`;
};

// ---------------------------------------------------------------------------
// Per-shape geometry rendering.

// EMU → CSS px coordinate. We project the whole slide onto a 96-DPI grid
// so the SVG numbers are friendly to the browser's HTML/CSS layer
// (which kicks in inside `<foreignObject>`).
const E = (n: number): string => (n / EMU_PER_PX).toFixed(2);

const renderShape = (
  shape: SlideShapeData,
  pres: PresentationData,
  theme: PresentationTheme | null,
): string => {
  // Use the inheriting resolver: shape → layout → master. Slide
  // placeholders routinely defer geometry to the master, so the literal
  // `getShapeBounds` returns null for them and the preview would hide
  // every title / body / footer slot.
  const bounds = getShapeBoundsResolved(pres, shape);
  if (!bounds) return '';
  const x = bounds.x as number;
  const y = bounds.y as number;
  const w = bounds.w as number;
  const h = bounds.h as number;
  if (w <= 0 || h <= 0) return '';

  const kind = getShapeKind(shape);
  const fill = getShapeFill(shape);
  const stroke = getShapeStroke(shape);
  const rotation = getShapeRotation(shape);
  const flip = getShapeFlip(shape) ?? { horizontal: false, vertical: false };
  const phType = getShapePlaceholderType(shape);
  const cx = x + w / 2;
  const cy = y + h / 2;

  // Build the transform string. Rotation around the shape's centre,
  // then flips around the same point if set.
  const transforms: string[] = [];
  if (rotation !== 0) transforms.push(`rotate(${rotation} ${E(cx)} ${E(cy)})`);
  if (flip.horizontal) transforms.push(`translate(${E(2 * cx)} 0) scale(-1 1)`);
  if (flip.vertical) transforms.push(`translate(0 ${E(2 * cy)}) scale(1 -1)`);
  const transform = transforms.length > 0 ? ` transform="${transforms.join(' ')}"` : '';

  const textOverlay = kind === 'shape' || kind === 'graphicFrame'
    ? renderTextBody(shape, { x, y, w, h }, theme, phType)
    : '';

  if (kind === 'picture') {
    const bytes = getShapeImageBytes(shape);
    const format = getShapeImageFormat(shape);
    if (bytes && format) {
      const mime = imageMime[format] ?? 'application/octet-stream';
      const dataUrl = `data:${mime};base64,${u8ToBase64(bytes)}`;
      return `<g${transform}><image x="${E(x)}" y="${E(y)}" width="${E(w)}" height="${E(h)}" href="${dataUrl}" preserveAspectRatio="none"/></g>`;
    }
    return `<g${transform}><rect x="${E(x)}" y="${E(y)}" width="${E(w)}" height="${E(h)}" fill="#F3F4F6" stroke="#9CA3AF" stroke-width="${E(9_525)}" stroke-dasharray="${E(50_000)},${E(30_000)}"/></g>`;
  }

  if (kind === 'connector') {
    const p = paint(fill, stroke, theme, false);
    const sw = p.strokeWidth || 19_050;
    let x1 = x;
    let y1 = y;
    let x2 = x + w;
    let y2 = y + h;
    if (flip.horizontal) {
      x1 = x + w;
      x2 = x;
    }
    if (flip.vertical) {
      y1 = y + h;
      y2 = y;
    }
    const strokeColor =
      p.stroke === 'none' ? resolveColor('scheme:tx1', theme, '#1F2937') : p.stroke;
    return `<line x1="${E(x1)}" y1="${E(y1)}" x2="${E(x2)}" y2="${E(y2)}" stroke="${strokeColor}" stroke-width="${E(sw)}" stroke-linecap="round"${transform}/>`;
  }

  if (kind === 'group') {
    return `<g${transform}><rect x="${E(x)}" y="${E(y)}" width="${E(w)}" height="${E(h)}" fill="none" stroke="#9CA3AF" stroke-width="${E(9_525)}" stroke-dasharray="${E(50_000)},${E(30_000)}"/></g>`;
  }

  const p = paint(fill, stroke, theme, phType !== null);

  if (kind === 'graphicFrame') {
    // Chart / table / SmartArt frame. Draw a faint placeholder behind
    // any text the frame might carry.
    return `<g${transform}><rect x="${E(x)}" y="${E(y)}" width="${E(w)}" height="${E(h)}" fill="${p.fill === 'none' ? '#F9FAFB' : p.fill}" stroke="#9CA3AF" stroke-width="${E(9_525)}" stroke-dasharray="${E(50_000)},${E(30_000)}"/>${textOverlay}</g>`;
  }

  // kind === 'shape'
  const preset = getShapePreset(shape) ?? 'rect';

  let geomSvg: string;
  if (preset === 'rect') {
    geomSvg = `<rect x="${E(x)}" y="${E(y)}" width="${E(w)}" height="${E(h)}" fill="${p.fill}" stroke="${p.stroke}" stroke-width="${E(p.strokeWidth)}"/>`;
  } else if (preset === 'roundRect') {
    const r = E(Math.min(w, h) * 0.18);
    geomSvg = `<rect x="${E(x)}" y="${E(y)}" width="${E(w)}" height="${E(h)}" rx="${r}" ry="${r}" fill="${p.fill}" stroke="${p.stroke}" stroke-width="${E(p.strokeWidth)}"/>`;
  } else if (preset === 'ellipse' || preset === 'oval') {
    geomSvg = `<ellipse cx="${E(cx)}" cy="${E(cy)}" rx="${E(w / 2)}" ry="${E(h / 2)}" fill="${p.fill}" stroke="${p.stroke}" stroke-width="${E(p.strokeWidth)}"/>`;
  } else {
    const pointsFn = PRESET_POINTS[preset];
    if (pointsFn) {
      const points = pointsFn()
        .map(([nx, ny]) => `${E(x + nx * w)},${E(y + ny * h)}`)
        .join(' ');
      geomSvg = `<polygon points="${points}" fill="${p.fill}" stroke="${p.stroke}" stroke-width="${E(p.strokeWidth)}"/>`;
    } else {
      // Unrecognised preset — fall back to a labelled rectangle.
      geomSvg = `<rect x="${E(x)}" y="${E(y)}" width="${E(w)}" height="${E(h)}" fill="${p.fill}" stroke="${p.stroke}" stroke-width="${E(p.strokeWidth)}"/>`;
    }
  }

  return `<g${transform}>${geomSvg}${textOverlay}</g>`;
};

// ---------------------------------------------------------------------------
// Slide composition.

export const renderSlideSvg = (pres: PresentationData, slide: SlideData): string => {
  const size = getSlideSize(pres) ?? DEFAULT_SIZE;
  const W = size.width as number;
  const H = size.height as number;
  const theme = getPresentationTheme(pres);

  const bg = getSlideBackground(slide);
  let bgColor = '#FFFFFF';
  if (bg.kind === 'solid') {
    bgColor = resolveColor(bg.color, theme, '#FFFFFF');
  } else if (theme && bg.kind === 'inherit') {
    bgColor = normalizeHex(theme.light1);
  }

  const shapesSvg = getSlideShapes(slide).map((s) => renderShape(s, pres, theme)).join('');

  return [
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${E(W)} ${E(H)}" preserveAspectRatio="xMidYMid meet">`,
    `<rect width="${E(W)}" height="${E(H)}" fill="${bgColor}"/>`,
    shapesSvg,
    '</svg>',
  ].join('');
};

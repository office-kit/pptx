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
  getShapeChartKind,
  getShapeChartSpec,
  getShapeImageBytes,
  getShapeImageFillBytes,
  getShapeImagePartName,
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
  getTableDimensions,
  isChartShape,
  isTableShape,
  type PresentationData,
  type PresentationTheme,
  type ChartSpec,
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
// Single quotes only — these strings are emitted into `style="..."`
// attributes, so any embedded double quote would close the attribute
// early and silently drop every property after it (notably `color:`,
// which then inherits the site's dark-mode white = invisible text).
const DEFAULT_FONT = "Calibri, 'Helvetica Neue', Arial, sans-serif";

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

// Fallback for when pptx-kit's `getShapeImageFormat` returns `null` —
// usually that means the magic bytes don't match a recognised format
// (EMF / WMF / HEIC / etc.), but the file extension on the media part
// is often enough to get the right MIME for the browser to try.
const EXT_TO_MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  webp: 'image/webp',
  svg: 'image/svg+xml',
  avif: 'image/avif',
  heic: 'image/heic',
  heif: 'image/heif',
};

const mimeFromPartName = (name: string | null): string | null => {
  if (!name) return null;
  const dot = name.lastIndexOf('.');
  if (dot < 0) return null;
  const ext = name.slice(dot + 1).toLowerCase();
  return EXT_TO_MIME[ext] ?? null;
};

// Render the picture if pptx-kit handed us bytes; fall back to a
// labelled placeholder describing why nothing is drawn. EMF / WMF
// pictures still won't display (no browser can decode them) but the
// label tells the user what's there.
const renderPicture = (
  shape: SlideShapeData,
  x: number,
  y: number,
  w: number,
  h: number,
  transform: string,
  textOverlay: string,
  bytes: Uint8Array | null,
  format: string | null,
): string => {
  let mime: string | null = null;
  if (bytes && format) {
    mime = imageMime[format] ?? null;
  }
  if (bytes && !mime) {
    // Format detection failed (likely EMF/WMF/HEIC). Try the part
    // name's extension — most browsers can still render HEIC / AVIF
    // and friends if labelled correctly.
    mime = mimeFromPartName(getShapeImagePartName(shape));
  }
  if (bytes && mime) {
    const dataUrl = `data:${mime};base64,${u8ToBase64(bytes)}`;
    return `<g${transform}><image x="${E(x)}" y="${E(y)}" width="${E(w)}" height="${E(h)}" href="${dataUrl}" xlink:href="${dataUrl}" preserveAspectRatio="none"/>${textOverlay}</g>`;
  }
  const label = !bytes
    ? 'picture (no bytes)'
    : `picture (${format ?? 'unknown'}${bytes ? `, ${bytes.byteLength} B` : ''})`;
  return `<g${transform}><rect x="${E(x)}" y="${E(y)}" width="${E(w)}" height="${E(h)}" fill="#F3F4F6" stroke="#9CA3AF" stroke-width="${E(9_525)}" stroke-dasharray="${E(50_000)},${E(30_000)}"/>${renderPicturePlaceholderLabel(x, y, w, h, label)}${textOverlay}</g>`;
};

const renderPicturePlaceholderLabel = (
  x: number,
  y: number,
  w: number,
  h: number,
  text: string,
): string => {
  const cx = x + w / 2;
  const cy = y + h / 2;
  // <title> shows on hover in every SVG viewer — useful when the
  // shape is small enough that the inline <text> gets cropped.
  const title = `<title>${escapeXml(text)}</title>`;
  const label = `<text x="${E(cx)}" y="${E(cy)}" text-anchor="middle" dominant-baseline="middle" font-family="sans-serif" font-weight="600" font-size="${(13 * PX_PER_PT).toFixed(2)}" fill="#374151">${escapeXml(text)}</text>`;
  return `${title}${label}`;
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

// ---------------------------------------------------------------------------
// Chart rendering. We don't ship a real chart engine — the goal is just
// enough visual fidelity that a column / line / pie chart in the deck
// looks like a chart in the preview. Axes, gridlines, data-label
// styling, secondary axes, and chart styles are intentionally skipped.

// CSS px helpers that operate directly in CSS-px space (the chart math
// uses plain numbers without EMU conversion).
const px = (n: number): string => n.toFixed(2);

const accentSequence = (theme: PresentationTheme | null): string[] => {
  const fallbacks = ['#5B9BD5', '#ED7D31', '#A5A5A5', '#FFC000', '#4472C4', '#70AD47'];
  if (!theme) return fallbacks;
  const hexes = [theme.accent1, theme.accent2, theme.accent3, theme.accent4, theme.accent5, theme.accent6]
    .map((c) => normalizeHex(c))
    .filter((c): c is string => /^#[0-9A-Fa-f]{6}$/.test(c));
  return hexes.length > 0 ? hexes : fallbacks;
};

// Project EMU bounds → CSS-px chart frame. Title and legend get fixed
// vertical strips; the plot area takes whatever's left.
interface ChartFrame {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
  readonly plotX: number;
  readonly plotY: number;
  readonly plotW: number;
  readonly plotH: number;
  readonly titleY: number;
  readonly legendY: number;
}

const layoutChart = (xEmu: number, yEmu: number, wEmu: number, hEmu: number, hasTitle: boolean): ChartFrame => {
  const x = xEmu / EMU_PER_PX;
  const y = yEmu / EMU_PER_PX;
  const w = wEmu / EMU_PER_PX;
  const h = hEmu / EMU_PER_PX;
  const titleStrip = hasTitle ? 18 : 0;
  const legendStrip = 18;
  const padding = 8;
  return {
    x,
    y,
    w,
    h,
    plotX: x + padding,
    plotY: y + titleStrip + padding,
    plotW: Math.max(0, w - 2 * padding),
    plotH: Math.max(0, h - titleStrip - legendStrip - 2 * padding),
    titleY: y + titleStrip - 2,
    legendY: y + h - legendStrip / 2,
  };
};

const seriesMinMax = (spec: ChartSpec): { min: number; max: number } => {
  let min = Infinity;
  let max = -Infinity;
  for (const s of spec.series) {
    for (const v of s.values) {
      if (v !== null && Number.isFinite(v)) {
        if (v < min) min = v;
        if (v > max) max = v;
      }
    }
  }
  if (!Number.isFinite(min)) min = 0;
  if (!Number.isFinite(max)) max = 1;
  if (max === min) max = min + 1;
  if (min > 0) min = 0; // include the zero line, like PowerPoint does
  return { min, max };
};

const renderChartTitle = (f: ChartFrame, title: string): string => {
  if (!title) return '';
  return `<text x="${px(f.x + f.w / 2)}" y="${px(f.titleY)}" text-anchor="middle" dominant-baseline="middle" font-family="sans-serif" font-size="13" fill="#1F2937" font-weight="600">${escapeXml(title)}</text>`;
};

const renderChartLegend = (f: ChartFrame, names: ReadonlyArray<string>, colors: ReadonlyArray<string>): string => {
  if (names.length === 0) return '';
  const itemPx = Math.min(140, f.w / names.length);
  const totalW = itemPx * names.length;
  const startX = f.x + (f.w - totalW) / 2;
  const out: string[] = [];
  for (let i = 0; i < names.length; i++) {
    const cx = startX + i * itemPx;
    const swatchX = cx + 4;
    const swatchY = f.legendY - 4;
    const labelX = swatchX + 14;
    out.push(`<rect x="${px(swatchX)}" y="${px(swatchY)}" width="9" height="9" fill="${colors[i % colors.length]}"/>`);
    out.push(`<text x="${px(labelX)}" y="${px(f.legendY)}" dominant-baseline="middle" font-family="sans-serif" font-size="11" fill="#374151">${escapeXml(names[i] ?? `Series ${i + 1}`)}</text>`);
  }
  return out.join('');
};

const renderColumnChart = (f: ChartFrame, spec: ChartSpec, colors: ReadonlyArray<string>): string => {
  const cats = spec.categories;
  if (cats.length === 0) return '';
  const { min, max } = seriesMinMax(spec);
  const range = max - min;
  const groupW = f.plotW / cats.length;
  const barW = (groupW * 0.8) / spec.series.length;
  const baseY = f.plotY + f.plotH - ((0 - min) / range) * f.plotH;
  const out: string[] = [];
  for (let c = 0; c < cats.length; c++) {
    for (let s = 0; s < spec.series.length; s++) {
      const v = spec.series[s]?.values[c] ?? 0;
      const x0 = f.plotX + c * groupW + groupW * 0.1 + s * barW;
      const top = f.plotY + f.plotH - ((v - min) / range) * f.plotH;
      const y0 = Math.min(top, baseY);
      const h = Math.abs(top - baseY);
      out.push(`<rect x="${px(x0)}" y="${px(y0)}" width="${px(barW)}" height="${px(h)}" fill="${colors[s % colors.length]}"/>`);
    }
  }
  // Zero baseline for visual reference.
  out.push(`<line x1="${px(f.plotX)}" y1="${px(baseY)}" x2="${px(f.plotX + f.plotW)}" y2="${px(baseY)}" stroke="#9CA3AF" stroke-width="0.5"/>`);
  return out.join('');
};

const renderBarChart = (f: ChartFrame, spec: ChartSpec, colors: ReadonlyArray<string>): string => {
  const cats = spec.categories;
  if (cats.length === 0) return '';
  const { min, max } = seriesMinMax(spec);
  const range = max - min;
  const groupH = f.plotH / cats.length;
  const barH = (groupH * 0.8) / spec.series.length;
  const baseX = f.plotX + ((0 - min) / range) * f.plotW;
  const out: string[] = [];
  for (let c = 0; c < cats.length; c++) {
    for (let s = 0; s < spec.series.length; s++) {
      const v = spec.series[s]?.values[c] ?? 0;
      const y0 = f.plotY + c * groupH + groupH * 0.1 + s * barH;
      const tip = f.plotX + ((v - min) / range) * f.plotW;
      const x0 = Math.min(tip, baseX);
      const w = Math.abs(tip - baseX);
      out.push(`<rect x="${px(x0)}" y="${px(y0)}" width="${px(w)}" height="${px(barH)}" fill="${colors[s % colors.length]}"/>`);
    }
  }
  out.push(`<line x1="${px(baseX)}" y1="${px(f.plotY)}" x2="${px(baseX)}" y2="${px(f.plotY + f.plotH)}" stroke="#9CA3AF" stroke-width="0.5"/>`);
  return out.join('');
};

const renderLineChart = (f: ChartFrame, spec: ChartSpec, colors: ReadonlyArray<string>, fill: boolean): string => {
  const cats = spec.categories;
  if (cats.length === 0) return '';
  const { min, max } = seriesMinMax(spec);
  const range = max - min;
  const step = cats.length > 1 ? f.plotW / (cats.length - 1) : 0;
  const baseY = f.plotY + f.plotH - ((0 - min) / range) * f.plotH;
  const out: string[] = [];
  out.push(`<line x1="${px(f.plotX)}" y1="${px(baseY)}" x2="${px(f.plotX + f.plotW)}" y2="${px(baseY)}" stroke="#E5E7EB" stroke-width="0.5"/>`);
  for (let s = 0; s < spec.series.length; s++) {
    const series = spec.series[s];
    if (!series) continue;
    const color = colors[s % colors.length];
    const pts: Array<[number, number]> = [];
    for (let c = 0; c < cats.length; c++) {
      const v = series.values[c] ?? 0;
      const xp = f.plotX + c * step;
      const yp = f.plotY + f.plotH - ((v - min) / range) * f.plotH;
      pts.push([xp, yp]);
    }
    const dPath = pts.map(([xp, yp], i) => `${i === 0 ? 'M' : 'L'}${px(xp)},${px(yp)}`).join(' ');
    if (fill) {
      const areaPath = `${dPath} L${px(f.plotX + (cats.length - 1) * step)},${px(baseY)} L${px(f.plotX)},${px(baseY)} Z`;
      out.push(`<path d="${areaPath}" fill="${color}" fill-opacity="0.25" stroke="none"/>`);
    }
    out.push(`<path d="${dPath}" fill="none" stroke="${color}" stroke-width="1.5" stroke-linejoin="round" stroke-linecap="round"/>`);
    // Data point markers.
    for (const [xp, yp] of pts) {
      out.push(`<circle cx="${px(xp)}" cy="${px(yp)}" r="2.2" fill="${color}"/>`);
    }
  }
  return out.join('');
};

const renderPieChart = (f: ChartFrame, spec: ChartSpec, colors: ReadonlyArray<string>, doughnut: boolean): string => {
  const series = spec.series[0];
  if (!series) return '';
  const values = series.values.map((v) => Math.max(0, v ?? 0));
  const total = values.reduce((a, b) => a + b, 0);
  if (total === 0) return '';
  const radius = Math.min(f.plotW, f.plotH) / 2 - 2;
  const cx = f.plotX + f.plotW / 2;
  const cy = f.plotY + f.plotH / 2;
  const innerR = doughnut ? radius * 0.55 : 0;
  let acc = -Math.PI / 2; // start at 12 o'clock
  const out: string[] = [];
  for (let i = 0; i < values.length; i++) {
    const v = values[i] ?? 0;
    const angle = (v / total) * 2 * Math.PI;
    const start = acc;
    const end = acc + angle;
    acc = end;
    const largeArc = angle > Math.PI ? 1 : 0;
    const ox1 = cx + radius * Math.cos(start);
    const oy1 = cy + radius * Math.sin(start);
    const ox2 = cx + radius * Math.cos(end);
    const oy2 = cy + radius * Math.sin(end);
    const color = colors[i % colors.length];
    if (doughnut) {
      const ix1 = cx + innerR * Math.cos(start);
      const iy1 = cy + innerR * Math.sin(start);
      const ix2 = cx + innerR * Math.cos(end);
      const iy2 = cy + innerR * Math.sin(end);
      const d = `M${px(ox1)},${px(oy1)} A${px(radius)},${px(radius)} 0 ${largeArc} 1 ${px(ox2)},${px(oy2)} L${px(ix2)},${px(iy2)} A${px(innerR)},${px(innerR)} 0 ${largeArc} 0 ${px(ix1)},${px(iy1)} Z`;
      out.push(`<path d="${d}" fill="${color}" stroke="#FFFFFF" stroke-width="0.6"/>`);
    } else {
      const d = `M${px(cx)},${px(cy)} L${px(ox1)},${px(oy1)} A${px(radius)},${px(radius)} 0 ${largeArc} 1 ${px(ox2)},${px(oy2)} Z`;
      out.push(`<path d="${d}" fill="${color}" stroke="#FFFFFF" stroke-width="0.6"/>`);
    }
  }
  return out.join('');
};

const renderChart = (
  shape: SlideShapeData,
  x: number,
  y: number,
  w: number,
  h: number,
  transform: string,
  theme: PresentationTheme | null,
): string | null => {
  let spec: ChartSpec | null = null;
  try {
    spec = getShapeChartSpec(shape);
  } catch {
    return null;
  }
  if (!spec) return null;
  const colors = accentSequence(theme);
  const f = layoutChart(x, y, w, h, !!spec.title);
  const seriesNamesForLegend: string[] =
    spec.kind === 'pie' || spec.kind === 'doughnut'
      ? Array.from(spec.categories)
      : spec.series.map((s) => s.name);
  const seriesColorsForLegend: string[] =
    spec.kind === 'pie' || spec.kind === 'doughnut'
      ? spec.categories.map((_, i) => colors[i % colors.length] ?? '#888')
      : spec.series.map((_, i) => colors[i % colors.length] ?? '#888');

  // Count finite values across all series — when zero, draw a hint
  // label so an empty chart isn't indistinguishable from a working
  // one whose data we failed to read.
  let finiteCount = 0;
  for (const s of spec.series) {
    for (const v of s.values) if (v !== null && Number.isFinite(v)) finiteCount++;
  }

  let plot = '';
  switch (spec.kind) {
    case 'column':
    case 'bar':
      // pptx-kit reports both as `bar` / `column` via separate `kind`;
      // legacy `barDir` distinction. We branch on `kind`.
      plot = spec.kind === 'column' ? renderColumnChart(f, spec, colors) : renderBarChart(f, spec, colors);
      break;
    case 'line':
      plot = renderLineChart(f, spec, colors, false);
      break;
    case 'area':
      plot = renderLineChart(f, spec, colors, true);
      break;
    case 'pie':
      plot = renderPieChart(f, spec, colors, false);
      break;
    case 'doughnut':
      plot = renderPieChart(f, spec, colors, true);
      break;
    default:
      return null;
  }

  const emptyHint =
    finiteCount === 0
      ? `<text x="${px(f.plotX + f.plotW / 2)}" y="${px(f.plotY + f.plotH / 2)}" text-anchor="middle" dominant-baseline="middle" font-family="sans-serif" font-size="12" fill="#9CA3AF">${escapeXml(`chart (${spec.kind}) — no data`)}</text>`
      : '';

  return [
    `<g${transform}>`,
    `<rect x="${px(f.x)}" y="${px(f.y)}" width="${px(f.w)}" height="${px(f.h)}" fill="#FFFFFF" stroke="#E5E7EB" stroke-width="0.6"/>`,
    renderChartTitle(f, spec.title ?? ''),
    plot,
    emptyHint,
    renderChartLegend(f, seriesNamesForLegend, seriesColorsForLegend),
    '</g>',
  ].join('');
};

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
    return renderPicture(
      shape,
      x,
      y,
      w,
      h,
      transform,
      textOverlay,
      getShapeImageBytes(shape),
      getShapeImageFormat(shape),
    );
  }

  // Shapes with an image fill (`<p:sp>` + `<a:blipFill>` instead of a
  // solid / gradient / pattern). PowerPoint's "Insert Picture from
  // File" and several third-party tools emit pictures this way rather
  // than as top-level `<p:pic>`.
  if (kind === 'shape' && fill.kind === 'image') {
    return renderPicture(
      shape,
      x,
      y,
      w,
      h,
      transform,
      textOverlay,
      getShapeImageFillBytes(shape),
      getShapeImageFormat(shape),
    );
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
    // Chart frames: render a simplified chart (bars / line / pie /
    // etc.) from the chart spec when we can read it. Tables and
    // SmartArt still render as labelled placeholders — drawing them
    // needs a real layout engine.
    if (isChartShape(shape)) {
      const chartSvg = renderChart(shape, x, y, w, h, transform, theme);
      if (chartSvg) return chartSvg;
    }
    let label = 'graphicFrame';
    try {
      if (isChartShape(shape)) {
        const chartKind = getShapeChartKind(shape);
        label = chartKind ? `chart (${chartKind})` : 'chart';
      } else if (isTableShape(shape)) {
        const t = getTableDimensions(shape);
        label = `table (${t.rows}×${t.cols})`;
      }
    } catch {
      // Fall through with the generic label.
    }
    return `<g${transform}><rect x="${E(x)}" y="${E(y)}" width="${E(w)}" height="${E(h)}" fill="${p.fill === 'none' ? '#F9FAFB' : p.fill}" stroke="#9CA3AF" stroke-width="${E(9_525)}" stroke-dasharray="${E(50_000)},${E(30_000)}"/>${renderPicturePlaceholderLabel(x, y, w, h, label)}${textOverlay}</g>`;
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
    `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 ${E(W)} ${E(H)}" preserveAspectRatio="xMidYMid meet">`,
    `<rect width="${E(W)}" height="${E(H)}" fill="${bgColor}"/>`,
    shapesSvg,
    '</svg>',
  ].join('');
};

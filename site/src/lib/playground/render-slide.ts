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
  getShapeGradientFill,
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
  getGroupChildren,
  getGroupTransform,
  getSlideBackground,
  getSlideBackgroundImageBytes,
  getSlideShapes,
  getSlideSize,
  getTableCellAlignment,
  getTableCellFill,
  getTableCellText,
  getTableCells,
  getTableColumnWidths,
  getTableDimensions,
  getTableRowHeights,
  isChartShape,
  isTableShape,
  type PresentationData,
  type PresentationTheme,
  type ChartSpec,
  type GradientFillOptions,
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

// Conservative defaults — pptx-kit doesn't walk the lstStyle / master
// cascade to find the real default size, so we err on the small side
// to avoid the rendered text overflowing tight placeholders. PowerPoint
// real defaults are 44pt title / 28pt body; we'd clip many titles at
// those sizes since the placeholder height comes from the master,
// which expects autofit.
const DEFAULT_BODY_PT = 14;
const DEFAULT_TITLE_PT = 26;
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

// Module-scoped id counter for SVG `<defs>` references (gradients,
// patterns). Each renderSlideSvg call mints fresh ids; collisions
// across slides don't matter because each slide's SVG is a separate
// document. Plain monotonic counter is fine.
let nextDefId = 0;
const mintId = (): string => `pkdef-${(nextDefId++).toString(36)}`;

// `<linearGradient>` definition + `fill="url(#…)"` reference, projected
// from pptx-kit's `{ stops, angleDeg }` shape onto SVG's
// objectBoundingBox unit cube. ECMA-376 measures `angleDeg` clockwise
// from 3 o'clock, which matches the trig below (0° = +x, 90° = +y).
const gradientDef = (
  grad: GradientFillOptions,
  theme: PresentationTheme | null,
): { defs: string; fillAttr: string } => {
  const id = mintId();
  const angleRad = ((grad.angleDeg ?? 0) * Math.PI) / 180;
  const dx = Math.cos(angleRad) / 2;
  const dy = Math.sin(angleRad) / 2;
  const x1 = 0.5 - dx;
  const y1 = 0.5 - dy;
  const x2 = 0.5 + dx;
  const y2 = 0.5 + dy;
  const stops = grad.stops
    .map((s) => `<stop offset="${s.offset.toFixed(4)}" stop-color="${resolveColor(s.color, theme, '#E5E7EB')}"/>`)
    .join('');
  const defs = `<defs><linearGradient id="${id}" gradientUnits="objectBoundingBox" x1="${x1.toFixed(4)}" y1="${y1.toFixed(4)}" x2="${x2.toFixed(4)}" y2="${y2.toFixed(4)}">${stops}</linearGradient></defs>`;
  return { defs, fillAttr: `url(#${id})` };
};

interface PaintResult {
  fill: string;
  stroke: string;
  strokeWidth: number;
  /** Extra SVG `<defs>` the caller should emit before the shape. */
  defs: string;
}

const paint = (
  shape: SlideShapeData | null,
  fill: ShapeFill,
  stroke: ShapeStroke,
  theme: PresentationTheme | null,
  isPlaceholder: boolean,
): PaintResult => {
  let fillColor: string;
  let defs = '';
  if (fill.kind === 'solid') {
    fillColor = resolveColor(fill.color, theme, '#E5E7EB');
  } else if (fill.kind === 'none') {
    fillColor = 'none';
  } else if (fill.kind === 'gradient') {
    // Best-effort: paint a real `<linearGradient>` when the
    // `<a:gradFill>` is on the shape itself. Inherited gradients (no
    // explicit gradFill on the shape) fall back to the orange tint
    // since we don't yet walk the layout/master cascade for fills.
    const grad = shape ? getShapeGradientFill(shape) : null;
    if (grad) {
      const built = gradientDef(grad, theme);
      defs = built.defs;
      fillColor = built.fillAttr;
    } else {
      fillColor = '#FDBA74';
    }
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
  return { fill: fillColor, stroke: strokeColor, strokeWidth, defs };
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
  // Additional block arrows beyond the cardinal four.
  bentArrow: () => [
    [0, 0.45],
    [0.55, 0.45],
    [0.55, 0.25],
    [0.55, 0.05],
    [0.95, 0.05],
    [0.95, 0.55],
    [0.8, 0.55],
    [0.8, 0.85],
    [0, 0.85],
  ],
  // Right-pointing pentagon (often used for flowcharts).
  homePlate: () => [
    [0, 0],
    [0.75, 0],
    [1, 0.5],
    [0.75, 1],
    [0, 1],
  ],
  // Hearts / smileyFace / cloud handled via path renderers below.
};

// Path-based shape renderers — for shapes that can't be expressed as a
// closed polygon (curves, multiple sub-paths, etc.). Returns an SVG
// `d` attribute already scaled to the shape's bounding box.
const PRESET_PATHS: Record<string, (x: number, y: number, w: number, h: number) => string> = {
  // Wedge callouts: rect / round-rect / ellipse body covering ~85% of
  // the bounding box plus a triangular tail pointing down-left.
  // Default adj1/adj2 values aren't read from the XML; the tail
  // position is approximate.
  wedgeRectCallout: (x, y, w, h) => {
    const bodyH = h * 0.78;
    const tailTipX = x + w * 0.12;
    const tailTipY = y + h;
    const tailBaseLeft = x + w * 0.18;
    const tailBaseRight = x + w * 0.32;
    const bodyB = y + bodyH;
    return `M${x},${y} L${x + w},${y} L${x + w},${bodyB} L${tailBaseRight},${bodyB} L${tailTipX},${tailTipY} L${tailBaseLeft},${bodyB} L${x},${bodyB} Z`;
  },
  wedgeRoundRectCallout: (x, y, w, h) => {
    const r = Math.min(w, h) * 0.08;
    const bodyH = h * 0.78;
    const tailTipX = x + w * 0.12;
    const tailTipY = y + h;
    const tailBaseLeft = x + w * 0.18;
    const tailBaseRight = x + w * 0.32;
    const bodyB = y + bodyH;
    return `M${x + r},${y} L${x + w - r},${y} A${r},${r} 0 0 1 ${x + w},${y + r} L${x + w},${bodyB - r} A${r},${r} 0 0 1 ${x + w - r},${bodyB} L${tailBaseRight},${bodyB} L${tailTipX},${tailTipY} L${tailBaseLeft},${bodyB} L${x + r},${bodyB} A${r},${r} 0 0 1 ${x},${bodyB - r} L${x},${y + r} A${r},${r} 0 0 1 ${x + r},${y} Z`;
  },
  wedgeEllipseCallout: (x, y, w, h) => {
    const bodyCy = y + h * 0.39;
    const bodyRy = h * 0.39;
    const cx = x + w / 2;
    const bodyRx = w / 2;
    // Ellipse perimeter via two arcs, then a triangle to the tail.
    const tailTipX = x + w * 0.12;
    const tailTipY = y + h;
    const tailBaseAngle = 1.5; // radians from positive x — bottom-ish
    const tailBase1X = cx + bodyRx * Math.cos(tailBaseAngle - 0.18);
    const tailBase1Y = bodyCy + bodyRy * Math.sin(tailBaseAngle - 0.18);
    const tailBase2X = cx + bodyRx * Math.cos(tailBaseAngle + 0.18);
    const tailBase2Y = bodyCy + bodyRy * Math.sin(tailBaseAngle + 0.18);
    return `M${tailBase1X},${tailBase1Y} A${bodyRx},${bodyRy} 0 1 0 ${tailBase2X},${tailBase2Y} L${tailTipX},${tailTipY} Z`;
  },
  // Cloud callout — body is 8 lobes around an ellipse, plus a small
  // dot trail towards the tail point.
  cloudCallout: (x, y, w, h) => {
    const bodyH = h * 0.78;
    const cx = x + w / 2;
    const cy = y + bodyH / 2;
    const rx = (w / 2) * 0.92;
    const ry = (bodyH / 2) * 0.85;
    const lobes = 10;
    const path: string[] = [];
    for (let i = 0; i < lobes; i++) {
      const a = (i / lobes) * 2 * Math.PI - Math.PI / 2;
      const lobeRx = rx * 0.32;
      const lobeRy = ry * 0.32;
      const px0 = cx + rx * Math.cos(a);
      const py0 = cy + ry * Math.sin(a);
      if (i === 0) path.push(`M${px0 - lobeRx},${py0}`);
      path.push(`A${lobeRx},${lobeRy} 0 1 1 ${px0 + lobeRx},${py0}`);
      const nextA = ((i + 1) / lobes) * 2 * Math.PI - Math.PI / 2;
      const nextX = cx + rx * Math.cos(nextA) - lobeRx;
      const nextY = cy + ry * Math.sin(nextA);
      path.push(`L${nextX},${nextY}`);
    }
    path.push('Z');
    // Two small trailing circles for the tail.
    const tailX = x + w * 0.18;
    const tailY = y + h * 0.95;
    return `${path.join(' ')} M${tailX - 6},${tailY - 14} a4,3 0 1 0 1,0 Z M${tailX},${tailY} a6,4 0 1 0 1,0 Z`;
  },
  // Hearts / sun / lightning / smiley — common decorative shapes.
  heart: (x, y, w, h) => {
    const cx = x + w / 2;
    const top = y + h * 0.27;
    return `M${cx},${y + h} C${x},${y + h * 0.55} ${x},${top} ${cx},${y + h * 0.4} C${x + w},${top} ${x + w},${y + h * 0.55} ${cx},${y + h} Z`;
  },
  sun: (x, y, w, h) => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const innerR = Math.min(w, h) * 0.25;
    const outerR = Math.min(w, h) * 0.5;
    const rays = 12;
    const path: string[] = [];
    for (let i = 0; i < rays * 2; i++) {
      const r = i % 2 === 0 ? outerR : innerR;
      const a = ((i / (rays * 2)) * 2 * Math.PI) - Math.PI / 2;
      const px0 = cx + r * Math.cos(a);
      const py0 = cy + r * Math.sin(a);
      path.push(`${i === 0 ? 'M' : 'L'}${px0},${py0}`);
    }
    path.push('Z');
    return path.join(' ');
  },
  smileyFace: (x, y, w, h) => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const r = Math.min(w, h) / 2 - 1;
    const eyeR = r * 0.07;
    const eyeOff = r * 0.32;
    const mouthW = r * 0.5;
    const mouthY = cy + r * 0.18;
    // Face circle, two eye holes (subpath, even-odd-filled), smiling arc.
    return `M${cx + r},${cy} A${r},${r} 0 1 0 ${cx - r},${cy} A${r},${r} 0 1 0 ${cx + r},${cy} Z M${cx - eyeOff + eyeR},${cy - eyeOff} A${eyeR},${eyeR} 0 1 1 ${cx - eyeOff - eyeR},${cy - eyeOff} A${eyeR},${eyeR} 0 1 1 ${cx - eyeOff + eyeR},${cy - eyeOff} Z M${cx + eyeOff + eyeR},${cy - eyeOff} A${eyeR},${eyeR} 0 1 1 ${cx + eyeOff - eyeR},${cy - eyeOff} A${eyeR},${eyeR} 0 1 1 ${cx + eyeOff + eyeR},${cy - eyeOff} Z M${cx - mouthW},${mouthY} Q${cx},${mouthY + r * 0.32} ${cx + mouthW},${mouthY}`;
  },
  lightningBolt: (x, y, w, h) => {
    return `M${x + w * 0.5},${y} L${x + w * 0.15},${y + h * 0.55} L${x + w * 0.45},${y + h * 0.55} L${x + w * 0.3},${y + h} L${x + w * 0.85},${y + h * 0.4} L${x + w * 0.55},${y + h * 0.4} L${x + w * 0.7},${y} Z`;
  },
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
  // PowerPoint uses tight line-height (~1.0) by default for placeholders;
  // the previous 1.2 left enough vertical slack to push the top/bottom of
  // glyphs outside short placeholders.
  styles.push(`line-height:1.05`);
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
  // overflow:visible — the placeholder's box is sized for PowerPoint's
  // autofit text, which we don't model. With overflow:hidden, long
  // titles or body bullets that should have shrunk to fit instead get
  // their ascenders / descenders clipped. Letting them bleed outside
  // the shape is uglier but never *loses* information.
  const body = `<div xmlns="http://www.w3.org/1999/xhtml" style="display:flex;flex-direction:column;justify-content:${justify};width:100%;height:100%;box-sizing:border-box;overflow:visible;font-family:${DEFAULT_FONT};color:${defaultColor};word-break:break-word">${paragraphs.join('')}</div>`;
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

const layoutChart = (
  xEmu: number,
  yEmu: number,
  wEmu: number,
  hEmu: number,
  hasTitle: boolean,
  hasAxes: boolean,
): ChartFrame => {
  const x = xEmu / EMU_PER_PX;
  const y = yEmu / EMU_PER_PX;
  const w = wEmu / EMU_PER_PX;
  const h = hEmu / EMU_PER_PX;
  const titleStrip = hasTitle ? 18 : 0;
  const legendStrip = 18;
  const padding = 8;
  // Carve out axis gutters so numeric / category labels have room.
  const yAxisGutter = hasAxes ? 40 : 0;
  const xAxisGutter = hasAxes ? 18 : 0;
  return {
    x,
    y,
    w,
    h,
    plotX: x + padding + yAxisGutter,
    plotY: y + titleStrip + padding,
    plotW: Math.max(0, w - 2 * padding - yAxisGutter),
    plotH: Math.max(0, h - titleStrip - legendStrip - xAxisGutter - 2 * padding),
    titleY: y + titleStrip - 2,
    legendY: y + h - legendStrip / 2,
  };
};

// ---------------------------------------------------------------------------
// Axis labels + gridlines for bar / column / line / area charts.

// Pick ~5 "nice" tick values between min and max. The step is rounded
// to a 1 / 2 / 5 × 10ⁿ that gives 4-6 ticks total — same rule
// Excel / PowerPoint use.
const niceTicks = (min: number, max: number, target = 5): number[] => {
  const range = max - min;
  if (range <= 0) return [min];
  const rawStep = range / target;
  const exp = Math.floor(Math.log10(rawStep));
  const base = Math.pow(10, exp);
  const m = rawStep / base;
  const stepMultiplier = m < 1.5 ? 1 : m < 3 ? 2 : m < 7 ? 5 : 10;
  const step = stepMultiplier * base;
  const start = Math.ceil(min / step) * step;
  const ticks: number[] = [];
  for (let v = start; v <= max + step / 2; v += step) ticks.push(v);
  return ticks;
};

const formatTick = (v: number): string => {
  if (v === 0) return '0';
  const abs = Math.abs(v);
  if (abs >= 1_000_000_000) return `${(v / 1_000_000_000).toFixed(1)}B`;
  if (abs >= 1_000_000) return `${(v / 1_000_000).toFixed(1)}M`;
  if (abs >= 10_000) return `${(v / 1000).toFixed(0)}K`;
  if (abs >= 1000) return `${(v / 1000).toFixed(1)}K`;
  if (Number.isInteger(v)) return v.toString();
  return v.toFixed(abs < 1 ? 2 : 1);
};

// Numeric tick labels + horizontal gridlines on the value axis (Y for
// column/line/area, X for bar). Shared between every cartesian chart
// kind so axis styling stays consistent.
interface AxisSpec {
  readonly orientation: 'vertical' | 'horizontal';
  readonly min: number;
  readonly max: number;
}

const renderValueAxis = (f: ChartFrame, axis: AxisSpec): string => {
  const ticks = niceTicks(axis.min, axis.max);
  const out: string[] = [];
  const range = axis.max - axis.min || 1;
  for (const t of ticks) {
    if (axis.orientation === 'vertical') {
      const yp = f.plotY + f.plotH - ((t - axis.min) / range) * f.plotH;
      // Gridline.
      out.push(`<line x1="${px(f.plotX)}" y1="${px(yp)}" x2="${px(f.plotX + f.plotW)}" y2="${px(yp)}" stroke="#E5E7EB" stroke-width="0.5"/>`);
      // Numeric label, right-aligned to the plot's left edge.
      out.push(`<text x="${px(f.plotX - 4)}" y="${px(yp)}" text-anchor="end" dominant-baseline="middle" font-family="sans-serif" font-size="10" fill="#6B7280">${escapeXml(formatTick(t))}</text>`);
    } else {
      const xp = f.plotX + ((t - axis.min) / range) * f.plotW;
      out.push(`<line x1="${px(xp)}" y1="${px(f.plotY)}" x2="${px(xp)}" y2="${px(f.plotY + f.plotH)}" stroke="#E5E7EB" stroke-width="0.5"/>`);
      out.push(`<text x="${px(xp)}" y="${px(f.plotY + f.plotH + 12)}" text-anchor="middle" dominant-baseline="middle" font-family="sans-serif" font-size="10" fill="#6B7280">${escapeXml(formatTick(t))}</text>`);
    }
  }
  return out.join('');
};

// Category labels along the non-value axis.
const renderCategoryAxis = (
  f: ChartFrame,
  orientation: 'horizontal' | 'vertical',
  cats: ReadonlyArray<string>,
  pointCount: number,
): string => {
  const labels: string[] = [];
  for (let i = 0; i < pointCount; i++) {
    labels.push(cats[i] ?? (i + 1).toString());
  }
  const out: string[] = [];
  if (orientation === 'horizontal') {
    // Categories along x-axis (column / line / area charts).
    const step = pointCount > 1 ? f.plotW / pointCount : 0;
    for (let i = 0; i < pointCount; i++) {
      // Center labels under each category slot.
      const cx = f.plotX + (i + 0.5) * step;
      const truncated = labels[i] !== undefined && labels[i]!.length > 14
        ? `${labels[i]!.slice(0, 12)}…`
        : labels[i] ?? '';
      out.push(`<text x="${px(cx)}" y="${px(f.plotY + f.plotH + 12)}" text-anchor="middle" dominant-baseline="middle" font-family="sans-serif" font-size="10" fill="#6B7280">${escapeXml(truncated)}</text>`);
    }
  } else {
    // Categories down the y-axis (bar chart).
    const step = pointCount > 0 ? f.plotH / pointCount : 0;
    for (let i = 0; i < pointCount; i++) {
      const cy = f.plotY + (i + 0.5) * step;
      const truncated = labels[i] !== undefined && labels[i]!.length > 14
        ? `${labels[i]!.slice(0, 12)}…`
        : labels[i] ?? '';
      out.push(`<text x="${px(f.plotX - 4)}" y="${px(cy)}" text-anchor="end" dominant-baseline="middle" font-family="sans-serif" font-size="10" fill="#6B7280">${escapeXml(truncated)}</text>`);
    }
  }
  return out.join('');
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

// Charts can ship without an explicit `<c:cat>` channel — the series'
// `<c:val>` array alone is enough; PowerPoint then labels the x-axis
// 1, 2, 3, ... Use the longest series as the point count when
// `spec.categories` is empty so those charts still plot.
const pointCount = (spec: ChartSpec): number => {
  if (spec.categories.length > 0) return spec.categories.length;
  let n = 0;
  for (const s of spec.series) if (s.values.length > n) n = s.values.length;
  return n;
};

const renderColumnChart = (f: ChartFrame, spec: ChartSpec, colors: ReadonlyArray<string>): string => {
  const N = pointCount(spec);
  if (N === 0 || spec.series.length === 0) return '';
  const { min, max } = seriesMinMax(spec);
  const range = max - min;
  const groupW = f.plotW / N;
  const barW = (groupW * 0.8) / spec.series.length;
  const baseY = f.plotY + f.plotH - ((0 - min) / range) * f.plotH;
  const out: string[] = [];
  for (let c = 0; c < N; c++) {
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
  const N = pointCount(spec);
  if (N === 0 || spec.series.length === 0) return '';
  const { min, max } = seriesMinMax(spec);
  const range = max - min;
  const groupH = f.plotH / N;
  const barH = (groupH * 0.8) / spec.series.length;
  const baseX = f.plotX + ((0 - min) / range) * f.plotW;
  const out: string[] = [];
  for (let c = 0; c < N; c++) {
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
  const N = pointCount(spec);
  if (N === 0 || spec.series.length === 0) return '';
  const { min, max } = seriesMinMax(spec);
  const range = max - min;
  const step = N > 1 ? f.plotW / (N - 1) : 0;
  const baseY = f.plotY + f.plotH - ((0 - min) / range) * f.plotH;
  const out: string[] = [];
  out.push(`<line x1="${px(f.plotX)}" y1="${px(baseY)}" x2="${px(f.plotX + f.plotW)}" y2="${px(baseY)}" stroke="#E5E7EB" stroke-width="0.5"/>`);
  for (let s = 0; s < spec.series.length; s++) {
    const series = spec.series[s];
    if (!series) continue;
    const color = colors[s % colors.length];
    const pts: Array<[number, number]> = [];
    for (let c = 0; c < N; c++) {
      const v = series.values[c] ?? 0;
      const xp = f.plotX + c * step;
      const yp = f.plotY + f.plotH - ((v - min) / range) * f.plotH;
      pts.push([xp, yp]);
    }
    const dPath = pts.map(([xp, yp], i) => `${i === 0 ? 'M' : 'L'}${px(xp)},${px(yp)}`).join(' ');
    if (fill) {
      const areaPath = `${dPath} L${px(f.plotX + (N - 1) * step)},${px(baseY)} L${px(f.plotX)},${px(baseY)} Z`;
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
  const isCartesian =
    spec.kind === 'column' ||
    spec.kind === 'bar' ||
    spec.kind === 'line' ||
    spec.kind === 'area';
  const f = layoutChart(x, y, w, h, !!spec.title, isCartesian);
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
  let axes = '';
  if (isCartesian) {
    const { min, max } = seriesMinMax(spec);
    const N = pointCount(spec);
    const valueAxis: AxisSpec =
      spec.kind === 'bar'
        ? { orientation: 'horizontal', min, max }
        : { orientation: 'vertical', min, max };
    axes = renderValueAxis(f, valueAxis);
    if (N > 0) {
      axes += renderCategoryAxis(f, spec.kind === 'bar' ? 'vertical' : 'horizontal', spec.categories, N);
    }
  }
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
    axes,
    plot,
    emptyHint,
    renderChartLegend(f, seriesNamesForLegend, seriesColorsForLegend),
    '</g>',
  ].join('');
};

// ---------------------------------------------------------------------------
// Table rendering. Real table layout (cell borders, banded rows, header
// row, merged cells, per-run text formatting) needs a much bigger pass;
// this version draws the grid, fills, and centred cell text — enough to
// recognise the table at a glance.

const ALIGNMENT_TEXT_ANCHOR: Record<string, string> = {
  left: 'start',
  center: 'middle',
  right: 'end',
  justify: 'start',
};

const renderTableCellText = (
  text: string,
  cx: number,
  cy: number,
  cw: number,
  ch: number,
  alignment: string | null,
  color: string,
): string => {
  if (!text.trim()) return '';
  // Use foreignObject so cell text wraps and aligns the same way
  // proper PowerPoint cells do.
  const ta = alignment && ALIGNMENT_TEXT_ANCHOR[alignment] !== undefined ? alignment : 'left';
  const pad = 4; // px inset for cell text
  const innerX = cx + pad;
  const innerY = cy + pad;
  const innerW = Math.max(0, cw - 2 * pad);
  const innerH = Math.max(0, ch - 2 * pad);
  if (innerW <= 0 || innerH <= 0) return '';
  const justify = ta === 'center' ? 'center' : ta === 'right' ? 'flex-end' : 'flex-start';
  const lines = text.split('\n').slice(0, 8);
  const body = lines
    .map((line) => `<div style="text-align:${ta}">${escapeXml(line)}</div>`)
    .join('');
  return `<foreignObject x="${px(innerX)}" y="${px(innerY)}" width="${px(innerW)}" height="${px(innerH)}"><div xmlns="http://www.w3.org/1999/xhtml" style="display:flex;flex-direction:column;justify-content:center;align-items:${justify};width:100%;height:100%;box-sizing:border-box;overflow:hidden;font-family:${DEFAULT_FONT};color:${color};font-size:10px;line-height:1.15;word-break:break-word">${body}</div></foreignObject>`;
};

const renderTable = (
  shape: SlideShapeData,
  x: number,
  y: number,
  w: number,
  h: number,
  transform: string,
  theme: PresentationTheme | null,
): string | null => {
  let dims: { rows: number; cols: number };
  let widths: ReadonlyArray<number>;
  let heights: ReadonlyArray<number>;
  let cells: ReadonlyArray<ReadonlyArray<unknown>>;
  try {
    dims = getTableDimensions(shape);
    widths = getTableColumnWidths(shape);
    heights = getTableRowHeights(shape);
    cells = getTableCells(shape);
  } catch {
    return null;
  }
  if (dims.rows === 0 || dims.cols === 0) return null;

  const xPx = x / EMU_PER_PX;
  const yPx = y / EMU_PER_PX;
  const widthsPx = widths.map((w0) => w0 / EMU_PER_PX);
  const heightsPx = heights.map((h0) => h0 / EMU_PER_PX);
  // If the declared widths / heights don't fill the table bounds (or
  // overflow them), scale them so the rendered grid matches the shape's
  // outer rect. Without scaling, narrow tables would render at the
  // declared cell widths but readers expect the bounding box.
  const wSum = widthsPx.reduce((a, b) => a + b, 0);
  const hSum = heightsPx.reduce((a, b) => a + b, 0);
  const wScale = wSum > 0 ? (w / EMU_PER_PX) / wSum : 1;
  const hScale = hSum > 0 ? (h / EMU_PER_PX) / hSum : 1;
  const colXs: number[] = [xPx];
  for (let c = 0; c < widthsPx.length; c++) {
    colXs.push((colXs[c] ?? xPx) + (widthsPx[c] ?? 0) * wScale);
  }
  const rowYs: number[] = [yPx];
  for (let r = 0; r < heightsPx.length; r++) {
    rowYs.push((rowYs[r] ?? yPx) + (heightsPx[r] ?? 0) * hScale);
  }

  const textColor = resolveColor('scheme:tx1', theme, '#000000');
  const out: string[] = [];
  out.push(`<g${transform}>`);
  // Whole-table backdrop so cells with no explicit fill still
  // contrast against whatever's behind.
  out.push(`<rect x="${px(xPx)}" y="${px(yPx)}" width="${px((colXs[widthsPx.length] ?? xPx) - xPx)}" height="${px((rowYs[heightsPx.length] ?? yPx) - yPx)}" fill="#FFFFFF"/>`);
  for (let r = 0; r < dims.rows; r++) {
    for (let c = 0; c < dims.cols; c++) {
      const cell = cells[r]?.[c];
      if (!cell) continue;
      const cx = colXs[c] ?? xPx;
      const cy = rowYs[r] ?? yPx;
      const cw = (colXs[c + 1] ?? cx) - cx;
      const ch = (rowYs[r + 1] ?? cy) - cy;
      const fill = getTableCellFill(cell as Parameters<typeof getTableCellFill>[0]);
      const fillColor = fill ? resolveColor(fill, theme, '#FFFFFF') : 'none';
      out.push(`<rect x="${px(cx)}" y="${px(cy)}" width="${px(cw)}" height="${px(ch)}" fill="${fillColor}" stroke="#9CA3AF" stroke-width="0.5"/>`);
      const text = getTableCellText(cell as Parameters<typeof getTableCellText>[0]);
      const align = getTableCellAlignment(cell as Parameters<typeof getTableCellAlignment>[0]);
      out.push(renderTableCellText(text, cx, cy, cw, ch, align, textColor));
    }
  }
  out.push('</g>');
  return out.join('');
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
    const p = paint(shape, fill, stroke, theme, false);
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
    // Recurse into the group's children. Their bounds live in the
    // group's internal coordinate system; an SVG transform maps that
    // onto the slide. Children are rendered the same way as
    // top-level shapes — nested groups recurse naturally.
    const xform = getGroupTransform(shape);
    const children = getGroupChildren(shape);
    if (children.length === 0) return '';
    let groupTransform = '';
    if (xform) {
      const ox = xform.outer.x as number;
      const oy = xform.outer.y as number;
      const ow = xform.outer.w as number;
      const oh = xform.outer.h as number;
      const ix = xform.inner.x as number;
      const iy = xform.inner.y as number;
      const iw = (xform.inner.w as number) || 1;
      const ih = (xform.inner.h as number) || 1;
      const sx = (ow / iw).toFixed(6);
      const sy = (oh / ih).toFixed(6);
      // translate first, then scale, so the child's natural coords
      // (px) project: ox/EMU_PER_PX + (cx - ix) * sx / EMU_PER_PX,
      // which factors as translate(ox/EMU - ix*sx/EMU) scale(sx).
      const tx = ((ox - ix * (ow / iw)) / EMU_PER_PX).toFixed(2);
      const ty = ((oy - iy * (oh / ih)) / EMU_PER_PX).toFixed(2);
      groupTransform = ` transform="translate(${tx} ${ty}) scale(${sx} ${sy})"`;
    }
    const childrenSvg = children.map((c) => renderShape(c, pres, theme)).join('');
    return `<g${groupTransform}>${childrenSvg}</g>`;
  }

  const p = paint(shape, fill, stroke, theme, phType !== null);

  if (kind === 'graphicFrame') {
    // Charts and tables get real renders. SmartArt and the
    // graphicFrame variants pptx-kit doesn't model fall through to a
    // labelled placeholder.
    if (isChartShape(shape)) {
      const chartSvg = renderChart(shape, x, y, w, h, transform, theme);
      if (chartSvg) return chartSvg;
    }
    if (isTableShape(shape)) {
      const tableSvg = renderTable(shape, x, y, w, h, transform, theme);
      if (tableSvg) return tableSvg;
    }
    let label = 'graphicFrame';
    try {
      if (isChartShape(shape)) {
        // `isChartShape` was true but renderChart returned null —
        // pptx-kit couldn't model this chart kind (3D bar, scatter,
        // bubble, stock, radar, surface, of-pie all land here).
        label = 'chart (unsupported kind)';
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
    const pathFn = PRESET_PATHS[preset];
    if (pathFn) {
      // The path generators output CSS-px coords directly (post-E).
      const d = pathFn(x / EMU_PER_PX, y / EMU_PER_PX, w / EMU_PER_PX, h / EMU_PER_PX);
      geomSvg = `<path d="${d}" fill="${p.fill}" stroke="${p.stroke}" stroke-width="${E(p.strokeWidth)}" fill-rule="evenodd"/>`;
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
  }

  return `${p.defs}<g${transform}>${geomSvg}${textOverlay}</g>`;
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

  // Slide background image (`<p:bgPr><a:blipFill>`). Paint it above the
  // solid bg-color rect so the bytes show through; shapes still draw
  // on top of the image.
  let bgImage = '';
  if (bg.kind === 'image') {
    const bytes = getSlideBackgroundImageBytes(slide);
    if (bytes) {
      const fmt = detectImageFormatLocal(bytes);
      const mime = fmt ? imageMime[fmt] ?? 'image/png' : 'image/png';
      const dataUrl = `data:${mime};base64,${u8ToBase64(bytes)}`;
      bgImage = `<image x="0" y="0" width="${E(W)}" height="${E(H)}" href="${dataUrl}" xlink:href="${dataUrl}" preserveAspectRatio="xMidYMid slice"/>`;
    }
  }

  const shapesSvg = getSlideShapes(slide).map((s) => renderShape(s, pres, theme)).join('');

  return [
    `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 ${E(W)} ${E(H)}" preserveAspectRatio="xMidYMid meet">`,
    `<rect width="${E(W)}" height="${E(H)}" fill="${bgColor}"/>`,
    bgImage,
    shapesSvg,
    '</svg>',
  ].join('');
};

// Minimal magic-byte sniffer covering the formats we already know how
// to MIME-type. Used for slide background images, which pptx-kit
// returns as raw bytes without exposing the format.
const detectImageFormatLocal = (bytes: Uint8Array): string | null => {
  if (bytes.length < 4) return null;
  if (
    bytes[0] === 0x89 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x4e &&
    bytes[3] === 0x47
  ) return 'png';
  if (bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return 'jpeg';
  if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) return 'gif';
  if (bytes[0] === 0x42 && bytes[1] === 0x4d) return 'bmp';
  if (
    bytes.length >= 12 &&
    bytes[0] === 0x52 && bytes[1] === 0x49 && bytes[2] === 0x46 && bytes[3] === 0x46 &&
    bytes[8] === 0x57 && bytes[9] === 0x45 && bytes[10] === 0x42 && bytes[11] === 0x50
  ) return 'webp';
  return null;
};

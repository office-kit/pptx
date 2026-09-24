import { resolveTextBodyRect, shapeCustomTextRect } from './text-body-rect.ts';
import { textColumnsStyle, verticalTextStyle } from './text-body-style.ts';
import { paragraphNumberLabels } from './paragraph-number-labels.ts';
// Per-slide SVG renderer for the playground.
//
// @office-kit/pptx does not ship a full DrawingML renderer — that would be a
// project in its own right (see python-pptx-renderer, pptxgenjs viewer,
// or the LibreOffice headless pipeline). What we do here is build a
// reasonable approximation by:
//
//   1. Painting each shape's preset geometry on the slide canvas at its
//      actual EMU bounds, with the rotation / flip / solid fill / stroke
//      @office-kit/pptx reports.
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

import type { getParagraphLineSpacing, getShapeEffects } from '@office-kit/pptx';
import {
  getParagraphAlignment,
  getParagraphBullet,
  getParagraphBulletStyle,
  getParagraphBulletImageBytes,
  isParagraphBulletPicture,
  getParagraphIndent,
  getParagraphLevel,
  getParagraphPropertiesEffective,
  getEffectiveColorMap,
  getParagraphSpacing,
  getPresentationFonts,
  getPresentationTheme,
  resolveDeckBodyTextColor,
  getShapeBoundsResolved,
  getShapeEffectsEffective,
  getShapeFillEffective,
  getShapeFillOpacity,
  getShapeFillColorResolved,
  getShapeFlip,
  getShapeGradientFill,
  getShapeGradientFillEffective,
  getShapePatternFill,
  getShapeBodyPrEffective,
  getShapeChartSpec,
  getShapeClickAction,
  getShapeSlide,
  getPresentationFirstSlideNumber,
  getSlides,
  isSlideHidden,
  isSlideBackgroundGraphicsHidden,
  isSlideLayoutBackgroundGraphicsHidden,
  getShapeAltTitle,
  getShapeDescription,
  getShapeHyperlinkTooltip,
  getShapeId,
  getShapeName,
  getShapeTextBodyRotationDeg,
  getShapeTextDirection,
  getShapeImageBiLevelThreshold,
  getShapeImageBrightness,
  getShapeImageLinkUrl,
  getShapeImageBytes,
  getShapeImageContrast,
  getShapeImageCrop,
  getShapeImageDuotone,
  getShapeImageFillBytes,
  getShapeImageFillLayout,
  getShapeImageIntrinsicSize,
  getShapeImageOpacity,
  getShapeImagePartName,
  getShapeImageFormat,
  getShapeAdjustValues,
  getShapeBounds,
  getShapeCustomGeometry,
  getShapeKind,
  getShapeParagraphCount,
  getShapeParagraphElements,
  getShapeRunClickAction,
  getShapeRunHyperlink,
  getShapeRunHyperlinkTooltip,
  getShapePlaceholderType,
  getShapePreset,
  getShapeXmlString,
  getShapeRotation,
  getShapeRunFormat,
  getShapeRunFormatEffective,
  getShapeStrokeEffective,
  getShapeStrokeArrow,
  getShapeStrokeCap,
  getShapeStrokeColorResolved,
  getShapeStrokeCompound,
  getShapeStrokeDash,
  getShapeStrokeJoin,
  getShapeStrokeOpacity,
  getShapeTextAnchor,
  getShapeTextAutoFitParams,
  getShapeTextAutoFit,
  getShapeTextColumns,
  getShapeTextMargins,
  getGroupChildren,
  getGroupTransform,
  getSlideBackground,
  getSlideBackgroundGradientFill,
  getSlideLayoutBackgroundGradientFill,
  getSlideMasterBackgroundGradientFill,
  getSlideBackgroundImageBytes,
  getSlideBackgroundPatternFill,
  getSlideIndex,
  getSlideLayout,
  getSlideLayoutBackground,
  getSlideLayoutBackgroundImageBytes,
  getSlideLayoutBackgroundPatternFill,
  getSlideLayoutShapes,
  getSlideMasterShapes,
  getSlideMasterBackground,
  getSlideMasterBackgroundImageBytes,
  getSlideMasterBackgroundPatternFill,
  getSlideShapes,
  isShapeHidden,
  getSlideSize,
  getTableCellAnchor,
  getTableCellMargins,
  getTableCellBorders,
  getTableCellFill,
  getTableCellParagraphs,
  getTableCellSpan,
  getTableCells,
  getTableStyleFlags,
  getTableColumnWidths,
  getTableDimensions,
  getTableRowHeights,
  isChartShape,
  isShapeImageGrayscale,
  isShapePlaceholder,
  isShapeTextBox,
  isTableShape,
  type PresentationData,
  type PresentationTheme,
  type ChartKind,
  type ChartSeries,
  type ReadChartSpec,
  type ChartTextStyle,
  type CustomGeometry,
  type ReadGradientFill,
  type ShapeFill,
  type ShapeStroke,
  type SlideData,
  type SlideShapeData,
  type TableCellParagraph,
  type ReadTextFormat,
} from '@office-kit/pptx';
import { renderEmfToSvg } from './emf.ts';
import {
  defaultMeasurer,
  layoutTextSvg,
  layoutCore,
  measureTextBodyHeight,
  substituteFamily,
  type BulletInput,
  type ColumnLayout,
  type ParaInput,
  type PieceInput,
  type RenderSlideOptions,
  type TextBodyInput,
  type TextLayoutMode,
  type TextMeasurer,
  type VerticalLayout,
} from './text-layout.ts';

export type { RenderSlideOptions, TextMeasurer, FontSpec, MeasureResult } from './text-layout.ts';

// Threaded through the shape walk so the text path (foreignObject vs pure SVG)
// and its measurer are chosen once at the top and never re-derived.
interface LayoutCtx {
  readonly mode: TextLayoutMode;
  readonly measure: TextMeasurer;
  // Product of the scale factors of every enclosing group (`<a:ext>` over
  // `<a:chExt>`). Geometry scales with the group; text does not — PowerPoint
  // and LibreOffice both keep a group child's glyphs at their authored size
  // and aspect when the group is resized, so the text path renders into the
  // group-scaled rect and cancels the scale back out (see `renderShape`).
  readonly groupScale: { readonly sx: number; readonly sy: number };
  // An odd number of ancestor reflections reverses glyph handedness.
  readonly groupReflected: boolean;
  // Whether the shapes being drawn are the slide's own. Only those carry
  // `data-pptx-shape-id`: a slide's `<p:spTgt spid>` names shape ids on the
  // slide, and the layout and master have their own id spaces where the same
  // number means a different shape.
  readonly ownShapes: boolean;
  readonly background: { id: string; width: number; height: number };
  readonly inverseGroupTransform: string;
}

const clickActionHref = (
  pres: PresentationData,
  shape: SlideShapeData,
  action: ReturnType<typeof getShapeClickAction>,
): string | undefined => {
  if (!action) return undefined;
  if (action.kind === 'url') return action.url;
  if (action.kind === 'slide') {
    const index = getSlideIndex(pres, action.slide);
    return index >= 0 ? `#slide-${index + 1}` : undefined;
  }
  const slides = getSlides(pres);
  const current = getSlideIndex(pres, getShapeSlide(shape));
  if (current < 0) return undefined;
  const backwards = action.kind === 'prevSlide' || action.kind === 'lastSlide';
  const step = backwards ? -1 : 1;
  const start =
    action.kind === 'firstSlide'
      ? 0
      : action.kind === 'lastSlide'
        ? slides.length - 1
        : current + step;
  for (let index = start; index >= 0 && index < slides.length; index += step) {
    if (!isSlideHidden(slides[index]!)) return `#slide-${index + 1}`;
  }
  return `#slide-${current + 1}`;
};

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
export const EMU_PER_PX = 9525;

// CSS px per typographic point.
const PX_PER_PT = 96 / 72;

// PowerPoint's stock master defaults — we now honour
// `<a:normAutofit fontScale=…>` when it's set on the shape's text
// body, so the title can claim its full 44pt size without
// auto-shrinking blindly. A heuristic autofit still kicks in only
// when neither an explicit normAutofit nor a small enough authored
// size keeps the text inside the placeholder.
const DEFAULT_BODY_PT = 18;
const DEFAULT_TITLE_PT = 44;
// Single quotes only — these strings are emitted into `style="..."`
// attributes, so any embedded double quote would close the attribute
// early and silently drop every property after it (notably `color:`,
// which then inherits the site's dark-mode white = invisible text).
const DEFAULT_FONT = "Calibri, 'Helvetica Neue', Arial, sans-serif";
// Bullet font when no buFont is authored or inherited. The stock PowerPoint
// template's master bodyStyle sets buFont="Arial", and its '•'/'◦' glyphs are
// smaller and higher than the theme minor face (Calibri) — so Arial, not the
// body text face, is the right default.
const DEFAULT_BULLET_FONT = 'Arial';

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

// Inlines raw image bytes as a `data:` URL, sniffing the format from the
// magic bytes (defaulting to PNG when unknown).
const bytesToDataUrl = (bytes: Uint8Array): string => {
  const fmt = detectImageFormatLocal(bytes);
  const mime = fmt ? (imageMime[fmt] ?? 'image/png') : 'image/png';
  return `data:${mime};base64,${u8ToBase64(bytes)}`;
};

// Fallback for when @office-kit/pptx's `getShapeImageFormat` returns `null` —
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

// Unsupported image formats retain a labelled placeholder so users can
// identify artwork that the preview cannot render.
const renderPicture = (
  shape: SlideShapeData,
  pres: PresentationData,
  x: number,
  y: number,
  w: number,
  h: number,
  transform: string,
  textOverlay: string,
  bytes: Uint8Array | null,
  format: string | null,
  outline: PaintResult,
): string => {
  let mime: string | null = null;
  if (bytes && format) {
    mime = imageMime[format] ?? null;
  }
  if (bytes && !mime) {
    const svg = renderEmfToSvg(bytes);
    if (svg !== null) {
      bytes = new TextEncoder().encode(svg);
      mime = 'image/svg+xml';
    }
  }
  if (bytes && !mime) {
    // Format detection failed (likely EMF/WMF/HEIC). Try the part
    // name's extension — most browsers can still render HEIC / AVIF
    // and friends if labelled correctly.
    mime = mimeFromPartName(getShapeImagePartName(shape));
  }
  if (bytes && mime) {
    const dataUrl = `data:${mime};base64,${u8ToBase64(bytes)}`;
    // Apply <a:srcRect> crop, brightness/contrast (lum),
    // and opacity (alphaModFix) so PowerPoint's "Picture Format >
    // Corrections" matches what the playground paints.
    const crop = getShapeImageCrop(shape);
    const layout = getShapeImageFillLayout(shape);
    const stretch = layout?.mode === 'stretch' ? layout : null;
    const fillL = stretch?.left ?? 0;
    const fillT = stretch?.top ?? 0;
    const fillR = stretch?.right ?? 0;
    const fillB = stretch?.bottom ?? 0;
    const hasFillOffsets = fillL !== 0 || fillT !== 0 || fillR !== 0 || fillB !== 0;
    let imgX = x + w * fillL,
      imgY = y + h * fillT,
      imgW = Math.max(0, w * (1 - fillL - fillR)),
      imgH = Math.max(0, h * (1 - fillT - fillB));
    let clipDef = '';
    let clipAttr = '';
    const cropL = crop?.left ?? 0;
    const cropT = crop?.top ?? 0;
    const cropR = crop?.right ?? 0;
    const cropB = crop?.bottom ?? 0;
    if (cropL > 0 || cropT > 0 || cropR > 0 || cropB > 0) {
      // ECMA-376 <a:srcRect> sides are fractions of the source image;
      // PowerPoint crops by adjusting the visible region. We project
      // the same effect by scaling the <image> larger and clipping it
      // to the shape's bounds.
      const scaleX = 1 / Math.max(0.001, 1 - cropL - cropR);
      const scaleY = 1 / Math.max(0.001, 1 - cropT - cropB);
      imgW *= scaleX;
      imgH *= scaleY;
      imgX -= imgW * cropL;
      imgY -= imgH * cropT;
    }
    const preset = getShapePreset(shape) ?? 'rect';
    if (crop || hasFillOffsets || preset !== 'rect') {
      const clipId = mintId();
      const geometry = pictureClipGeometry(shape, preset, x, y, w, h);
      clipDef = `<defs><clipPath id="${clipId}">${geometry}</clipPath></defs>`;
      clipAttr = ` clip-path="url(#${clipId})"`;
    }
    const brightness = getShapeImageBrightness(shape) ?? 0;
    const contrast = getShapeImageContrast(shape) ?? 0;
    const opacity = getShapeImageOpacity(shape) ?? 1;
    const grayscale = isShapeImageGrayscale(shape);
    const biLevel = getShapeImageBiLevelThreshold(shape);
    const duotone = getShapeImageDuotone(pres, shape);
    let filterAttr = '';
    if (
      brightness !== 0 ||
      contrast !== 0 ||
      grayscale ||
      biLevel !== null ||
      (duotone && (duotone.firstColor || duotone.secondColor))
    ) {
      // SVG filter pipeline: contrast/brightness (linear), then optional
      // grayscale (luminance matrix), then optional biLevel two-tone
      // (discrete table that snaps each channel to 0 or 1 at thresh).
      const fid = mintId();
      const prims: string[] = [];
      if (brightness !== 0 || contrast !== 0) {
        // DrawingML contrast is a signed percentage change, with zero neutral.
        // Scale channel distances from mid-gray; negative values reduce contrast.
        const slope = 1 + contrast;
        const intercept = brightness + (1 - slope) / 2;
        prims.push(
          `<feComponentTransfer><feFuncR type="linear" slope="${slope}" intercept="${intercept}"/><feFuncG type="linear" slope="${slope}" intercept="${intercept}"/><feFuncB type="linear" slope="${slope}" intercept="${intercept}"/></feComponentTransfer>`,
        );
      }
      if (grayscale) {
        // Luminance-preserving desaturation matrix.
        prims.push(
          `<feColorMatrix type="matrix" values="0.2126 0.7152 0.0722 0 0  0.2126 0.7152 0.0722 0 0  0.2126 0.7152 0.0722 0 0  0 0 0 1 0"/>`,
        );
      }
      if (duotone && duotone.firstColor && duotone.secondColor) {
        // Duotone: gray → lerp(firstColor, secondColor, gray). We emit
        // a feColorMatrix that desaturates first (Rec. 709 luminance),
        // then a feComponentTransfer with tableValues sampling the
        // gradient between the two colors.
        const [r1, g1, b1] = hexChannels(duotone.firstColor);
        const [r2, g2, b2] = hexChannels(duotone.secondColor);
        const steps = 16;
        const tR: string[] = [];
        const tG: string[] = [];
        const tB: string[] = [];
        for (let i = 0; i < steps; i++) {
          const t = i / (steps - 1);
          tR.push(((r1 + (r2 - r1) * t) / 255).toFixed(4));
          tG.push(((g1 + (g2 - g1) * t) / 255).toFixed(4));
          tB.push(((b1 + (b2 - b1) * t) / 255).toFixed(4));
        }
        prims.push(
          `<feColorMatrix type="matrix" values="0.2126 0.7152 0.0722 0 0  0.2126 0.7152 0.0722 0 0  0.2126 0.7152 0.0722 0 0  0 0 0 1 0"/>`,
          `<feComponentTransfer><feFuncR type="table" tableValues="${tR.join(' ')}"/><feFuncG type="table" tableValues="${tG.join(' ')}"/><feFuncB type="table" tableValues="${tB.join(' ')}"/></feComponentTransfer>`,
        );
      }
      if (biLevel !== null) {
        const t = biLevel / 100;
        // discrete tables snap channels to 0 below `t` and to 1 at/above.
        const table = `0 1`;
        // Use a step function: tableValues with 2 entries split at thresh.
        // feFuncR/G/B with type=discrete + 2-entry table snaps at the midpoint;
        // for an arbitrary threshold we shift via tableValues with more samples.
        const steps = 32;
        const vals: string[] = [];
        for (let i = 0; i < steps; i++) {
          vals.push(i / (steps - 1) >= t ? '1' : '0');
        }
        void table;
        const tableStr = vals.join(' ');
        prims.push(
          `<feComponentTransfer><feFuncR type="discrete" tableValues="${tableStr}"/><feFuncG type="discrete" tableValues="${tableStr}"/><feFuncB type="discrete" tableValues="${tableStr}"/></feComponentTransfer>`,
        );
      }
      clipDef += `<defs><filter id="${fid}">${prims.join('')}</filter></defs>`;
      filterAttr = ` filter="url(#${fid})"`;
    }
    const opacityAttr = opacity !== 1 ? ` opacity="${opacity.toFixed(3)}"` : '';
    // Paint the outline outside the image clip: otherwise half its width is
    // cut off at the mask boundary. It shares the image's rotation and flips.
    const border =
      outline.stroke !== 'none' && outline.strokeWidth > 0
        ? `<g${transform} fill="none" stroke="${outline.stroke}" stroke-width="${E(outline.strokeWidth)}"${outline.strokeAttrs ? ` ${outline.strokeAttrs}` : ''}>${pictureClipGeometry(shape, preset, x, y, w, h)}</g>`
        : '';
    if (layout?.mode === 'tile') {
      const intrinsic = getShapeImageIntrinsicSize(shape);
      if (intrinsic) {
        const sourceW = intrinsic.width * (layout.scaleX ?? 1);
        const sourceH = intrinsic.height * (layout.scaleY ?? 1);
        const tileW = sourceW * (1 - cropL - cropR);
        const tileH = sourceH * (1 - cropT - cropB);
        if (tileW <= 0 || tileH <= 0) return `${clipDef}${border}<g${transform}>${textOverlay}</g>`;
        const alignment = layout.alignment ?? 'tl';
        const horizontal = ['t', 'ctr', 'b'].includes(alignment)
          ? 0.5
          : ['tr', 'r', 'br'].includes(alignment)
            ? 1
            : 0;
        const vertical = ['l', 'ctr', 'r'].includes(alignment)
          ? 0.5
          : ['bl', 'b', 'br'].includes(alignment)
            ? 1
            : 0;
        const tileX = x + (w - tileW) * horizontal + (layout.offsetX ?? 0);
        const tileY = y + (h - tileH) * vertical + (layout.offsetY ?? 0);
        const mirrorX = layout.flip === 'x' || layout.flip === 'xy';
        const mirrorY = layout.flip === 'y' || layout.flip === 'xy';
        const patternId = mintId();
        const images: string[] = [];
        for (let row = 0; row < (mirrorY ? 2 : 1); row++) {
          for (let col = 0; col < (mirrorX ? 2 : 1); col++) {
            const reflection =
              col || row
                ? ` transform="translate(${E(col * 2 * tileW)} ${E(row * 2 * tileH)}) scale(${col ? -1 : 1} ${row ? -1 : 1})"`
                : '';
            images.push(
              `<g${reflection}><svg width="${E(tileW)}" height="${E(tileH)}" overflow="hidden"><image x="${E(-sourceW * cropL)}" y="${E(-sourceH * cropT)}" width="${E(sourceW)}" height="${E(sourceH)}" href="${dataUrl}" xlink:href="${dataUrl}" preserveAspectRatio="none"/></svg></g>`,
            );
          }
        }
        const pattern = `<defs><pattern id="${patternId}" patternUnits="userSpaceOnUse" x="${E(tileX)}" y="${E(tileY)}" width="${E(tileW * (mirrorX ? 2 : 1))}" height="${E(tileH * (mirrorY ? 2 : 1))}">${images.join('')}</pattern></defs>`;
        const geometry = pictureClipGeometry(shape, preset, x, y, w, h);
        return `${clipDef}${pattern}<g${transform} fill="url(#${patternId})"${filterAttr}${opacityAttr}>${geometry}</g>${border}<g${transform}>${textOverlay}</g>`;
      }
    }
    return `${clipDef}<g${transform}${clipAttr}><image x="${E(imgX)}" y="${E(imgY)}" width="${E(imgW)}" height="${E(imgH)}" href="${dataUrl}" xlink:href="${dataUrl}" preserveAspectRatio="none"${filterAttr}${opacityAttr}/></g>${border}<g${transform}>${textOverlay}</g>`;
  }
  // B14 — external r:link pictures don't ship bytes in the package.
  // Surface the URL in the placeholder so users can see where the
  // picture lives.
  const linkUrl = getShapeImageLinkUrl(shape);
  const label = !bytes
    ? linkUrl
      ? `picture (link: ${linkUrl.length > 48 ? linkUrl.slice(0, 45) + '…' : linkUrl})`
      : 'picture (no bytes)'
    : `picture (${format ?? 'unknown'}${bytes ? `, ${bytes.byteLength} B` : ''})`;
  return `<g data-pptx-fallback="image"${transform}><rect x="${E(x)}" y="${E(y)}" width="${E(w)}" height="${E(h)}" fill="#F3F4F6" stroke="#9CA3AF" stroke-width="${E(9_525)}" stroke-dasharray="${E(50_000)},${E(30_000)}"/>${renderPicturePlaceholderLabel(x, y, w, h, label)}${textOverlay}</g>`;
};

// Share preset path generators with native shapes so image masks use the
// same slide coordinates and are transformed together with the cropped image.
const pictureClipGeometry = (
  shape: SlideShapeData,
  preset: string,
  x: number,
  y: number,
  w: number,
  h: number,
): string => {
  if (preset === 'ellipse')
    return `<ellipse cx="${E(x + w / 2)}" cy="${E(y + h / 2)}" rx="${E(w / 2)}" ry="${E(h / 2)}"/>`;
  if (preset === 'roundRect') {
    const radius = E(
      Math.min(w, h) *
        Math.max(0, Math.min(0.5, (getShapeAdjustValues(shape).adj ?? 16667) / 100000)),
    );
    return `<rect x="${E(x)}" y="${E(y)}" width="${E(w)}" height="${E(h)}" rx="${radius}" ry="${radius}"/>`;
  }
  const path = PRESET_PATHS[preset];
  if (path)
    return `<path d="${path(x / EMU_PER_PX, y / EMU_PER_PX, w / EMU_PER_PX, h / EMU_PER_PX)}" clip-rule="evenodd"/>`;
  const points = PRESET_POINTS[preset];
  if (points)
    return `<polygon points="${points(w / EMU_PER_PX, h / EMU_PER_PX)
      .map(([nx, ny]) => `${E(x + nx * w)},${E(y + ny * h)}`)
      .join(' ')}"/>`;
  return `<rect x="${E(x)}" y="${E(y)}" width="${E(w)}" height="${E(h)}"/>`;
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

// Parse `#RRGGBB` (or bare `RRGGBB`) into 0-255 channels.
const hexChannels = (hex: string): [number, number, number] => {
  const h = hex.startsWith('#') ? hex.slice(1) : hex;
  return [
    Number.parseInt(h.slice(0, 2), 16) || 0,
    Number.parseInt(h.slice(2, 4), 16) || 0,
    Number.parseInt(h.slice(4, 6), 16) || 0,
  ];
};

// Linearly mix two #RRGGBB colors. `t` = weight of `a` in [0,1].
const mixHex = (aHex: string, bHex: string, t: number): string => {
  const aa = aHex.startsWith('#') ? aHex.slice(1) : aHex;
  const bb = bHex.startsWith('#') ? bHex.slice(1) : bHex;
  const part = (h: string, off: number): number => Number.parseInt(h.slice(off, off + 2), 16);
  const r = Math.round(part(aa, 0) * t + part(bb, 0) * (1 - t));
  const g = Math.round(part(aa, 2) * t + part(bb, 2) * (1 - t));
  const b = Math.round(part(aa, 4) * t + part(bb, 4) * (1 - t));
  const h = (n: number): string =>
    Math.max(0, Math.min(255, n)).toString(16).padStart(2, '0').toUpperCase();
  return `#${h(r)}${h(g)}${h(b)}`;
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
  // Accept either `scheme:tx1` (the form fills + strokes use) or a bare
  // ECMA-376 token (`tx1`, `accent1`, ...) from the run-format cascade.
  let token: string | null = null;
  if (c.startsWith('scheme:')) token = c.slice('scheme:'.length);
  else if (SCHEME_TO_THEME[c]) token = c;
  if (token !== null) {
    if (theme) {
      // Remap the slide token through the effective color map first, then index
      // the theme — so an inverted map (tx1→lt1) paints what PowerPoint paints.
      const mapped = activeColorMap?.[token] ?? token;
      const key = SCHEME_TO_THEME[mapped] ?? SCHEME_TO_THEME[token];
      if (key) return normalizeHex(theme[key]);
    }
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

// Effective color map (`<p:clrMap>` overlaid by `<p:clrMapOvr>`) for the slide
// currently rendering. PowerPoint remaps the slide tokens (`tx1`, `bg1`, …) to
// theme slots through this map before the theme lookup; an inverted map
// (`bg1="dk1" tx1="lt1"`) is common in Google-Slides / Canva exports and makes
// `tx1` resolve to the light slot. `renderSlideSvg` sets this at the start of
// each (synchronous, single-slide) render and `resolveColor` reads it. Like
// `nextDefId` above, every entry overwrites it, so it needs no reset.
let activeColorMap: Readonly<Record<string, string>> | null = null;

// Default body-text color for the slide currently rendering. PowerPoint takes
// the fallback text color from the master's `bodyStyle` (e.g. `schemeClr bg1`),
// not from the `tx1` token — and in an inverted color map `tx1` resolves to the
// background slot, which would paint default text the same color as the surface
// (white-on-white, an invisible slide). renderSlideSvg sets this from
// `resolveDeckBodyTextColor`; overwritten on every entry, so no reset needed.
let activeDeckTextColor = '#000000';

// The number a `slidenum` field shows on the slide being rendered: the deck's
// `firstSlideNum` plus the slide's position. Set on every renderSlideSvg entry
// alongside the color map, for the same reason — the text path is several
// calls deep and the slide is not one of its arguments.
let activeSlideNumber = '1';

// `<linearGradient>` definition + `fill="url(#…)"` reference, projected
// from @office-kit/pptx's `{ stops, angleDeg }` shape onto SVG's
// objectBoundingBox unit cube. ECMA-376 measures `angleDeg` clockwise
// from 3 o'clock, which matches the trig below (0° = +x, 90° = +y).
const gradientDef = (
  grad: ReadGradientFill,
  theme: PresentationTheme | null,
  transform = '',
): { defs: string; fillAttr: string } => {
  const id = mintId();
  const orderedStops = [...grad.stops].sort((a, b) => a.offset - b.offset);
  const stops = orderedStops
    .map(
      (s) =>
        `<stop offset="${s.offset.toFixed(4)}" stop-color="${s.resolvedColor ?? resolveColor(s.color, theme, '#E5E7EB')}"${s.opacity !== undefined ? ` stop-opacity="${s.opacity}"` : ''}/>`,
    )
    .join('');
  if (grad.path === 'rect') {
    const focus = grad.focus ?? { left: 0.5, top: 0.5, right: 0.5, bottom: 0.5 };
    const tile = grad.tileRect ?? { left: 0, top: 0, right: 0, bottom: 0 };
    const l = tile.left;
    const t = tile.top;
    const r = 1 - tile.right;
    const b = 1 - tile.bottom;
    const fl = focus.left;
    const ft = focus.top;
    const fr = 1 - focus.right;
    const fb = 1 - focus.bottom;
    // Each side interpolates perpendicular to its edge. Together the four
    // trapezoids produce rectangular contours, including PowerPoint's corner
    // presets whose tiles extend beyond the shape. Inverted imported rectangles
    // keep the radial fallback below.
    if (l <= fl && fl <= fr && fr <= r && t <= ft && ft <= fb && fb <= b && l < r && t < b) {
      const sides = [
        { points: `${l},${t} ${r},${t} ${fr},${ft} ${fl},${ft}`, x1: l, y1: t, x2: l, y2: ft },
        { points: `${r},${t} ${r},${b} ${fr},${fb} ${fr},${ft}`, x1: r, y1: t, x2: fr, y2: t },
        { points: `${r},${b} ${l},${b} ${fl},${fb} ${fr},${fb}`, x1: l, y1: b, x2: l, y2: fb },
        { points: `${l},${b} ${l},${t} ${fl},${ft} ${fl},${fb}`, x1: l, y1: t, x2: fl, y2: t },
      ];
      let gradients = '';
      let polygons = '';
      for (const side of sides) {
        if (side.x1 === side.x2 && side.y1 === side.y2) continue;
        const sideId = mintId();
        gradients += `<linearGradient id="${sideId}" gradientUnits="userSpaceOnUse" x1="${side.x2}" y1="${side.y2}" x2="${side.x1}" y2="${side.y1}">${stops}</linearGradient>`;
        polygons += `<polygon points="${side.points}" fill="url(#${sideId})"/>`;
      }
      const first = orderedStops[0];
      const center =
        first && fr > fl && fb > ft
          ? `<rect x="${fl}" y="${ft}" width="${fr - fl}" height="${fb - ft}" fill="${first.resolvedColor ?? resolveColor(first.color, theme, '#E5E7EB')}"${first.opacity !== undefined ? ` fill-opacity="${first.opacity}"` : ''}/>`
          : '';
      // Antialiasing adjacent transparent polygons separately leaves hairline
      // gaps. Their colors meet continuously, so rasterize only their shared
      // boundaries without antialiasing; the shape's outer clip stays smooth.
      const defs = `<defs>${gradients}<pattern id="${id}" patternUnits="objectBoundingBox" patternContentUnits="objectBoundingBox" x="${l}" y="${t}" width="${r - l}" height="${b - t}"${transform.replace('gradientTransform', 'patternTransform')}><g shape-rendering="crispEdges" transform="translate(${-l} ${-t})">${polygons}${center}</g></pattern></defs>`;
      return { defs, fillAttr: `url(#${id})` };
    }
  }
  if (grad.path === 'circle' || grad.path === 'rect' || grad.path === 'shape') {
    // Shape-following paths still use an elliptical approximation.
    const focus = grad.focus ?? { left: 0.5, top: 0.5, right: 0.5, bottom: 0.5 };
    // fillToRect describes insets from each edge, not absolute coordinates.
    // PowerPoint's bottom-right focus has l=t=1 and r=b=0.
    const cx = (focus.left + 1 - focus.right) / 2;
    const cy = (focus.top + 1 - focus.bottom) / 2;
    // Mac PowerPoint places the first stop at the focus, as SVG does.
    const defs = `<defs><radialGradient id="${id}" gradientUnits="objectBoundingBox" cx="${cx.toFixed(4)}" cy="${cy.toFixed(4)}" r="${Math.max(0.5, Math.max(cx, cy, 1 - cx, 1 - cy)).toFixed(4)}">${stops}</radialGradient></defs>`;
    return { defs, fillAttr: `url(#${id})` };
  }
  const angleRad = ((grad.angleDeg ?? 0) * Math.PI) / 180;
  const dx = Math.cos(angleRad) / 2;
  const dy = Math.sin(angleRad) / 2;
  const x1 = 0.5 - dx;
  const y1 = 0.5 - dy;
  const x2 = 0.5 + dx;
  const y2 = 0.5 + dy;
  const defs = `<defs><linearGradient id="${id}" gradientUnits="objectBoundingBox"${transform} x1="${x1.toFixed(4)}" y1="${y1.toFixed(4)}" x2="${x2.toFixed(4)}" y2="${y2.toFixed(4)}">${stops}</linearGradient></defs>`;
  return { defs, fillAttr: `url(#${id})` };
};

// SVG `<pattern>` definitions for the ECMA-376 ST_PresetPatternVal
// presets (§20.1.10.49), tuned against LibreOffice's own substitution for
// these fills (not GDI's original hatch-brush metrics — LibreOffice draws
// none of them as literal bitmaps). Two calibrated base pitches (SVG
// user-space px at 96 DPI) cover the whole family: orthogonal
// (horizontal/vertical) hatches read far finer than diagonal ones at the
// same nominal weight, and "lg"-prefixed grids use a materially larger
// cell than "sm"/plain ones. Unknown presets fall through to a 50%-style
// diagonal crosshatch.
const ORTHO_TILE = 4;
const ORTHO_WIDE_TILE = 16;
const DIAG_TILE = 16;
const MOTIF_TILE = 8; // wave / weave / sphere / diamond motifs are sized to an 8-unit cell
const PATTERN_TILE_SIZE: Record<string, number> = {
  horz: ORTHO_TILE,
  ltHorz: ORTHO_TILE,
  narHorz: ORTHO_TILE,
  dashHorz: ORTHO_TILE,
  horzBrick: ORTHO_TILE,
  dkHorz: ORTHO_TILE,
  vert: ORTHO_TILE,
  ltVert: ORTHO_TILE,
  narVert: ORTHO_TILE,
  dashVert: ORTHO_TILE,
  dkVert: ORTHO_TILE,
  ltHorzCross: ORTHO_TILE,
  cross: ORTHO_TILE,
  dotGrid: ORTHO_TILE,
  smGrid: ORTHO_TILE,
  dkHorzCross: ORTHO_WIDE_TILE,
  lgGrid: ORTHO_WIDE_TILE,
  plaid: ORTHO_WIDE_TILE,
  wave: MOTIF_TILE,
  zigZag: MOTIF_TILE,
  weave: MOTIF_TILE,
  divot: MOTIF_TILE,
  sphere: MOTIF_TILE,
  solidDmnd: MOTIF_TILE,
  openDmnd: MOTIF_TILE,
};
const patternDef = (pat: {
  preset: string;
  foreground: string;
  background: string;
}): { defs: string; fillAttr: string } => {
  const id = mintId();
  const fg = pat.foreground;
  const bg = pat.background;
  const preset = pat.preset;
  let body = '';
  // pct{N} — LibreOffice substitutes a diagonal hatch for these (not a literal
  // N%-coverage bitmap): a sparse single-direction hatch below ~30% density, a full
  // crosshatch at and above it, with the pitch shrinking as density rises. The
  // pitch formula is calibrated so pct5 (at the floor) comes out to
  // `DIAG_TILE * 2`; PCT_DENSITY_FLOOR must stay the single reference density
  // for both the clamp and the ratio below it or that calibration drifts.
  const PCT_DENSITY_FLOOR = 0.05;
  const PCT_TILE_MIN = 6;
  const pctMatch = /^pct(\d+)$/.exec(preset);
  const pctDensity = pctMatch
    ? Math.min(100, Math.max(0, Number.parseInt(pctMatch[1]!, 10))) / 100
    : null;
  const W =
    pctDensity !== null
      ? Math.max(
          PCT_TILE_MIN,
          DIAG_TILE * 2 * Math.sqrt(PCT_DENSITY_FLOOR / Math.max(pctDensity, PCT_DENSITY_FLOOR)),
        )
      : (PATTERN_TILE_SIZE[preset] ?? DIAG_TILE);
  const H = W;
  const stripe = (orientation: 'h' | 'v' | 'd' | 'a', width = 1): string => {
    if (orientation === 'h')
      return `<path d="M0 ${H / 2}H${W}" stroke="${fg}" stroke-width="${width}"/>`;
    if (orientation === 'v')
      return `<path d="M${W / 2} 0V${H}" stroke="${fg}" stroke-width="${width}"/>`;
    if (orientation === 'd')
      return `<path d="M0 0L${W} ${H}" stroke="${fg}" stroke-width="${width}"/>`;
    return `<path d="M${W} 0L0 ${H}" stroke="${fg}" stroke-width="${width}"/>`;
  };
  if (pctDensity !== null) {
    body = pctDensity >= 0.3 ? stripe('d', 0.8) + stripe('a', 0.8) : stripe('d', 0.8);
  } else if (
    preset === 'horz' ||
    preset === 'ltHorz' ||
    preset === 'narHorz' ||
    preset === 'dashHorz' ||
    preset === 'horzBrick'
  ) {
    // ST_PresetPatternVal horizontal family (§20.1.10.49) — horizontal lines.
    // (The old branch keyed on GDI HatchStyle names like 'ltHorizontal', which
    // no valid OOXML emits, so 'horz'/'ltHorz' fell through to the 50% checker.)
    body = stripe('h', 0.8);
  } else if (preset === 'dkHorz') {
    body = stripe('h', 2);
  } else if (
    preset === 'vert' ||
    preset === 'ltVert' ||
    preset === 'narVert' ||
    preset === 'dashVert'
  ) {
    body = stripe('v', 0.8);
  } else if (preset === 'dkVert') {
    body = stripe('v', 2);
  } else if (preset === 'ltUpDiag' || preset === 'wdUpDiag') {
    body = stripe('d', 0.8);
  } else if (preset === 'dkUpDiag') {
    body = stripe('d', 2);
  } else if (preset === 'ltDnDiag' || preset === 'wdDnDiag') {
    body = stripe('a', 0.8);
  } else if (preset === 'dkDnDiag') {
    body = stripe('a', 2);
  } else if (
    preset === 'ltHorzCross' ||
    preset === 'smGrid' ||
    preset === 'cross' ||
    preset === 'dotGrid'
  ) {
    // dotGrid is a dotted grid in the spec; LibreOffice renders it as a thin
    // line grid, so a light cross-grid is the closest match to the ground truth.
    body = stripe('h', 0.8) + stripe('v', 0.8);
  } else if (preset === 'dkHorzCross' || preset === 'lgGrid' || preset === 'plaid') {
    body = stripe('h', 2) + stripe('v', 2);
  } else if (
    preset === 'diagCross' ||
    preset === 'trellis' ||
    preset === 'shingle' ||
    preset === 'dashUpDiag' ||
    preset === 'dashDnDiag'
  ) {
    body = stripe('d', 0.8) + stripe('a', 0.8);
  } else if (preset === 'dkUpDiagStripe' || preset === 'dkDnDiagStripe') {
    body = stripe(preset === 'dkUpDiagStripe' ? 'd' : 'a', 2);
  } else if (preset === 'wave' || preset === 'zigZag') {
    body = `<path d="M0 4Q2 2 4 4T8 4" stroke="${fg}" stroke-width="0.8" fill="none"/>`;
  } else if (preset === 'weave' || preset === 'divot') {
    body = `<path d="M0 0L4 4 0 8M4 0L8 4 4 8" stroke="${fg}" stroke-width="0.8" fill="none"/>`;
  } else if (preset === 'sphere') {
    body = `<circle cx="4" cy="4" r="3" fill="${fg}" fill-opacity="0.7"/>`;
  } else if (preset === 'solidDmnd' || preset === 'openDmnd') {
    body = `<path d="M4 1L7 4 4 7 1 4Z" fill="${preset === 'solidDmnd' ? fg : 'none'}" stroke="${fg}" stroke-width="0.6"/>`;
  } else {
    // Unrecognized preset — fall back to a pct50-style diagonal crosshatch.
    body = stripe('d', 0.8) + stripe('a', 0.8);
  }
  const defs = `<defs><pattern id="${id}" patternUnits="userSpaceOnUse" width="${W}" height="${H}"><rect width="${W}" height="${H}" fill="${bg}"/>${body}</pattern></defs>`;
  return { defs, fillAttr: `url(#${id})` };
};

interface PaintResult {
  fill: string;
  /** Pre-built `fill-opacity` attribute for a translucent solid fill, else ''. */
  fillAttrs: string;
  stroke: string;
  strokeWidth: number;
  /** Extra SVG `<defs>` the caller should emit before the shape. */
  defs: string;
  /** Pre-built stroke style attributes (dasharray, linecap, linejoin, etc.). */
  strokeAttrs: string;
  /** Pre-built marker-start / marker-end attributes for arrowheads. */
  markerAttrs: string;
}

// ECMA-376 `<a:prstDash val="…"/>` → SVG `stroke-dasharray` (in stroke widths).
// PowerPoint scales the pattern by the line width; the multipliers below were
// reverse-engineered from real PPTX exports so the visual cadence matches.
const DASH_PATTERNS: Record<string, string> = {
  solid: '',
  dot: '1 3',
  dash: '4 3',
  lgDash: '8 3',
  dashDot: '4 3 1 3',
  lgDashDot: '8 3 1 3',
  lgDashDotDot: '8 3 1 3 1 3',
  sysDash: '3 1',
  sysDot: '1 1',
  sysDashDot: '3 1 1 1',
  sysDashDotDot: '3 1 1 1 1 1',
};

const arrowSize = (size: 'sm' | 'med' | 'lg' | undefined): number =>
  size === 'sm' ? 3 : size === 'lg' ? 7 : 5;

const buildArrowMarker = (
  type: string,
  width: 'sm' | 'med' | 'lg' | undefined,
  length: 'sm' | 'med' | 'lg' | undefined,
  color: string,
  orient: 'auto' | 'auto-start-reverse',
): { id: string; def: string } => {
  const id = mintId();
  const w = arrowSize(width);
  const h = arrowSize(length);
  // refX = w so the tip lands on the line's endpoint; refY centers vertically.
  let body: string;
  switch (type) {
    case 'triangle':
    case 'arrow':
      body = `<path d="M0 0L${w} ${h / 2}L0 ${h}z" fill="${color}"/>`;
      break;
    case 'stealth':
      body = `<path d="M0 0L${w} ${h / 2}L0 ${h}L${w * 0.4} ${h / 2}z" fill="${color}"/>`;
      break;
    case 'diamond':
      body = `<path d="M0 ${h / 2}L${w / 2} 0L${w} ${h / 2}L${w / 2} ${h}z" fill="${color}"/>`;
      break;
    case 'oval':
      body = `<ellipse cx="${w / 2}" cy="${h / 2}" rx="${w / 2}" ry="${h / 2}" fill="${color}"/>`;
      break;
    case 'none':
    default:
      body = '';
  }
  const def = `<defs><marker id="${id}" viewBox="0 0 ${w} ${h}" refX="${w}" refY="${h / 2}" markerWidth="${w}" markerHeight="${h}" orient="${orient}">${body}</marker></defs>`;
  return { id, def };
};

const paint = (
  shape: SlideShapeData | null,
  fill: ShapeFill,
  stroke: ShapeStroke,
  theme: PresentationTheme | null,
  isPlaceholder: boolean,
  pres?: PresentationData,
): PaintResult => {
  let fillColor: string;
  let fillAttrs = '';
  let defs = '';
  if (fill.kind === 'solid') {
    // Prefer the transform-aware reader when we have a presentation
    // handle — `<a:lumMod>` / `<a:shade>` / `<a:tint>` etc. on the
    // shape's fill produce the actual painted color, while `fill.color`
    // is only the base value. Falls through to the legacy path when
    // the resolver can't produce a hex (theme missing, etc.).
    let resolved: string | null = null;
    if (shape && pres) resolved = getShapeFillColorResolved(pres, shape);
    fillColor = resolved ?? resolveColor(fill.color, theme, '#E5E7EB');
    // `<a:alpha>` lives beside the color, not in it — a 27% "veil" over
    // layout artwork must not paint as an opaque block.
    const opacity = shape ? getShapeFillOpacity(shape) : null;
    if (opacity !== null && opacity < 1) fillAttrs = ` fill-opacity="${opacity.toFixed(3)}"`;
  } else if (fill.kind === 'none') {
    fillColor = 'none';
  } else if (fill.kind === 'gradient') {
    // Resolve the gradient through the shape → layout → master
    // placeholder cascade so inherited gradFills paint a real
    // `<linearGradient>` / `<radialGradient>`, not a placeholder color.
    const grad = shape
      ? pres
        ? getShapeGradientFillEffective(pres, shape)
        : getShapeGradientFill(shape)
      : null;
    if (grad) {
      let transform = '';
      const bounds = shape && (pres ? getShapeBoundsResolved(pres, shape) : getShapeBounds(shape));
      if (grad.rotateWithShape === false && shape && bounds && bounds.w > 0 && bounds.h > 0) {
        // Mac PowerPoint anchors a non-rotating gradient to the rotated shape's
        // axis-aligned bounding box. Undo the shape transform in physical space;
        // rotating the unit square alone distorts wide or tall shapes.
        const angle = (getShapeRotation(shape) * Math.PI) / 180;
        const cos = Math.cos(angle);
        const sin = Math.sin(angle);
        const w = bounds.w;
        const h = bounds.h;
        const rotatedW = Math.abs(cos) * w + Math.abs(sin) * h;
        const rotatedH = Math.abs(sin) * w + Math.abs(cos) * h;
        const flip = getShapeFlip(shape);
        const sx = flip?.horizontal ? -1 : 1;
        const sy = flip?.vertical ? -1 : 1;
        const a = (sx * cos * rotatedW) / w;
        const b = (-sy * sin * rotatedW) / h;
        const c = (sx * sin * rotatedH) / w;
        const d = (sy * cos * rotatedH) / h;
        const matrix = [a, b, c, d, (1 - a - c) / 2, (1 - b - d) / 2];
        transform = ` gradientTransform="matrix(${matrix.map((value) => value.toFixed(6)).join(' ')})"`;
      }
      const built = gradientDef(grad, theme, transform);
      defs = built.defs;
      fillColor = built.fillAttr;
    } else {
      // The fill reads as a gradient but no `<a:gradFill>` is reachable
      // (e.g. it's referenced through the theme style matrix, which the
      // core doesn't model yet). Fall back to the resolved solid color
      // when the shape carries one, else transparent — never a spurious
      // tint that masks real content.
      fillColor = (shape && pres ? getShapeFillColorResolved(pres, shape) : null) ?? 'none';
    }
  } else if (fill.kind === 'pattern') {
    const pat = shape && pres ? getShapePatternFill(pres, shape) : null;
    if (pat) {
      const built = patternDef(pat);
      defs = built.defs;
      fillColor = built.fillAttr;
    } else {
      fillColor = '#BFDBFE';
    }
  } else if (fill.kind === 'image') {
    fillColor = '#DDD6FE';
  } else {
    // `inherit`: the real fill lives in the layout/master cascade or the
    // shape's style matrix (`<p:style><a:fillRef>`), which we don't resolve
    // yet. Render transparent — that matches what PowerPoint / LibreOffice
    // show for unstyled text boxes, content placeholders, and bare autoshapes
    // (verified against ground truth), and never paints a spurious grey box
    // over real content. (`isPlaceholder` no longer changes the fill: a
    // no-type `<p:ph>` reads as phType null, so keying on it mislabelled
    // content placeholders as non-placeholders anyway.)
    void isPlaceholder;
    fillColor = 'none';
  }

  let strokeColor = 'none';
  let strokeWidth = 0;
  const strokeAttrParts: string[] = [];
  let markerAttrs = '';
  if (stroke.kind === 'solid') {
    let resolved: string | null = null;
    if (shape && pres) resolved = getShapeStrokeColorResolved(pres, shape);
    strokeColor = resolved ?? resolveColor(stroke.color, theme, '#9CA3AF');
    strokeWidth = stroke.widthEmu ?? 9_525; // 1pt
    if (shape) {
      const opacity = getShapeStrokeOpacity(shape);
      if (opacity !== null && opacity < 1) {
        strokeAttrParts.push(`stroke-opacity="${opacity.toFixed(3)}"`);
      }
      const dash = getShapeStrokeDash(shape);
      if (dash && dash !== 'solid') {
        const pattern = DASH_PATTERNS[dash];
        if (pattern) {
          // Scale by stroke width in CSS pixels so the cadence matches PowerPoint.
          const swPx = strokeWidth / EMU_PER_PX;
          const arr = pattern
            .split(' ')
            .map((n) => (Number.parseFloat(n) * swPx).toFixed(2))
            .join(' ');
          strokeAttrParts.push(`stroke-dasharray="${arr}"`);
        }
      }
      const cap = getShapeStrokeCap(shape);
      if (cap === 'rnd') strokeAttrParts.push('stroke-linecap="round"');
      else if (cap === 'sq') strokeAttrParts.push('stroke-linecap="square"');
      else if (cap === 'flat') strokeAttrParts.push('stroke-linecap="butt"');
      const join = getShapeStrokeJoin(shape);
      if (join === 'round') strokeAttrParts.push('stroke-linejoin="round"');
      else if (join === 'bevel') strokeAttrParts.push('stroke-linejoin="bevel"');
      else if (join === 'miter') strokeAttrParts.push('stroke-linejoin="miter"');
      const cmpd = getShapeStrokeCompound(shape);
      if (cmpd === 'dbl') {
        // Approximate a double line by widening + a transparent stripe down
        // the middle. SVG has no native compound-line primitive.
        strokeWidth = Math.max(strokeWidth, 19_050);
      }
      const head = getShapeStrokeArrow(shape, 'head');
      const tail = getShapeStrokeArrow(shape, 'tail');
      if (head && head.type !== 'none') {
        const m = buildArrowMarker(
          head.type,
          head.width,
          head.length,
          strokeColor,
          'auto-start-reverse',
        );
        defs += m.def;
        markerAttrs += ` marker-start="url(#${m.id})"`;
      }
      if (tail && tail.type !== 'none') {
        const m = buildArrowMarker(tail.type, tail.width, tail.length, strokeColor, 'auto');
        defs += m.def;
        markerAttrs += ` marker-end="url(#${m.id})"`;
      }
    }
  }
  return {
    fill: fillColor,
    fillAttrs,
    stroke: strokeColor,
    strokeWidth,
    defs,
    strokeAttrs: strokeAttrParts.join(' '),
    markerAttrs,
  };
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

// Preset polygons in normalized [0,1] space. Most ignore the shape size, but
// the block-arrow presets need it: in OOXML their arrowhead length is a
// fraction of `min(w,h)` (not of the side it points along), so a fixed-fraction
// polygon is wrong for any non-square arrow. `w`/`h` are the shape's px size;
// only their ratio is used, so the unit cancels.
const PRESET_POINTS: Record<string, (w: number, h: number) => Array<[number, number]>> = {
  triangle: () => [
    [0.5, 0],
    [1, 1],
    [0, 1],
  ],
  rtTriangle: () => [
    [0, 0],
    [1, 1],
    [0, 1],
  ],
  diamond: () => [
    [0.5, 0],
    [1, 0.5],
    [0.5, 1],
    [0, 0.5],
  ],
  parallelogram: () => [
    [0.25, 0],
    [1, 0],
    [0.75, 1],
    [0, 1],
  ],
  trapezoid: () => [
    [0.25, 0],
    [0.75, 0],
    [1, 1],
    [0, 1],
  ],
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
  // Block arrows: default adj1=adj2=50000 → shaft is 0.5·(cross dimension),
  // arrowhead length is 0.5·min(w,h) along the pointing axis (ECMA-376
  // rightArrow et al.). Shaft spans 0.25..0.75 of the cross dimension.
  rightArrow: (w, h) => {
    const bx = 1 - Math.min(1, (0.5 * Math.min(w, h)) / w);
    return [
      [0, 0.25],
      [bx, 0.25],
      [bx, 0],
      [1, 0.5],
      [bx, 1],
      [bx, 0.75],
      [0, 0.75],
    ];
  },
  leftArrow: (w, h) => {
    const bx = Math.min(1, (0.5 * Math.min(w, h)) / w);
    return [
      [1, 0.25],
      [bx, 0.25],
      [bx, 0],
      [0, 0.5],
      [bx, 1],
      [bx, 0.75],
      [1, 0.75],
    ];
  },
  upArrow: (w, h) => {
    const by = Math.min(1, (0.5 * Math.min(w, h)) / h);
    return [
      [0.25, 1],
      [0.25, by],
      [0, by],
      [0.5, 0],
      [1, by],
      [0.75, by],
      [0.75, 1],
    ];
  },
  downArrow: (w, h) => {
    const by = 1 - Math.min(1, (0.5 * Math.min(w, h)) / h);
    return [
      [0.25, 0],
      [0.25, by],
      [0, by],
      [0.5, 1],
      [1, by],
      [0.75, by],
      [0.75, 0],
    ];
  },
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
      const a = (i / (rays * 2)) * 2 * Math.PI - Math.PI / 2;
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

  // -- Flowchart shapes ---------------------------------------------------
  // Lightweight approximations of the ~28 ECMA-376 flowchart presets.
  // They're laid out so the shape's bounding box matches the slide's,
  // and the geometry is what most viewers expect at a glance.
  flowChartProcess: (x, y, w, h) => `M${x},${y} L${x + w},${y} L${x + w},${y + h} L${x},${y + h} Z`,
  flowChartAlternateProcess: (x, y, w, h) => {
    const r = Math.min(w, h) * 0.18;
    return `M${x + r},${y} L${x + w - r},${y} A${r},${r} 0 0 1 ${x + w},${y + r} L${x + w},${y + h - r} A${r},${r} 0 0 1 ${x + w - r},${y + h} L${x + r},${y + h} A${r},${r} 0 0 1 ${x},${y + h - r} L${x},${y + r} A${r},${r} 0 0 1 ${x + r},${y} Z`;
  },
  flowChartDecision: (x, y, w, h) => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    return `M${cx},${y} L${x + w},${cy} L${cx},${y + h} L${x},${cy} Z`;
  },
  flowChartTerminator: (x, y, w, h) => {
    const r = h / 2;
    return `M${x + r},${y} L${x + w - r},${y} A${r},${r} 0 0 1 ${x + w - r},${y + h} L${x + r},${y + h} A${r},${r} 0 0 1 ${x + r},${y} Z`;
  },
  flowChartConnector: (x, y, w, h) => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const r = Math.min(w, h) / 2;
    return `M${cx + r},${cy} A${r},${r} 0 1 1 ${cx - r},${cy} A${r},${r} 0 1 1 ${cx + r},${cy} Z`;
  },
  flowChartDocument: (x, y, w, h) => {
    // Bottom is a wave (cubic) instead of a flat line.
    const wave = h * 0.18;
    return `M${x},${y} L${x + w},${y} L${x + w},${y + h - wave} C${x + w * 0.75},${y + h + wave * 0.5} ${x + w * 0.25},${y + h - wave * 2} ${x},${y + h - wave * 0.5} Z`;
  },
  flowChartMultidocument: (x, y, w, h) => {
    // Two stacked documents (back document offset by 6%).
    const inset = w * 0.06;
    const back = `M${x + inset},${y + inset * 0.6} L${x + w},${y + inset * 0.6} L${x + w},${y + h - inset * 0.6} L${x + w - inset},${y + h - inset * 0.6} L${x + w - inset},${y + inset * 0.6}`;
    const front = `M${x},${y + inset * 1.2} L${x + w - inset},${y + inset * 1.2} L${x + w - inset},${y + h * 0.85} C${x + (w - inset) * 0.75},${y + h + 6} ${x + (w - inset) * 0.25},${y + h * 0.75} ${x},${y + h * 0.95} Z`;
    return `${back} Z ${front}`;
  },
  flowChartPredefinedProcess: (x, y, w, h) => {
    // Process box with two vertical bars carved on each side.
    const inset = w * 0.1;
    return `M${x},${y} L${x + w},${y} L${x + w},${y + h} L${x},${y + h} Z M${x + inset},${y} L${x + inset},${y + h} M${x + w - inset},${y} L${x + w - inset},${y + h}`;
  },
  flowChartInternalStorage: (x, y, w, h) => {
    const inset = Math.min(w, h) * 0.1;
    return `M${x},${y} L${x + w},${y} L${x + w},${y + h} L${x},${y + h} Z M${x + inset},${y} L${x + inset},${y + h} M${x},${y + inset} L${x + w},${y + inset}`;
  },
  flowChartManualInput: (x, y, w, h) => {
    return `M${x},${y + h * 0.35} L${x + w},${y} L${x + w},${y + h} L${x},${y + h} Z`;
  },
  flowChartManualOperation: (x, y, w, h) => {
    return `M${x},${y} L${x + w},${y} L${x + w * 0.8},${y + h} L${x + w * 0.2},${y + h} Z`;
  },
  flowChartInputOutput: (x, y, w, h) => {
    const skew = w * 0.18;
    return `M${x + skew},${y} L${x + w},${y} L${x + w - skew},${y + h} L${x},${y + h} Z`;
  },
  flowChartPunchedTape: (x, y, w, h) => {
    const wave = h * 0.12;
    return `M${x},${y + wave} C${x + w * 0.25},${y - wave} ${x + w * 0.75},${y + wave * 2} ${x + w},${y + wave} L${x + w},${y + h - wave} C${x + w * 0.75},${y + h + wave} ${x + w * 0.25},${y + h - wave * 2} ${x},${y + h - wave} Z`;
  },
  flowChartCard: (x, y, w, h) => {
    const cut = h * 0.3;
    return `M${x + cut},${y} L${x + w},${y} L${x + w},${y + h} L${x},${y + h} L${x},${y + cut} Z`;
  },
  flowChartPunchedCard: (x, y, w, h) => {
    const cut = h * 0.2;
    return `M${x + cut},${y} L${x + w},${y} L${x + w},${y + h} L${x},${y + h} L${x},${y + cut} Z`;
  },
  flowChartOnlineStorage: (x, y, w, h) => {
    // Tape-like: ellipse-capped rectangle, left side open.
    const cap = w * 0.12;
    return `M${x + cap},${y} L${x + w},${y} L${x + w},${y + h} L${x + cap},${y + h} A${cap},${h / 2} 0 0 1 ${x + cap},${y} Z`;
  },
  flowChartMagneticDisk: (x, y, w, h) => {
    // Cylinder: rectangle with ellipses top and bottom.
    const er = h * 0.12;
    return `M${x},${y + er} A${w / 2},${er} 0 0 1 ${x + w},${y + er} L${x + w},${y + h - er} A${w / 2},${er} 0 0 1 ${x},${y + h - er} Z M${x},${y + er} A${w / 2},${er} 0 0 0 ${x + w},${y + er}`;
  },
  flowChartMagneticDrum: (x, y, w, h) => {
    // Horizontal cylinder.
    const er = w * 0.12;
    return `M${x + er},${y} L${x + w - er},${y} A${er},${h / 2} 0 0 1 ${x + w - er},${y + h} L${x + er},${y + h} A${er},${h / 2} 0 0 1 ${x + er},${y} Z M${x + w - er},${y} A${er},${h / 2} 0 0 0 ${x + w - er},${y + h}`;
  },
  flowChartMagneticTape: (x, y, w, h) => {
    // Circle with a small "tail" notch at the bottom-right.
    const cx = x + w / 2;
    const cy = y + h / 2;
    const r = Math.min(w, h) / 2;
    return `M${cx + r},${cy} A${r},${r} 0 1 0 ${cx - r * 0.7},${cy + r * 0.7} L${x + w},${y + h} L${cx + r},${cy} Z`;
  },
  flowChartSummingJunction: (x, y, w, h) => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const r = Math.min(w, h) / 2;
    // Circle with an X inside (drawn as two crossing lines via subpaths).
    const off = r * Math.SQRT1_2;
    return `M${cx + r},${cy} A${r},${r} 0 1 0 ${cx - r},${cy} A${r},${r} 0 1 0 ${cx + r},${cy} Z M${cx - off},${cy - off} L${cx + off},${cy + off} M${cx - off},${cy + off} L${cx + off},${cy - off}`;
  },
  flowChartOr: (x, y, w, h) => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const r = Math.min(w, h) / 2;
    return `M${cx + r},${cy} A${r},${r} 0 1 0 ${cx - r},${cy} A${r},${r} 0 1 0 ${cx + r},${cy} Z M${cx - r},${cy} L${cx + r},${cy} M${cx},${cy - r} L${cx},${cy + r}`;
  },
  flowChartCollate: (x, y, w, h) => {
    // Hourglass.
    return `M${x},${y} L${x + w},${y} L${x},${y + h} L${x + w},${y + h} Z`;
  },
  flowChartSort: (x, y, w, h) => {
    // Diamond with a horizontal line.
    const cx = x + w / 2;
    const cy = y + h / 2;
    return `M${cx},${y} L${x + w},${cy} L${cx},${y + h} L${x},${cy} Z M${x},${cy} L${x + w},${cy}`;
  },
  flowChartExtract: (x, y, w, h) => {
    return `M${x + w / 2},${y} L${x + w},${y + h} L${x},${y + h} Z`;
  },
  flowChartMerge: (x, y, w, h) => {
    return `M${x},${y} L${x + w},${y} L${x + w / 2},${y + h} Z`;
  },
  flowChartOfflineStorage: (x, y, w, h) => {
    // Triangle with a horizontal bar at the bottom (storage symbol).
    return `M${x + w / 2},${y} L${x + w},${y + h} L${x},${y + h} Z M${x + w * 0.25},${y + h * 0.75} L${x + w * 0.75},${y + h * 0.75}`;
  },
  flowChartDelay: (x, y, w, h) => {
    // Rectangle with the right side capped by a semicircle.
    const r = h / 2;
    return `M${x},${y} L${x + w - r},${y} A${r},${r} 0 0 1 ${x + w - r},${y + h} L${x},${y + h} Z`;
  },
  flowChartDisplay: (x, y, w, h) => {
    // Trapezoid-ish display: pointed-left, rounded-right.
    const r = h / 2;
    return `M${x + r * 0.5},${y} L${x + w - r},${y} A${r},${r} 0 0 1 ${x + w - r},${y + h} L${x + r * 0.5},${y + h} L${x},${y + h / 2} Z`;
  },
  flowChartPreparation: (x, y, w, h) => {
    // Elongated hexagon.
    const cut = w * 0.18;
    return `M${x + cut},${y} L${x + w - cut},${y} L${x + w},${y + h / 2} L${x + w - cut},${y + h} L${x + cut},${y + h} L${x},${y + h / 2} Z`;
  },

  // -- Block arrows -------------------------------------------------------
  notchedRightArrow: (x, y, w, h) => {
    return `M${x},${y + h * 0.3} L${x + w * 0.65},${y + h * 0.3} L${x + w * 0.65},${y} L${x + w},${y + h / 2} L${x + w * 0.65},${y + h} L${x + w * 0.65},${y + h * 0.7} L${x},${y + h * 0.7} L${x + w * 0.15},${y + h / 2} Z`;
  },
  stripedRightArrow: (x, y, w, h) => {
    // Right arrow with two parallel "stripes" cut at the left (rendered as
    // separate strokes via subpaths).
    const stripe = w * 0.04;
    return `M${x},${y + h * 0.3} L${x + stripe},${y + h * 0.3} L${x + stripe},${y + h * 0.7} L${x},${y + h * 0.7} Z M${x + stripe * 2.5},${y + h * 0.3} L${x + stripe * 3.5},${y + h * 0.3} L${x + stripe * 3.5},${y + h * 0.7} L${x + stripe * 2.5},${y + h * 0.7} Z M${x + stripe * 5},${y + h * 0.3} L${x + w * 0.65},${y + h * 0.3} L${x + w * 0.65},${y} L${x + w},${y + h / 2} L${x + w * 0.65},${y + h} L${x + w * 0.65},${y + h * 0.7} L${x + stripe * 5},${y + h * 0.7} Z`;
  },
  curvedRightArrow: (x, y, w, h) => {
    // Quarter-arc with a triangular tip on the right.
    return `M${x},${y + h * 0.6} Q${x + w * 0.5},${y} ${x + w * 0.85},${y + h * 0.25} L${x + w},${y + h * 0.45} L${x + w * 0.85},${y + h * 0.55} L${x + w * 0.7},${y + h * 0.35} Q${x + w * 0.45},${y + h * 0.18} ${x + w * 0.15},${y + h * 0.75} Z`;
  },
  uturnArrow: (x, y, w, h) => {
    // Half-arc + tip pointing down on the right side.
    return `M${x},${y + h} L${x},${y + h / 2} A${w * 0.4},${h * 0.4} 0 0 1 ${x + w * 0.8},${y + h / 2} L${x + w * 0.8},${y + h * 0.25} L${x + w},${y + h * 0.45} L${x + w * 0.8},${y + h * 0.65} L${x + w * 0.8},${y + h / 2} A${w * 0.2},${h * 0.25} 0 0 0 ${x + w * 0.2},${y + h / 2} L${x + w * 0.2},${y + h} Z`;
  },

  // -- Brackets / braces --------------------------------------------------
  // Drawn as strokes (open paths); fill is none for these by convention.
  leftBracket: (x, y, w, h) => {
    return `M${x + w},${y} L${x},${y} L${x},${y + h} L${x + w},${y + h}`;
  },
  rightBracket: (x, y, w, h) => {
    return `M${x},${y} L${x + w},${y} L${x + w},${y + h} L${x},${y + h}`;
  },
  bracketPair: (x, y, w, h) => {
    return `M${x + w * 0.1},${y} L${x},${y} L${x},${y + h} L${x + w * 0.1},${y + h} M${x + w * 0.9},${y} L${x + w},${y} L${x + w},${y + h} L${x + w * 0.9},${y + h}`;
  },
  leftBrace: (x, y, w, h) => {
    const mid = y + h / 2;
    return `M${x + w},${y} Q${x},${y} ${x},${mid - 8} Q${x},${mid} ${x - 4},${mid} Q${x},${mid} ${x},${mid + 8} Q${x},${y + h} ${x + w},${y + h}`;
  },
  rightBrace: (x, y, w, h) => {
    const mid = y + h / 2;
    return `M${x},${y} Q${x + w},${y} ${x + w},${mid - 8} Q${x + w},${mid} ${x + w + 4},${mid} Q${x + w},${mid} ${x + w},${mid + 8} Q${x + w},${y + h} ${x},${y + h}`;
  },
  bracePair: (x, y, w, h) => {
    const mid = y + h / 2;
    return `M${x + w * 0.12},${y} Q${x},${y} ${x},${mid - 8} Q${x},${mid} ${x - 4},${mid} Q${x},${mid} ${x},${mid + 8} Q${x},${y + h} ${x + w * 0.12},${y + h} M${x + w * 0.88},${y} Q${x + w},${y} ${x + w},${mid - 8} Q${x + w},${mid} ${x + w + 4},${mid} Q${x + w},${mid} ${x + w},${mid + 8} Q${x + w},${y + h} ${x + w * 0.88},${y + h}`;
  },

  // -- Snip / round corner rects -----------------------------------------
  snip1Rect: (x, y, w, h) => {
    const c = Math.min(w, h) * 0.18;
    return `M${x},${y} L${x + w - c},${y} L${x + w},${y + c} L${x + w},${y + h} L${x},${y + h} Z`;
  },
  snip2SameRect: (x, y, w, h) => {
    const c = Math.min(w, h) * 0.18;
    return `M${x + c},${y} L${x + w - c},${y} L${x + w},${y + c} L${x + w},${y + h} L${x},${y + h} L${x},${y + c} Z`;
  },
  snip2DiagRect: (x, y, w, h) => {
    const c = Math.min(w, h) * 0.18;
    return `M${x},${y} L${x + w - c},${y} L${x + w},${y + c} L${x + w},${y + h} L${x + c},${y + h} L${x},${y + h - c} Z`;
  },
  snipRoundRect: (x, y, w, h) => {
    const c = Math.min(w, h) * 0.18;
    return `M${x + c},${y} A${c},${c} 0 0 0 ${x},${y + c} L${x},${y + h} L${x + w},${y + h} L${x + w},${y + c} L${x + w - c},${y} Z`;
  },
  round1Rect: (x, y, w, h) => {
    const r = Math.min(w, h) * 0.18;
    return `M${x},${y} L${x + w - r},${y} A${r},${r} 0 0 1 ${x + w},${y + r} L${x + w},${y + h} L${x},${y + h} Z`;
  },
  round2SameRect: (x, y, w, h) => {
    const r = Math.min(w, h) * 0.18;
    return `M${x + r},${y} L${x + w - r},${y} A${r},${r} 0 0 1 ${x + w},${y + r} L${x + w},${y + h} L${x},${y + h} L${x},${y + r} A${r},${r} 0 0 1 ${x + r},${y} Z`;
  },
  round2DiagRect: (x, y, w, h) => {
    const r = Math.min(w, h) * 0.18;
    return `M${x},${y} L${x + w - r},${y} A${r},${r} 0 0 1 ${x + w},${y + r} L${x + w},${y + h} L${x + r},${y + h} A${r},${r} 0 0 1 ${x},${y + h - r} Z`;
  },

  // -- Banners & ribbons --------------------------------------------------
  ribbon: (x, y, w, h) => {
    // Centre ribbon with two notched tails at the bottom corners.
    const notch = w * 0.06;
    const bodyTop = y + h * 0.2;
    const bodyBot = y + h * 0.8;
    return `M${x},${bodyTop} L${x + notch * 2},${y} L${x + w - notch * 2},${y} L${x + w},${bodyTop} L${x + w * 0.85},${bodyTop + (bodyBot - bodyTop) / 2} L${x + w},${bodyBot} L${x + w - notch * 2},${y + h} L${x + w - notch * 4},${bodyBot} L${x + notch * 4},${bodyBot} L${x + notch * 2},${y + h} L${x},${bodyBot} L${x + w * 0.15},${bodyTop + (bodyBot - bodyTop) / 2} Z`;
  },
  ribbon2: (x, y, w, h) => {
    // Like ribbon but the band is at the bottom.
    const notch = w * 0.06;
    const bodyTop = y + h * 0.2;
    const bodyBot = y + h * 0.8;
    return `M${x},${bodyBot} L${x + notch * 2},${y + h} L${x + w - notch * 2},${y + h} L${x + w},${bodyBot} L${x + w * 0.85},${bodyTop + (bodyBot - bodyTop) / 2} L${x + w},${bodyTop} L${x + w - notch * 2},${y} L${x + w - notch * 4},${bodyTop} L${x + notch * 4},${bodyTop} L${x + notch * 2},${y} L${x},${bodyTop} L${x + w * 0.15},${bodyTop + (bodyBot - bodyTop) / 2} Z`;
  },
  verticalScroll: (x, y, w, h) => {
    const r = w * 0.08;
    return `M${x + r},${y + r} A${r},${r} 0 0 1 ${x + r * 2},${y} L${x + w},${y} L${x + w},${y + h - r} A${r},${r} 0 0 1 ${x + w - r * 2},${y + h} L${x},${y + h} L${x},${y + r} A${r},${r} 0 0 1 ${x + r},${y + r} Z`;
  },
  horizontalScroll: (x, y, w, h) => {
    const r = h * 0.08;
    return `M${x + r},${y + r} A${r},${r} 0 0 1 ${x},${y + r * 2} L${x},${y + h} L${x + w - r},${y + h} A${r},${r} 0 0 1 ${x + w},${y + h - r * 2} L${x + w},${y} L${x + r},${y} A${r},${r} 0 0 1 ${x + r},${y + r} Z`;
  },
  wave: (x, y, w, h) => {
    return `M${x},${y + h * 0.5} C${x + w * 0.25},${y - h * 0.1} ${x + w * 0.5},${y + h * 0.85} ${x + w * 0.75},${y + h * 0.3} C${x + w * 0.85},${y + h * 0.05} ${x + w * 0.95},${y + h * 0.4} ${x + w},${y + h * 0.5} L${x + w},${y + h} C${x + w * 0.75},${y + h * 0.4} ${x + w * 0.5},${y + h * 1.1} ${x + w * 0.25},${y + h * 0.55} C${x + w * 0.15},${y + h * 0.3} ${x + w * 0.05},${y + h * 0.6} ${x},${y + h * 0.55} Z`;
  },
  doubleWave: (x, y, w, h) => {
    return `M${x},${y + h * 0.4} C${x + w * 0.15},${y - h * 0.05} ${x + w * 0.35},${y + h * 0.65} ${x + w * 0.5},${y + h * 0.3} C${x + w * 0.65},${y - h * 0.05} ${x + w * 0.85},${y + h * 0.65} ${x + w},${y + h * 0.4} L${x + w},${y + h} C${x + w * 0.85},${y + h * 0.4} ${x + w * 0.65},${y + h * 1.05} ${x + w * 0.5},${y + h * 0.7} C${x + w * 0.35},${y + h * 1.05} ${x + w * 0.15},${y + h * 0.4} ${x},${y + h} Z`;
  },

  // -- Math operators -----------------------------------------------------
  mathPlus: (x, y, w, h) => {
    const t = Math.min(w, h) * 0.2;
    const cx = x + w / 2;
    const cy = y + h / 2;
    return `M${cx - t / 2},${y} L${cx + t / 2},${y} L${cx + t / 2},${cy - t / 2} L${x + w},${cy - t / 2} L${x + w},${cy + t / 2} L${cx + t / 2},${cy + t / 2} L${cx + t / 2},${y + h} L${cx - t / 2},${y + h} L${cx - t / 2},${cy + t / 2} L${x},${cy + t / 2} L${x},${cy - t / 2} L${cx - t / 2},${cy - t / 2} Z`;
  },
  mathMinus: (x, y, w, h) => {
    const t = h * 0.3;
    const cy = y + h / 2;
    return `M${x},${cy - t / 2} L${x + w},${cy - t / 2} L${x + w},${cy + t / 2} L${x},${cy + t / 2} Z`;
  },
  mathMultiply: (x, y, w, h) => {
    const t = Math.min(w, h) * 0.16;
    const cx = x + w / 2;
    const cy = y + h / 2;
    // Two crossed rects (diagonals).
    return `M${x},${y + t} L${x + t},${y} L${cx},${cy - t} L${x + w - t},${y} L${x + w},${y + t} L${cx + t},${cy} L${x + w},${y + h - t} L${x + w - t},${y + h} L${cx},${cy + t} L${x + t},${y + h} L${x},${y + h - t} L${cx - t},${cy} Z`;
  },
  mathDivide: (x, y, w, h) => {
    const dot = Math.min(w, h) * 0.1;
    const cx = x + w / 2;
    const cy = y + h / 2;
    return `M${x},${cy - dot * 0.4} L${x + w},${cy - dot * 0.4} L${x + w},${cy + dot * 0.4} L${x},${cy + dot * 0.4} Z M${cx - dot},${y + h * 0.18} A${dot},${dot} 0 1 0 ${cx + dot},${y + h * 0.18} A${dot},${dot} 0 1 0 ${cx - dot},${y + h * 0.18} Z M${cx - dot},${y + h * 0.82} A${dot},${dot} 0 1 0 ${cx + dot},${y + h * 0.82} A${dot},${dot} 0 1 0 ${cx - dot},${y + h * 0.82} Z`;
  },
  mathEqual: (x, y, w, h) => {
    const t = h * 0.2;
    return `M${x},${y + h * 0.3 - t / 2} L${x + w},${y + h * 0.3 - t / 2} L${x + w},${y + h * 0.3 + t / 2} L${x},${y + h * 0.3 + t / 2} Z M${x},${y + h * 0.7 - t / 2} L${x + w},${y + h * 0.7 - t / 2} L${x + w},${y + h * 0.7 + t / 2} L${x},${y + h * 0.7 + t / 2} Z`;
  },
  mathNotEqual: (x, y, w, h) => {
    const t = h * 0.15;
    const cy = y + h / 2;
    // Equal sign + diagonal bar (use path subpaths).
    return `M${x},${cy - h * 0.18 - t / 2} L${x + w},${cy - h * 0.18 - t / 2} L${x + w},${cy - h * 0.18 + t / 2} L${x},${cy - h * 0.18 + t / 2} Z M${x},${cy + h * 0.18 - t / 2} L${x + w},${cy + h * 0.18 - t / 2} L${x + w},${cy + h * 0.18 + t / 2} L${x},${cy + h * 0.18 + t / 2} Z M${x + w * 0.7},${y} L${x + w * 0.85},${y} L${x + w * 0.3},${y + h} L${x + w * 0.15},${y + h} Z`;
  },

  // -- Action button glyphs (the chrome is a roundRect; we just add the
  // glyph). Real action buttons are nested shapes; we approximate.
  actionButtonBlank: (x, y, w, h) => {
    const r = Math.min(w, h) * 0.06;
    return `M${x + r},${y} L${x + w - r},${y} A${r},${r} 0 0 1 ${x + w},${y + r} L${x + w},${y + h - r} A${r},${r} 0 0 1 ${x + w - r},${y + h} L${x + r},${y + h} A${r},${r} 0 0 1 ${x},${y + h - r} L${x},${y + r} A${r},${r} 0 0 1 ${x + r},${y} Z`;
  },

  // -- Explosion / starburst callouts (the "jagged" speech marks). ------
  // `irregularSeal1` and `irregularSeal2` are PowerPoint's two
  // explosion-style callouts. The geometry is a deterministic pseudo-
  // random pattern of long + short rays; we mimic that without trying
  // to match the spec point-for-point.
  irregularSeal1: (x, y, w, h) => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const rx = w / 2;
    const ry = h / 2;
    // 16 rays — outer radii drawn from a fixed offset table so the
    // shape comes out spiky-but-balanced like PowerPoint's.
    const offsets = [
      1.0, 0.45, 0.95, 0.5, 1.0, 0.4, 0.9, 0.55, 1.0, 0.45, 0.95, 0.5, 1.0, 0.4, 0.9, 0.55,
    ];
    const points: string[] = [];
    for (let i = 0; i < offsets.length; i++) {
      const a = (i / offsets.length) * 2 * Math.PI - Math.PI / 2;
      const offset = offsets[i] ?? 1;
      const r = offset;
      const px0 = cx + rx * r * Math.cos(a);
      const py0 = cy + ry * r * Math.sin(a);
      points.push(`${i === 0 ? 'M' : 'L'}${px0},${py0}`);
    }
    points.push('Z');
    return points.join(' ');
  },
  irregularSeal2: (x, y, w, h) => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const rx = w / 2;
    const ry = h / 2;
    // 24 alternating rays for a denser, more "scribble"-like burst.
    const offsets = [
      1.0, 0.4, 0.95, 0.5, 0.9, 0.35, 0.85, 0.45, 1.0, 0.4, 0.95, 0.5, 0.9, 0.35, 0.85, 0.45, 1.0,
      0.4, 0.95, 0.5, 0.9, 0.35, 0.85, 0.45,
    ];
    const points: string[] = [];
    for (let i = 0; i < offsets.length; i++) {
      const a = (i / offsets.length) * 2 * Math.PI - Math.PI / 2;
      const offset = offsets[i] ?? 1;
      const r = offset;
      const px0 = cx + rx * r * Math.cos(a);
      const py0 = cy + ry * r * Math.sin(a);
      points.push(`${i === 0 ? 'M' : 'L'}${px0},${py0}`);
    }
    points.push('Z');
    return points.join(' ');
  },
  // `cloudCallout` lives above; "cloud" without callout is just the
  // body without the tail dots.
  cloud: (x, y, w, h) => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const rx = (w / 2) * 0.92;
    const ry = (h / 2) * 0.85;
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
    return path.join(' ');
  },

  // -- Pies / chord / teardrop / arc / blockArc / moon ------------------
  pie: (x, y, w, h) => {
    // 270° pie (default), missing the top-right quadrant.
    const cx = x + w / 2;
    const cy = y + h / 2;
    const r = Math.min(w, h) / 2;
    return `M${cx},${cy} L${cx + r},${cy} A${r},${r} 0 1 1 ${cx},${cy - r} Z`;
  },
  chord: (x, y, w, h) => {
    // Ellipse minus a triangular chord (default cut along the top).
    const cx = x + w / 2;
    const cy = y + h / 2;
    const rx = w / 2;
    const ry = h / 2;
    return `M${cx + rx},${cy} A${rx},${ry} 0 1 1 ${cx - rx},${cy} L${cx + rx},${cy} Z`;
  },
  teardrop: (x, y, w, h) => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const rx = w / 2;
    const ry = h / 2;
    // Three-quarter circle + a pointed tip at the top-right corner.
    return `M${cx},${y} L${x + w},${y} L${x + w},${cy} A${rx},${ry} 0 1 1 ${cx - rx},${cy} A${rx},${ry} 0 0 1 ${cx},${y} Z`;
  },
  arc: (x, y, w, h) => {
    // Open arc (stroke only). Drawn as a half-pie outline.
    const cx = x + w / 2;
    const cy = y + h / 2;
    const rx = w / 2;
    const ry = h / 2;
    return `M${cx + rx},${cy} A${rx},${ry} 0 1 1 ${cx},${cy + ry}`;
  },
  blockArc: (x, y, w, h) => {
    // 270° annulus.
    const cx = x + w / 2;
    const cy = y + h / 2;
    const outerR = Math.min(w, h) / 2;
    const innerR = outerR * 0.6;
    return `M${cx + outerR},${cy} A${outerR},${outerR} 0 1 1 ${cx},${cy - outerR} L${cx},${cy - innerR} A${innerR},${innerR} 0 1 0 ${cx + innerR},${cy} Z`;
  },
  moon: (x, y, w, h) => {
    // Crescent: big ellipse minus a smaller offset ellipse.
    const cx = x + w / 2;
    const cy = y + h / 2;
    const rx = w / 2;
    const ry = h / 2;
    const innerRx = rx * 0.78;
    const offsetX = rx * 0.32;
    return `M${cx + rx},${cy} A${rx},${ry} 0 1 1 ${cx - rx},${cy} A${rx},${ry} 0 1 1 ${cx + rx},${cy} M${cx - rx + offsetX + innerRx},${cy} A${innerRx},${ry * 0.95} 0 1 1 ${cx - rx + offsetX - innerRx},${cy} A${innerRx},${ry * 0.95} 0 1 1 ${cx - rx + offsetX + innerRx},${cy}`;
  },

  // -- Plates / plaques / frames / corners ------------------------------
  plus: (x, y, w, h) => {
    const t = Math.min(w, h) * 0.3;
    const cx = x + w / 2;
    const cy = y + h / 2;
    return `M${cx - t / 2},${y} L${cx + t / 2},${y} L${cx + t / 2},${cy - t / 2} L${x + w},${cy - t / 2} L${x + w},${cy + t / 2} L${cx + t / 2},${cy + t / 2} L${cx + t / 2},${y + h} L${cx - t / 2},${y + h} L${cx - t / 2},${cy + t / 2} L${x},${cy + t / 2} L${x},${cy - t / 2} L${cx - t / 2},${cy - t / 2} Z`;
  },
  plaque: (x, y, w, h) => {
    const r = Math.min(w, h) * 0.18;
    // Rounded corners that arc INWARD.
    return `M${x + r},${y} L${x + w - r},${y} A${r},${r} 0 0 0 ${x + w},${y + r} L${x + w},${y + h - r} A${r},${r} 0 0 0 ${x + w - r},${y + h} L${x + r},${y + h} A${r},${r} 0 0 0 ${x},${y + h - r} L${x},${y + r} A${r},${r} 0 0 0 ${x + r},${y} Z`;
  },
  can: (x, y, w, h) => {
    // Cylinder rendered upright; same path as flowChartMagneticDisk.
    const er = h * 0.12;
    return `M${x},${y + er} A${w / 2},${er} 0 0 1 ${x + w},${y + er} L${x + w},${y + h - er} A${w / 2},${er} 0 0 1 ${x},${y + h - er} Z M${x},${y + er} A${w / 2},${er} 0 0 0 ${x + w},${y + er}`;
  },
  cube: (x, y, w, h) => {
    const d = Math.min(w, h) * 0.2;
    return `M${x},${y + d} L${x + d},${y} L${x + w},${y} L${x + w},${y + h - d} L${x + w - d},${y + h} L${x},${y + h} Z M${x},${y + d} L${x + w - d},${y + d} L${x + w},${y} M${x + w - d},${y + d} L${x + w - d},${y + h}`;
  },
  bevel: (x, y, w, h) => {
    const d = Math.min(w, h) * 0.12;
    return `M${x},${y} L${x + w},${y} L${x + w},${y + h} L${x},${y + h} Z M${x + d},${y + d} L${x + w - d},${y + d} L${x + w - d},${y + h - d} L${x + d},${y + h - d} Z M${x},${y} L${x + d},${y + d} M${x + w},${y} L${x + w - d},${y + d} M${x},${y + h} L${x + d},${y + h - d} M${x + w},${y + h} L${x + w - d},${y + h - d}`;
  },
  donut: (x, y, w, h) => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const r = Math.min(w, h) / 2;
    const innerR = r * 0.65;
    return `M${cx + r},${cy} A${r},${r} 0 1 0 ${cx - r},${cy} A${r},${r} 0 1 0 ${cx + r},${cy} Z M${cx + innerR},${cy} A${innerR},${innerR} 0 1 1 ${cx - innerR},${cy} A${innerR},${innerR} 0 1 1 ${cx + innerR},${cy} Z`;
  },
  noSmoking: (x, y, w, h) => {
    // Donut + diagonal bar.
    const cx = x + w / 2;
    const cy = y + h / 2;
    const r = Math.min(w, h) / 2;
    const innerR = r * 0.78;
    const t = r * 0.12;
    return `M${cx + r},${cy} A${r},${r} 0 1 0 ${cx - r},${cy} A${r},${r} 0 1 0 ${cx + r},${cy} Z M${cx + innerR},${cy} A${innerR},${innerR} 0 1 1 ${cx - innerR},${cy} A${innerR},${innerR} 0 1 1 ${cx + innerR},${cy} Z M${cx - r * 0.71 - t},${cy - r * 0.71 + t} L${cx - r * 0.71 + t},${cy - r * 0.71 - t} L${cx + r * 0.71 + t},${cy + r * 0.71 - t} L${cx + r * 0.71 - t},${cy + r * 0.71 + t} Z`;
  },
  frame: (x, y, w, h) => {
    // Picture-frame: outer rect minus inner rect.
    const f = Math.min(w, h) * 0.1;
    return `M${x},${y} L${x + w},${y} L${x + w},${y + h} L${x},${y + h} Z M${x + f},${y + f} L${x + f},${y + h - f} L${x + w - f},${y + h - f} L${x + w - f},${y + f} Z`;
  },
  halfFrame: (x, y, w, h) => {
    const f = Math.min(w, h) * 0.15;
    return `M${x},${y} L${x + w},${y} L${x + w - f},${y + f} L${x + f},${y + f} L${x + f},${y + h - f} L${x},${y + h} Z`;
  },
  corner: (x, y, w, h) => {
    // L-shaped corner piece.
    const tx = w * 0.4;
    const ty = h * 0.4;
    return `M${x},${y} L${x + tx},${y} L${x + tx},${y + h - ty} L${x + w},${y + h - ty} L${x + w},${y + h} L${x},${y + h} Z`;
  },
  diagStripe: (x, y, w, h) => {
    // Diagonal stripe filling the upper-left half of the bounding box.
    return `M${x},${y} L${x + w * 0.6},${y} L${x},${y + h * 0.6} Z`;
  },

  // -- Ellipse ribbons ---------------------------------------------------
  ellipseRibbon: (x, y, w, h) => {
    // Curved banner — top arc with two notched tails like `ribbon`.
    const notch = w * 0.08;
    const bodyTop = y + h * 0.2;
    const bodyBot = y + h * 0.85;
    const arcDip = h * 0.15;
    return `M${x},${bodyTop} C${x + w * 0.3},${bodyTop - arcDip} ${x + w * 0.7},${bodyTop - arcDip} ${x + w},${bodyTop} L${x + w * 0.85},${bodyTop + (bodyBot - bodyTop) / 2} L${x + w},${bodyBot} L${x + w - notch * 2},${y + h} L${x + w - notch * 4},${bodyBot} C${x + w * 0.7},${bodyBot + arcDip * 0.4} ${x + w * 0.3},${bodyBot + arcDip * 0.4} L${x + notch * 4},${bodyBot} L${x + notch * 2},${y + h} L${x},${bodyBot} L${x + w * 0.15},${bodyTop + (bodyBot - bodyTop) / 2} Z`;
  },
  ellipseRibbon2: (x, y, w, h) => {
    // Inverted version of ellipseRibbon (band at top, arc at bottom).
    const notch = w * 0.08;
    const bodyTop = y + h * 0.15;
    const bodyBot = y + h * 0.8;
    const arcRise = h * 0.15;
    return `M${x},${bodyBot} C${x + w * 0.3},${bodyBot + arcRise} ${x + w * 0.7},${bodyBot + arcRise} ${x + w},${bodyBot} L${x + w * 0.85},${bodyTop + (bodyBot - bodyTop) / 2} L${x + w},${bodyTop} L${x + w - notch * 2},${y} L${x + w - notch * 4},${bodyTop} C${x + w * 0.7},${bodyTop - arcRise * 0.4} ${x + w * 0.3},${bodyTop - arcRise * 0.4} L${x + notch * 4},${bodyTop} L${x + notch * 2},${y} L${x},${bodyTop} L${x + w * 0.15},${bodyTop + (bodyBot - bodyTop) / 2} Z`;
  },

  // -- Block arrows: the rest of the cardinal & curved family -----------
  quadArrow: (x, y, w, h) => {
    // Plus-sign of four arrows pointing outward.
    const cx = x + w / 2;
    const cy = y + h / 2;
    const tip = 0.2; // tip depth fraction
    const stem = 0.15; // stem width fraction
    const headW = 0.35; // head half-width fraction
    return `M${cx},${y} L${cx + headW * w},${y + tip * h} L${cx + stem * w},${y + tip * h} L${cx + stem * w},${cy - stem * h} L${x + w - tip * w},${cy - stem * h} L${x + w - tip * w},${cy - headW * h} L${x + w},${cy} L${x + w - tip * w},${cy + headW * h} L${x + w - tip * w},${cy + stem * h} L${cx + stem * w},${cy + stem * h} L${cx + stem * w},${y + h - tip * h} L${cx + headW * w},${y + h - tip * h} L${cx},${y + h} L${cx - headW * w},${y + h - tip * h} L${cx - stem * w},${y + h - tip * h} L${cx - stem * w},${cy + stem * h} L${x + tip * w},${cy + stem * h} L${x + tip * w},${cy + headW * h} L${x},${cy} L${x + tip * w},${cy - headW * h} L${x + tip * w},${cy - stem * h} L${cx - stem * w},${cy - stem * h} L${cx - stem * w},${y + tip * h} L${cx - headW * w},${y + tip * h} Z`;
  },
  leftRightUpArrow: (x, y, w, h) => {
    const cx = x + w / 2;
    const tip = 0.2;
    const stem = 0.15;
    const headW = 0.35;
    return `M${cx},${y} L${cx + headW * w},${y + tip * h} L${cx + stem * w},${y + tip * h} L${cx + stem * w},${y + h - tip * h} L${x + w - tip * w},${y + h - tip * h} L${x + w - tip * w},${y + h - headW * h} L${x + w},${y + h} L${x + w - tip * w},${y + h + headW * h} L${x + w - tip * w},${y + h - tip * h} L${cx - stem * w},${y + h - tip * h} L${cx - stem * w},${y + tip * h} L${cx - headW * w},${y + tip * h} Z`;
  },
  bentUpArrow: (x, y, w, h) => {
    const stem = 0.3;
    const tip = 0.25;
    return `M${x},${y + h * 0.55} L${x + w * 0.5},${y + h * 0.55} L${x + w * 0.5},${y + tip * h} L${x + w * (0.5 - stem * 0.5)},${y + tip * h} L${x + w * 0.75},${y} L${x + w},${y + tip * h} L${x + w * (0.5 + stem * 0.5)},${y + tip * h} L${x + w * (0.5 + stem * 0.5)},${y + h * 0.55 + h * 0.4} L${x},${y + h * 0.55 + h * 0.4} Z`;
  },
  curvedLeftArrow: (x, y, w, h) => {
    return `M${x + w},${y + h * 0.6} Q${x + w * 0.5},${y} ${x + w * 0.15},${y + h * 0.25} L${x},${y + h * 0.45} L${x + w * 0.15},${y + h * 0.55} L${x + w * 0.3},${y + h * 0.35} Q${x + w * 0.55},${y + h * 0.18} ${x + w * 0.85},${y + h * 0.75} Z`;
  },
  curvedUpArrow: (x, y, w, h) => {
    return `M${x + w * 0.4},${y + h} Q${x},${y + h * 0.5} ${x + w * 0.25},${y + h * 0.15} L${x + w * 0.45},${y} L${x + w * 0.55},${y + h * 0.15} L${x + w * 0.35},${y + h * 0.3} Q${x + w * 0.18},${y + h * 0.55} ${x + w * 0.75},${y + h * 0.85} Z`;
  },
  curvedDownArrow: (x, y, w, h) => {
    return `M${x + w * 0.4},${y} Q${x},${y + h * 0.5} ${x + w * 0.25},${y + h * 0.85} L${x + w * 0.45},${y + h} L${x + w * 0.55},${y + h * 0.85} L${x + w * 0.35},${y + h * 0.7} Q${x + w * 0.18},${y + h * 0.45} ${x + w * 0.75},${y + h * 0.15} Z`;
  },
  swooshArrow: (x, y, w, h) => {
    // Stylised curved arrow with a small tail to the upper-left.
    return `M${x},${y + h * 0.75} C${x + w * 0.35},${y + h} ${x + w * 0.65},${y + h * 0.3} ${x + w * 0.75},${y + h * 0.2} L${x + w * 0.65},${y + h * 0.05} L${x + w},${y + h * 0.18} L${x + w * 0.78},${y + h * 0.45} L${x + w * 0.7},${y + h * 0.3} C${x + w * 0.55},${y + h * 0.6} ${x + w * 0.35},${y + h * 0.9} ${x},${y + h * 0.85} Z`;
  },
  circularArrow: (x, y, w, h) => {
    // 270° annulus + small arrow tip at the open end.
    const cx = x + w / 2;
    const cy = y + h / 2;
    const outerR = Math.min(w, h) / 2;
    const innerR = outerR * 0.62;
    const midR = (outerR + innerR) / 2;
    // Arc goes from 12 o'clock (top) clockwise to 9 o'clock (left), with
    // a triangular tip pointing further clockwise from there.
    return `M${cx},${cy - outerR} A${outerR},${outerR} 0 1 1 ${cx - outerR},${cy} L${cx - midR - midR * 0.25},${cy + outerR * 0.15} L${cx - midR + midR * 0.25},${cy + outerR * 0.3} L${cx - innerR},${cy} A${innerR},${innerR} 0 1 0 ${cx},${cy - innerR} Z`;
  },
  leftCircularArrow: (x, y, w, h) => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const outerR = Math.min(w, h) / 2;
    const innerR = outerR * 0.62;
    return `M${cx},${cy - outerR} A${outerR},${outerR} 0 1 0 ${cx + outerR},${cy} L${cx + (outerR + innerR) / 2 + ((outerR + innerR) / 2) * 0.25},${cy + outerR * 0.15} L${cx + (outerR + innerR) / 2 - ((outerR + innerR) / 2) * 0.25},${cy + outerR * 0.3} L${cx + innerR},${cy} A${innerR},${innerR} 0 1 1 ${cx},${cy - innerR} Z`;
  },
  leftRightCircularArrow: (x, y, w, h) => {
    // Two-headed circular arrow (left + right tips on a 270° annulus).
    const cx = x + w / 2;
    const cy = y + h / 2;
    const outerR = Math.min(w, h) / 2;
    const innerR = outerR * 0.62;
    return `M${cx - outerR * 0.15},${cy - outerR * 1.05} L${cx + outerR * 0.15},${cy - outerR * 1.05} L${cx + outerR * 0.1},${cy - outerR * 0.85} A${outerR},${outerR} 0 1 1 ${cx - outerR * 0.1},${cy - outerR * 0.85} Z M${cx},${cy - innerR} A${innerR},${innerR} 0 1 0 ${cx},${cy + innerR}`;
  },

  // -- Arrow callouts (block arrow body + flat rect for text) -----------
  rightArrowCallout: (x, y, w, h) => {
    const headW = w * 0.25;
    return `M${x},${y + h * 0.3} L${x + w - headW},${y + h * 0.3} L${x + w - headW},${y} L${x + w},${y + h / 2} L${x + w - headW},${y + h} L${x + w - headW},${y + h * 0.7} L${x},${y + h * 0.7} Z`;
  },
  leftArrowCallout: (x, y, w, h) => {
    const headW = w * 0.25;
    return `M${x + headW},${y + h * 0.3} L${x + w},${y + h * 0.3} L${x + w},${y + h * 0.7} L${x + headW},${y + h * 0.7} L${x + headW},${y + h} L${x},${y + h / 2} L${x + headW},${y} Z`;
  },
  upArrowCallout: (x, y, w, h) => {
    const headH = h * 0.25;
    return `M${x + w * 0.3},${y + headH} L${x + w * 0.3},${y + h} L${x + w * 0.7},${y + h} L${x + w * 0.7},${y + headH} L${x + w},${y + headH} L${x + w / 2},${y} L${x},${y + headH} Z`;
  },
  downArrowCallout: (x, y, w, h) => {
    const headH = h * 0.25;
    return `M${x + w * 0.3},${y} L${x + w * 0.7},${y} L${x + w * 0.7},${y + h - headH} L${x + w},${y + h - headH} L${x + w / 2},${y + h} L${x},${y + h - headH} L${x + w * 0.3},${y + h - headH} Z`;
  },
  leftRightArrowCallout: (x, y, w, h) => {
    const headW = w * 0.2;
    return `M${x},${y + h / 2} L${x + headW},${y} L${x + headW},${y + h * 0.3} L${x + w - headW},${y + h * 0.3} L${x + w - headW},${y} L${x + w},${y + h / 2} L${x + w - headW},${y + h} L${x + w - headW},${y + h * 0.7} L${x + headW},${y + h * 0.7} L${x + headW},${y + h} Z`;
  },
  upDownArrowCallout: (x, y, w, h) => {
    const headH = h * 0.2;
    return `M${x + w / 2},${y} L${x},${y + headH} L${x + w * 0.3},${y + headH} L${x + w * 0.3},${y + h - headH} L${x},${y + h - headH} L${x + w / 2},${y + h} L${x + w},${y + h - headH} L${x + w * 0.7},${y + h - headH} L${x + w * 0.7},${y + headH} L${x + w},${y + headH} Z`;
  },
  quadArrowCallout: (x, y, w, h) => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const tip = 0.16;
    const stem = 0.18;
    const headW = 0.32;
    const txt = 0.28; // text-area half-width fraction
    return `M${cx - txt * w},${cy - txt * h} L${cx - txt * w},${cy - stem * h} L${cx - stem * w},${cy - stem * h} L${cx - stem * w},${y + tip * h} L${cx - headW * w},${y + tip * h} L${cx},${y} L${cx + headW * w},${y + tip * h} L${cx + stem * w},${y + tip * h} L${cx + stem * w},${cy - stem * h} L${cx + txt * w},${cy - stem * h} L${cx + txt * w},${cy - txt * h} L${x + w - tip * w},${cy - txt * h} L${x + w - tip * w},${cy - headW * h} L${x + w},${cy} L${x + w - tip * w},${cy + headW * h} L${x + w - tip * w},${cy + txt * h} L${cx + txt * w},${cy + txt * h} L${cx + txt * w},${cy + stem * h} L${cx + stem * w},${cy + stem * h} L${cx + stem * w},${y + h - tip * h} L${cx + headW * w},${y + h - tip * h} L${cx},${y + h} L${cx - headW * w},${y + h - tip * h} L${cx - stem * w},${y + h - tip * h} L${cx - stem * w},${cy + stem * h} L${cx - txt * w},${cy + stem * h} L${cx - txt * w},${cy + txt * h} L${x + tip * w},${cy + txt * h} L${x + tip * w},${cy + headW * h} L${x},${cy} L${x + tip * w},${cy - headW * h} L${x + tip * w},${cy - txt * h} Z`;
  },

  // -- Action button chrome (rounded rect) + glyph silhouettes ----------
  // Each button is rounded-rect chrome + a glyph drawn as a subpath
  // using the same path's `evenodd` fill rule so the glyph "punches"
  // out of the chrome.
  actionButtonHome: (x, y, w, h) => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const s = Math.min(w, h) * 0.3;
    return `${PRESET_PATHS.actionButtonBlank?.(x, y, w, h) ?? ''} M${cx - s},${cy + s * 0.6} L${cx - s},${cy - s * 0.1} L${cx},${cy - s * 0.7} L${cx + s},${cy - s * 0.1} L${cx + s},${cy + s * 0.6} L${cx + s * 0.3},${cy + s * 0.6} L${cx + s * 0.3},${cy + s * 0.1} L${cx - s * 0.3},${cy + s * 0.1} L${cx - s * 0.3},${cy + s * 0.6} Z`;
  },
  actionButtonForwardNext: (x, y, w, h) => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const s = Math.min(w, h) * 0.3;
    return `${PRESET_PATHS.actionButtonBlank?.(x, y, w, h) ?? ''} M${cx - s * 0.7},${cy - s} L${cx + s * 0.7},${cy} L${cx - s * 0.7},${cy + s} Z`;
  },
  actionButtonBackPrevious: (x, y, w, h) => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const s = Math.min(w, h) * 0.3;
    return `${PRESET_PATHS.actionButtonBlank?.(x, y, w, h) ?? ''} M${cx + s * 0.7},${cy - s} L${cx - s * 0.7},${cy} L${cx + s * 0.7},${cy + s} Z`;
  },
  actionButtonEnd: (x, y, w, h) => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const s = Math.min(w, h) * 0.3;
    return `${PRESET_PATHS.actionButtonBlank?.(x, y, w, h) ?? ''} M${cx - s},${cy - s} L${cx + s * 0.4},${cy} L${cx - s},${cy + s} Z M${cx + s * 0.5},${cy - s} L${cx + s},${cy - s} L${cx + s},${cy + s} L${cx + s * 0.5},${cy + s} Z`;
  },
  actionButtonBeginning: (x, y, w, h) => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const s = Math.min(w, h) * 0.3;
    return `${PRESET_PATHS.actionButtonBlank?.(x, y, w, h) ?? ''} M${cx + s},${cy - s} L${cx - s * 0.4},${cy} L${cx + s},${cy + s} Z M${cx - s * 0.5},${cy - s} L${cx - s},${cy - s} L${cx - s},${cy + s} L${cx - s * 0.5},${cy + s} Z`;
  },
  actionButtonReturn: (x, y, w, h) => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const s = Math.min(w, h) * 0.28;
    return `${PRESET_PATHS.actionButtonBlank?.(x, y, w, h) ?? ''} M${cx + s},${cy - s} L${cx + s * 0.4},${cy - s} L${cx + s * 0.4},${cy + s * 0.2} L${cx - s * 0.2},${cy + s * 0.2} L${cx - s * 0.2},${cy - s * 0.2} L${cx - s},${cy + s * 0.2} L${cx - s * 0.2},${cy + s} L${cx - s * 0.2},${cy + s * 0.5} L${cx + s},${cy + s * 0.5} Z`;
  },
  actionButtonHelp: (x, y, w, h) => {
    // Approximated "?" as a curved path; readable at typical button sizes.
    const cx = x + w / 2;
    const cy = y + h / 2;
    const s = Math.min(w, h) * 0.3;
    return `${PRESET_PATHS.actionButtonBlank?.(x, y, w, h) ?? ''} M${cx - s * 0.4},${cy - s * 0.4} Q${cx - s * 0.4},${cy - s} ${cx},${cy - s} Q${cx + s * 0.4},${cy - s} ${cx + s * 0.4},${cy - s * 0.4} Q${cx + s * 0.4},${cy} ${cx},${cy} L${cx},${cy + s * 0.4} M${cx - s * 0.18},${cy + s * 0.8} L${cx + s * 0.18},${cy + s * 0.8} L${cx + s * 0.18},${cy + s} L${cx - s * 0.18},${cy + s} Z`;
  },
  actionButtonInformation: (x, y, w, h) => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const s = Math.min(w, h) * 0.3;
    return `${PRESET_PATHS.actionButtonBlank?.(x, y, w, h) ?? ''} M${cx - s * 0.2},${cy - s * 0.7} L${cx + s * 0.2},${cy - s * 0.7} L${cx + s * 0.2},${cy - s * 0.35} L${cx - s * 0.2},${cy - s * 0.35} Z M${cx - s * 0.2},${cy - s * 0.1} L${cx + s * 0.2},${cy - s * 0.1} L${cx + s * 0.2},${cy + s} L${cx - s * 0.2},${cy + s} Z`;
  },
  actionButtonDocument: (x, y, w, h) => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const s = Math.min(w, h) * 0.3;
    return `${PRESET_PATHS.actionButtonBlank?.(x, y, w, h) ?? ''} M${cx - s * 0.6},${cy - s} L${cx + s * 0.3},${cy - s} L${cx + s * 0.6},${cy - s * 0.7} L${cx + s * 0.6},${cy + s} L${cx - s * 0.6},${cy + s} Z M${cx + s * 0.3},${cy - s} L${cx + s * 0.3},${cy - s * 0.7} L${cx + s * 0.6},${cy - s * 0.7}`;
  },
  actionButtonSound: (x, y, w, h) => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const s = Math.min(w, h) * 0.3;
    // Speaker silhouette: trapezoid cone + small rectangle.
    return `${PRESET_PATHS.actionButtonBlank?.(x, y, w, h) ?? ''} M${cx - s},${cy - s * 0.4} L${cx - s * 0.3},${cy - s * 0.4} L${cx + s * 0.3},${cy - s} L${cx + s * 0.3},${cy + s} L${cx - s * 0.3},${cy + s * 0.4} L${cx - s},${cy + s * 0.4} Z M${cx + s * 0.55},${cy - s * 0.4} Q${cx + s * 0.95},${cy} ${cx + s * 0.55},${cy + s * 0.4}`;
  },
  actionButtonMovie: (x, y, w, h) => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const s = Math.min(w, h) * 0.3;
    // Filmstrip — outer rect + sprocket holes.
    const out: string[] = [PRESET_PATHS.actionButtonBlank?.(x, y, w, h) ?? ''];
    out.push(
      `M${cx - s},${cy - s * 0.6} L${cx + s},${cy - s * 0.6} L${cx + s},${cy + s * 0.6} L${cx - s},${cy + s * 0.6} Z`,
    );
    for (let i = 0; i < 4; i++) {
      const px0 = cx - s + (i + 0.5) * ((s * 2) / 4);
      out.push(
        `M${px0 - s * 0.08},${cy - s * 0.45} L${px0 + s * 0.08},${cy - s * 0.45} L${px0 + s * 0.08},${cy - s * 0.3} L${px0 - s * 0.08},${cy - s * 0.3} Z`,
      );
      out.push(
        `M${px0 - s * 0.08},${cy + s * 0.3} L${px0 + s * 0.08},${cy + s * 0.3} L${px0 + s * 0.08},${cy + s * 0.45} L${px0 - s * 0.08},${cy + s * 0.45} Z`,
      );
    }
    return out.join(' ');
  },

  // -- Border/accent callouts (simplified — line callouts without
  // adjustable elbows; rendered as a rect + a single connecting line).
  borderCallout1: (x, y, w, h) => {
    return `M${x},${y} L${x + w * 0.6},${y} L${x + w * 0.6},${y + h * 0.5} L${x},${y + h * 0.5} Z M${x + w * 0.6},${y + h * 0.5} L${x + w},${y + h}`;
  },
  borderCallout2: (x, y, w, h) => {
    return `M${x},${y} L${x + w * 0.6},${y} L${x + w * 0.6},${y + h * 0.5} L${x},${y + h * 0.5} Z M${x + w * 0.6},${y + h * 0.5} L${x + w * 0.85},${y + h * 0.75} L${x + w},${y + h}`;
  },
  borderCallout3: (x, y, w, h) => {
    return `M${x},${y} L${x + w * 0.6},${y} L${x + w * 0.6},${y + h * 0.5} L${x},${y + h * 0.5} Z M${x + w * 0.6},${y + h * 0.5} L${x + w * 0.75},${y + h * 0.65} L${x + w * 0.9},${y + h * 0.65} L${x + w},${y + h}`;
  },
  accentCallout1: (x, y, w, h) => {
    // Frame + vertical accent bar on the right edge.
    return `M${x},${y} L${x + w * 0.6},${y} L${x + w * 0.6},${y + h * 0.5} L${x},${y + h * 0.5} Z M${x + w * 0.58},${y} L${x + w * 0.58},${y + h * 0.5} M${x + w * 0.6},${y + h * 0.5} L${x + w},${y + h}`;
  },
  accentBorderCallout1: (x, y, w, h) => {
    return PRESET_PATHS.accentCallout1?.(x, y, w, h) ?? '';
  },
  callout1: (x, y, w, h) => {
    // Single-segment line callout (no body box).
    return `M${x},${y + h * 0.5} L${x + w},${y + h}`;
  },
  callout2: (x, y, w, h) => {
    return `M${x},${y + h * 0.5} L${x + w * 0.6},${y + h * 0.7} L${x + w},${y + h}`;
  },
  callout3: (x, y, w, h) => {
    return `M${x},${y + h * 0.5} L${x + w * 0.4},${y + h * 0.55} L${x + w * 0.7},${y + h * 0.8} L${x + w},${y + h}`;
  },

  // -- Connectors / lines (rendered as straight diagonals; PowerPoint
  // routes these dynamically but the static preview just shows where
  // the endpoints are).
  straightConnector1: (x, y, w, h) => `M${x},${y} L${x + w},${y + h}`,
  bentConnector2: (x, y, w, h) => `M${x},${y} L${x + w},${y} L${x + w},${y + h}`,
  bentConnector3: (x, y, w, h) =>
    `M${x},${y} L${x + w / 2},${y} L${x + w / 2},${y + h} L${x + w},${y + h}`,
  bentConnector4: (x, y, w, h) =>
    `M${x},${y} L${x + w * 0.33},${y} L${x + w * 0.33},${y + h * 0.5} L${x + w * 0.66},${y + h * 0.5} L${x + w * 0.66},${y + h} L${x + w},${y + h}`,
  bentConnector5: (x, y, w, h) =>
    `M${x},${y} L${x + w * 0.25},${y} L${x + w * 0.25},${y + h * 0.5} L${x + w * 0.75},${y + h * 0.5} L${x + w * 0.75},${y + h} L${x + w},${y + h}`,
  curvedConnector2: (x, y, w, h) => `M${x},${y} Q${x + w},${y} ${x + w},${y + h}`,
  curvedConnector3: (x, y, w, h) =>
    `M${x},${y} C${x + w * 0.5},${y} ${x + w * 0.5},${y + h} ${x + w},${y + h}`,
  curvedConnector4: (x, y, w, h) =>
    `M${x},${y} C${x + w * 0.33},${y} ${x + w * 0.33},${y + h * 0.5} ${x + w * 0.5},${y + h * 0.5} C${x + w * 0.66},${y + h * 0.5} ${x + w * 0.66},${y + h} ${x + w},${y + h}`,
  curvedConnector5: (x, y, w, h) =>
    `M${x},${y} C${x + w * 0.25},${y} ${x + w * 0.25},${y + h * 0.25} ${x + w * 0.5},${y + h * 0.5} C${x + w * 0.75},${y + h * 0.75} ${x + w * 0.75},${y + h} ${x + w},${y + h}`,
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

// PowerPoint's stock master defaults per placeholder type. Used when
// the run has no `<a:rPr sz=…>` of its own — @office-kit/pptx doesn't walk
// the lstStyle cascade to find the resolved size, so we mirror the
// well-known master defaults here.
const placeholderDefaultPt = (phType: string | null): number => {
  if (phType === 'title' || phType === 'ctrTitle') return DEFAULT_TITLE_PT; // 44pt
  if (phType === 'subTitle') return 32;
  if (phType === 'ftr' || phType === 'dt' || phType === 'sldNum') return 12;
  return DEFAULT_BODY_PT; // 18pt
};

const bulletChar = (level: number): string => (level <= 0 ? '•' : level === 1 ? '◦' : '▪');

// `text-shadow` takes one color per layer and no separate opacity, so an
// effect's `<a:alpha>` has to travel inside the color.
const cssColorWithOpacity = (hex: string, opacity: number | undefined): string => {
  if (opacity === undefined || opacity >= 1) return hex;
  const match = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(hex);
  if (!match) return hex;
  const [r, g, b] = [match[1]!, match[2]!, match[3]!].map((part) => Number.parseInt(part, 16));
  return `rgba(${r},${g},${b},${Math.max(0, opacity).toFixed(3)})`;
};

// `effectivePt` is the post-autofit font size in points. Callers pass
// `format.size` (the authored size, if any) scaled by the body's
// autofit factor, or the placeholder default scaled the same way.
const renderRun = (
  text: string,
  format: ReadTextFormat | null,
  theme: PresentationTheme | null,
  effectivePt: number,
  /* unused but kept for forward compatibility */ _wasDefault = false,
): string => {
  if (text === '') return '';
  void _wasDefault;
  const styles: string[] = [];
  styles.push(`font-size:${(effectivePt * PX_PER_PT).toFixed(2)}px`);
  // PowerPoint uses tight line-height (~1.0) by default for placeholders;
  // the previous 1.2 left enough vertical slack to push the top/bottom of
  // glyphs outside short placeholders.
  styles.push(`line-height:1.05`);
  if (format?.font) styles.push(`font-family:${escapeXml(format.font)}, ${DEFAULT_FONT}`);
  if (format?.bold) styles.push('font-weight:700');
  if (format?.italic) styles.push('font-style:italic');
  const underline = format?.underline;
  const strike = format?.strike;
  const hasUnderline = underline !== undefined && underline !== false && underline !== 'none';
  const hasStrike = strike !== undefined && strike !== false && strike !== 'noStrike';
  // Unlike the SVG path (no text-decoration-style support in resvg), real
  // browsers render `text-decoration-style:wavy` fine, so the foreignObject
  // path can use CSS directly instead of a hand-drawn path.
  const isWavyUnderline = typeof underline === 'string' && underline.startsWith('wavy');
  // text-decoration-style is a single value for the whole underline+
  // line-through shorthand, so a wavy underline combined with a strikethrough
  // on the same run would wave the strikethrough too — PowerPoint always
  // draws strikethrough solid regardless of the underline style. Nest a nowrap
  // inner span carrying just the wavy underline so each element's decoration
  // lines render independently.
  const nestedWavyUnderline = hasUnderline && hasStrike && isWavyUnderline;
  if (nestedWavyUnderline) {
    styles.push('text-decoration:line-through');
  } else if (hasUnderline && hasStrike) {
    styles.push('text-decoration:underline line-through');
  } else if (hasUnderline) {
    styles.push('text-decoration:underline');
    if (isWavyUnderline) styles.push('text-decoration-style:wavy');
  } else if (hasStrike) {
    styles.push('text-decoration:line-through');
  }
  if (format?.color !== undefined && format.color !== null) {
    styles.push(`color:${resolveColor(format.color, theme, '#000000')}`);
  }
  // S3 — additional rPr attributes that change how the glyphs lay out.
  if (format?.spc !== undefined && format.spc !== 0) {
    // ECMA-376 `spc` is in 1/100 pt; convert to CSS px at the run's size.
    const trackingPx = (format.spc / 100) * PX_PER_PT;
    styles.push(`letter-spacing:${trackingPx.toFixed(3)}px`);
  }
  if (format?.baseline !== undefined && format.baseline !== 0) {
    // Positive = superscript, negative = subscript. Scale the glyph a
    // little smaller as PowerPoint does (~64% for super/subscript).
    const direction = format.baseline > 0 ? 'super' : 'sub';
    styles.push(`vertical-align:${direction}`);
    styles.push('font-size:0.65em');
  }
  if (format?.cap === 'all') styles.push('text-transform:uppercase');
  else if (format?.cap === 'small') styles.push('font-variant:small-caps');
  if (format?.highlight !== undefined && format.highlight !== null) {
    styles.push(`background-color:${resolveColor(format.highlight, theme, '#FFFF00')}`);
  }
  // Character-level `<a:ln>` and `<a:effectLst>`. The outline is drawn behind
  // the glyph fill, the way PowerPoint draws it — `paint-order` is the SVG
  // spelling and `-webkit-text-stroke` the HTML one, and only the latter has
  // any effect inside a `<foreignObject>`.
  const outline = format?.outline;
  if (outline && outline.color) {
    const widthPx = (outline.widthEmu ?? 9525) / EMU_PER_PX;
    styles.push(
      `-webkit-text-stroke:${widthPx.toFixed(2)}px ${resolveColor(outline.color, theme, '#000000')}`,
    );
    styles.push('paint-order:stroke fill');
  }
  const textShadows: string[] = [];
  const glow = format?.glow;
  if (glow?.color) {
    const radiusPx = (glow.radiusEmu ?? 63500) / EMU_PER_PX;
    const color = cssColorWithOpacity(resolveColor(glow.color, theme, '#FFFF00'), glow.opacity);
    textShadows.push(`0 0 ${radiusPx.toFixed(2)}px ${color}`);
  }
  const shadow = format?.shadow;
  if (shadow) {
    const distancePx = (shadow.offsetEmu ?? 38100) / EMU_PER_PX;
    const radians = ((shadow.angleDeg ?? 45) * Math.PI) / 180;
    const blurPx = (shadow.blurEmu ?? 50800) / EMU_PER_PX;
    const color = cssColorWithOpacity(
      resolveColor(shadow.color ?? '#000000', theme, '#000000'),
      shadow.opacity,
    );
    textShadows.push(
      `${(Math.cos(radians) * distancePx).toFixed(2)}px ${(Math.sin(radians) * distancePx).toFixed(2)}px ${blurPx.toFixed(2)}px ${color}`,
    );
  }
  if (textShadows.length) styles.push(`text-shadow:${textShadows.join(',')}`);
  // Explicit `\n` in the run text comes from <a:br> line breaks; project
  // each to an HTML <br/> so the foreignObject's CSS layout honours it.
  // Everything else is escaped as XML text.
  const html = text
    .split('\n')
    .map((part) => escapeXml(part))
    .join('<br/>');
  const content = nestedWavyUnderline
    ? `<span style="text-decoration:underline;text-decoration-style:wavy">${html}</span>`
    : html;
  return `<span style="${styles.join(';')}">${content}</span>`;
};

// CSS line-height factor we render text at. Keep in sync with the
// `line-height` declaration in renderRun.
const LINE_HEIGHT = 1.05;

// `<a:normAutofit/>` shrink-to-fit search bounds. PowerPoint reduces the font
// in discrete steps until the body fits its box; we sweep from 1.0 down to the
// floor in fixed decrements. The floor stops a pathologically small box from
// collapsing text to an unreadable size (PowerPoint clamps similarly).
const AUTOFIT_FLOOR = 0.25;
const AUTOFIT_STEP = 0.05;

type RunData = {
  text: string;
  fmt: ReadTextFormat | null;
  sizePt: number;
  href?: string;
  hrefTip?: string;
};
interface ParaData {
  readonly align: string;
  readonly level: number;
  readonly bulletStyle: ReturnType<typeof getParagraphBullet>;
  readonly bulletDetail: ReturnType<typeof getParagraphBulletStyle>;
  readonly bulletIsPicture: boolean;
  // Data URL for a picture bullet (`<a:buBlip>`) whose bytes resolved,
  // else null — null falls back to the "■" glyph.
  readonly bulletImageHref: string | null;
  readonly runs: RunData[];
  readonly lineSpacing: ReturnType<typeof getParagraphLineSpacing>;
  readonly spcBefPts: number | null;
  readonly spcAftPts: number | null;
  readonly indent: ReturnType<typeof getParagraphIndent>;
}

// Collapses every ST_TextUnderlineType token onto the 3 styles the SVG text
// engine actually distinguishes (see PieceInput.underline): the wavy family
// (wavy/wavyDbl/wavyHeavy) needs a hand-drawn path since SVG has no
// text-decoration-style, so it can't share a bucket with plain/dashed/dotted
// styles, which all render fine as a single line.
const underlineStyleOf = (fmt: ReadTextFormat | null): 'none' | 'single' | 'wavy' => {
  const u = fmt?.underline;
  if (u === undefined || u === false || u === 'none') return 'none';
  if (typeof u === 'string' && u.startsWith('wavy')) return 'wavy';
  return 'single';
};
const hasStrikeFmt = (fmt: ReadTextFormat | null): boolean => {
  const s = fmt?.strike;
  return s !== undefined && s !== false && s !== 'noStrike';
};

export interface SvgTextArgs {
  readonly pres: PresentationData;
  readonly shape: SlideShapeData;
  readonly theme: PresentationTheme | null;
  readonly paraData: readonly ParaData[];
  readonly numberLabels: ReadonlyArray<string | null>;
  readonly autoFitScale: number;
  readonly lineHeightScale: number;
  readonly defaultPt: number;
  readonly themeFace: string | null;
  readonly defaultColor: string;
  readonly anchor: 'top' | 'center' | 'bottom';
  readonly anchorCentered?: boolean;
  readonly wrap: boolean;
  readonly innerX: number;
  readonly innerY: number;
  readonly innerW: number;
  readonly innerH: number;
  readonly measure: TextMeasurer;
  readonly vert: VerticalLayout;
  readonly columns: ColumnLayout | null;
  /** Maps an authored font name onto the family the measurer keys off.
   *  The render paths leave this unset (= `substituteFamily`, whose output
   *  must match the bundled TTFs' internal names for resvg). The audit path
   *  passes the authored name through so a fontkit measurer with
   *  user-registered fonts can resolve it before falling back to the
   *  substitution map. */
  readonly resolveFamily?: (family: string | null) => string;
}

const alignOf = (a: string): ParaInput['align'] =>
  a === 'center' || a === 'right' || a === 'justify' ? a : 'left';

// Map the OOXML `vert` token onto the engine's layout modes, mirroring the
// foreignObject path's writing-mode grouping so server == browser: the
// `vertical-rl` family (columns right-to-left) rotates 90° clockwise, the
// `vertical-lr` family (columns left-to-right) rotates 270°. `wordArtVert`
// (ST_TextVerticalType, §21.1.2.1.1) means "one letter on top of another" —
// glyphs stay UPRIGHT and stack, not rotated — matching the browser path's
// `text-orientation:upright` and PowerPoint/LibreOffice.
//
// KNOWN DIVERGENCE (overflow only): when an upright run is taller than its box,
// the browser path wraps the overflow into a second column (CSS writing-mode),
// while the SVG path stacks it as one column that clips at the box edge. Both
// agree for text that fits; they differ only past the box, where the deck is
// already out of spec. Faithful multi-column upright wrapping in the SVG engine
// is disproportionate to that edge case, so we accept the clip.
export const verticalLayoutOf = (
  vert: ReturnType<typeof getShapeTextDirection>,
): VerticalLayout => {
  switch (vert) {
    case 'vert':
    case 'eaVert':
      return 'cw90';
    case 'wordArtVert':
    case 'wordArtVertRtl':
      return 'upright';
    case 'vert270':
    case 'mongolianVert':
      return 'cw270';
    case null:
      return 'none';
  }
};

// Build the px-native engine input from the resolved paraData (at a.autoFitScale).
export const buildSvgTextInput = (a: SvgTextArgs): TextBodyInput => {
  const scale = a.autoFitScale;
  const paragraphs: ParaInput[] = a.paraData.map((para, pi): ParaInput => {
    const pieces: PieceInput[] = [];
    for (const run of para.runs) {
      // <a:br> marker.
      if (run.text === '\n' && run.fmt === null) {
        pieces.push(breakPiece());
        continue;
      }
      let fmt = run.fmt;
      if (run.href) {
        const hlinkColor = a.theme ? normalizeHex(a.theme.hyperlink) : '#0563C1';
        fmt = {
          ...fmt,
          // PowerPoint and LibreOffice override a hyperlink run's direct
          // solidFill with the theme hlink color — the explicit run color
          // does NOT win for a link run, so the theme color is unconditional.
          color: hlinkColor,
          underline: fmt?.underline ?? true,
        };
      }
      const family = (a.resolveFamily ?? substituteFamily)(fmt?.font ?? a.themeFace);
      const sizePx = run.sizePt * scale * PX_PER_PT;
      const fillHex =
        fmt?.color !== undefined && fmt.color !== null
          ? resolveColor(fmt.color, a.theme, '#000000')
          : a.defaultColor;
      const superSub: 0 | 1 | -1 =
        fmt?.baseline !== undefined && fmt.baseline !== 0 ? (fmt.baseline > 0 ? 1 : -1) : 0;
      const letterSpacingPx =
        fmt?.spc !== undefined && fmt.spc !== 0 ? (fmt.spc / 100) * PX_PER_PT : 0;
      const caps = fmt?.cap === 'all' || fmt?.cap === 'small';
      const base: Omit<PieceInput, 'text' | 'isBreak'> = {
        family,
        sizePx,
        bold: fmt?.bold ?? false,
        italic: fmt?.italic ?? false,
        letterSpacingPx,
        fillHex,
        underline: underlineStyleOf(fmt),
        strike: hasStrikeFmt(fmt),
        ...(fmt?.highlight
          ? { highlightHex: resolveColor(fmt.highlight, a.theme, '#FFFF00') }
          : {}),
        ...(fmt?.outline?.color
          ? {
              outlineHex: resolveColor(fmt.outline.color, a.theme, '#000000'),
              outlineWidthPx: (fmt.outline.widthEmu ?? 9525) / EMU_PER_PX,
            }
          : {}),
        superSub,
        href: run.href ?? null,
        ...(run.hrefTip !== undefined ? { hrefTip: run.hrefTip } : {}),
      };
      // A run's text can carry embedded '\n' only via <a:br>, already split
      // out above; still split defensively so any stray newline becomes a break.
      const segs = run.text.split('\n');
      segs.forEach((seg, i) => {
        const segText = caps ? seg.toUpperCase() : seg;
        if (a.vert === 'upright') {
          // wordArtVert stacks one code point per line (spaces become blank
          // rows). Reuse the calibrated horizontal engine with a break before
          // every glyph rather than rotating the run. The break goes BEFORE each
          // glyph (skipped only at the very start or right after an existing
          // break) so adjacent runs stack too — a trailing per-glyph break would
          // leave the last glyph of run N sharing a row with the first of run
          // N+1, diverging from the browser's text-orientation:upright path.
          for (const g of Array.from(segText)) {
            const last = pieces[pieces.length - 1];
            if (last !== undefined && !last.isBreak) pieces.push(breakPiece());
            pieces.push({ ...base, text: g, isBreak: false });
          }
        } else {
          pieces.push({ ...base, text: segText, isBreak: false });
        }
        if (i < segs.length - 1) pieces.push(breakPiece());
      });
    }

    const marLpx =
      para.indent.leftEmu !== null
        ? (para.indent.leftEmu / EMU_PER_PX) * scale
        : para.level > 0
          ? para.level * 24 * PX_PER_PT * scale
          : 0;
    const marRpx = para.indent.rightEmu !== null ? (para.indent.rightEmu / EMU_PER_PX) * scale : 0;
    const firstIndentPx =
      para.indent.firstLineEmu !== null ? (para.indent.firstLineEmu / EMU_PER_PX) * scale : 0;
    const spcBefPx =
      para.spcBefPts !== null && para.spcBefPts > 0 ? para.spcBefPts * PX_PER_PT * scale : 0;
    const spcAftPx =
      para.spcAftPts !== null && para.spcAftPts > 0 ? para.spcAftPts * PX_PER_PT * scale : 0;
    const lineSpacing: ParaInput['lineSpacing'] =
      para.lineSpacing?.kind === 'pct'
        ? { kind: 'pct', value: para.lineSpacing.value }
        : para.lineSpacing?.kind === 'pts'
          ? { kind: 'pts', px: para.lineSpacing.value * PX_PER_PT * scale }
          : null;

    return {
      align: alignOf(para.align),
      marLpx,
      marRpx,
      firstIndentPx,
      spcBefPx,
      spcAftPx,
      lineSpacing,
      lineAdvanceScale: a.lineHeightScale,
      bullet: buildBullet(a, para, pi),
      pieces,
      fallbackSizePx: a.defaultPt * scale * PX_PER_PT,
    };
  });

  const input: TextBodyInput = {
    boxXpx: a.innerX / EMU_PER_PX,
    boxYpx: a.innerY / EMU_PER_PX,
    boxWpx: a.innerW / EMU_PER_PX,
    boxHpx: a.innerH / EMU_PER_PX,
    anchor: a.anchor,
    anchorCentered: a.anchorCentered ?? false,
    wrap: a.wrap,
    paragraphs,
    vert: a.vert,
    columns: a.columns,
  };
  return input;
};

const buildAndLayoutSvgText = (a: SvgTextArgs): string =>
  layoutTextSvg(buildSvgTextInput(a), a.measure);

const breakPiece = (): PieceInput => ({
  text: '',
  family: '',
  sizePx: 0,
  bold: false,
  italic: false,
  letterSpacingPx: 0,
  fillHex: '#000000',
  underline: 'none',
  strike: false,
  superSub: 0,
  href: null,
  isBreak: true,
});

const buildBullet = (a: SvgTextArgs, para: ParaData, pi: number): BulletInput | null => {
  const explicitChar =
    para.bulletStyle !== null && typeof para.bulletStyle === 'object' && 'char' in para.bulletStyle
      ? para.bulletStyle.char
      : null;
  const numberLabel = a.numberLabels[pi] ?? null;
  const showBullet =
    para.bulletStyle === 'bullet' ||
    explicitChar !== null ||
    numberLabel !== null ||
    para.bulletIsPicture ||
    (para.bulletStyle !== 'none' && para.level > 0);
  if (!showBullet) return null;
  // Picture bullet with resolved bytes → render the image; otherwise the
  // "■" glyph stands in (bytes unavailable, e.g. inherited buBlip).
  const char = para.bulletIsPicture ? '■' : (numberLabel ?? explicitChar ?? bulletChar(para.level));
  // An un-sized bullet (no buSzPct/buSzPts in the cascade) is 100% of the
  // paragraph's FIRST RUN size, not the placeholder default — a 32pt run gets a
  // 32pt bullet even on an 18pt-default body. Fall back to defaultPt only for an
  // empty paragraph (break-only) where there is no run size.
  const firstRunPt = para.runs.find((r) => r.text !== '\n' && r.text !== '')?.sizePt ?? a.defaultPt;
  const baseSizePx = firstRunPt * PX_PER_PT * a.autoFitScale;
  const sizePx =
    para.bulletDetail.sizePct !== null
      ? baseSizePx * para.bulletDetail.sizePct
      : para.bulletDetail.sizePts !== null
        ? para.bulletDetail.sizePts * PX_PER_PT * a.autoFitScale
        : baseSizePx;
  const fillHex = bulletFillOf(para, a.theme, a.defaultColor);
  return {
    text: char,
    family: (a.resolveFamily ?? substituteFamily)(para.bulletDetail.font ?? DEFAULT_BULLET_FONT),
    sizePx,
    fillHex,
    ...(para.bulletImageHref ? { imageHref: para.bulletImageHref } : {}),
  };
};

// A bullet with no `<a:buClr>` anywhere in the cascade takes the colour of the
// paragraph's FIRST RUN, the same rule an un-sized bullet already follows for
// its size. Falling back to the body default instead paints the marker black on
// a deck whose runs carry their own light colour, which is what a dark-themed
// agenda slide with `<a:buAutoNum>` looks like when the numbers go missing.
const bulletFillOf = (
  para: ParaData,
  theme: PresentationTheme | null,
  defaultColor: string,
): string => {
  if (para.bulletDetail.color) return resolveColor(para.bulletDetail.color, theme, '#000000');
  const firstRun = para.runs.find((r) => r.text !== '\n' && r.text !== '');
  const runColor = firstRun?.fmt?.color;
  return runColor === undefined || runColor === null
    ? defaultColor
    : resolveColor(runColor, theme, defaultColor);
};

// Preset geometry text rectangle (ECMA-376 `<a:rect>`), as fractions of w/h.
// The render-path-independent half of a shape's text body: the resolved
// paragraph/run model, the effective bodyPr cascade, the inner text rect, and
// the final autofit factor. Extracted from renderTextBody so the audit API
// (`auditTextLayout`) measures with EXACTLY the pipeline the renderer draws
// with — any drift between the two would make the audit lie about the preview.
export interface TextBodyModel {
  readonly paraData: readonly ParaData[];
  readonly numberLabels: ReadonlyArray<string | null>;
  readonly authoredAutofit: ReturnType<typeof getShapeTextAutoFitParams>;
  /** Final autofit factor: the baked `fontScale`, the bare-normAutofit shrink
   *  search result, or the foreignObject-only heuristic. SVG-semantics callers
   *  (the fidelity path, the audit) must apply it only when `authoredAutofit`
   *  is set — see the render path's `svgScale`. */
  readonly autoFitScale: number;
  readonly lineHeightScale: number;
  readonly defaultPt: number;
  readonly themeFace: string | null;
  readonly effectiveDefaultFont: string;
  readonly effectiveBody: ReturnType<typeof getShapeBodyPrEffective>;
  readonly anchor: 'top' | 'center' | 'bottom';
  readonly anchorOffset: { readonly x: number; readonly y: number };
  /** Inner text rect in EMU (preset-geometry text rect + insets applied). */
  readonly innerX: number;
  readonly innerY: number;
  readonly innerW: number;
  readonly innerH: number;
  readonly svgTextRect: (v: VerticalLayout) => { x: number; y: number; w: number; h: number };
}

// Returns null when the shape has no text body worth laying out (no
// paragraphs, no run text, or a degenerate inner rect).
export const resolveTextBodyModel = (
  pres: PresentationData,
  shape: SlideShapeData,
  bounds: { x: number; y: number; w: number; h: number },
  theme: PresentationTheme | null,
  phType: string | null,
  measure: TextMeasurer,
  defaultColor: string,
): TextBodyModel | null => {
  let paragraphCount: number;
  try {
    paragraphCount = getShapeParagraphCount(shape);
  } catch {
    return null;
  }
  if (paragraphCount === 0) return null;

  const defaultPt = placeholderDefaultPt(phType);
  // Theme font stack — `<a:fontScheme><a:majorFont>` is the title face,
  // `<a:minorFont>` is the body face. Use the major face for title /
  // ctrTitle placeholders, minor for everything else. Falls through to
  // DEFAULT_FONT so missing themes still render with our generic stack.
  const themeFonts = getPresentationFonts(pres);
  const isTitlePlaceholder = phType === 'title' || phType === 'ctrTitle';
  const themeFace = isTitlePlaceholder
    ? (themeFonts?.majorLatin ?? null)
    : (themeFonts?.minorLatin ?? null);
  const effectiveDefaultFont = themeFace
    ? `${escapeXml(themeFace)}, ${DEFAULT_FONT}`
    : DEFAULT_FONT;
  // Effective body-property cascade — anchor, wrap, vert, margins all
  // inherit from the layout / master placeholder when the slide doesn't
  // override them.
  let effectiveBody: ReturnType<typeof getShapeBodyPrEffective>;
  try {
    effectiveBody = getShapeBodyPrEffective(pres, shape);
  } catch {
    effectiveBody = {
      anchor: getShapeTextAnchor(shape),
      anchorCentered: null,
      autoFit: getShapeTextAutoFit(shape),
      autoFitParams: getShapeTextAutoFitParams(shape),
      columns: getShapeTextColumns(shape),
      wrap: null,
      vert: getShapeTextDirection(shape),
      margins: getShapeTextMargins(shape) ?? { left: null, top: null, right: null, bottom: null },
    };
  }
  // Default text alignment depends on the shape kind, matching PowerPoint /
  // LibreOffice: autoshapes (preset geometry that is neither a placeholder nor
  // a text box) center text horizontally and anchor it middle by default,
  // whereas text boxes and placeholder bodies default to left / top. Authored
  // algn / anchor still win.
  const isAutoshape = !isShapePlaceholder(shape) && !isShapeTextBox(shape);
  const defaultAlign = isAutoshape ? 'center' : 'left';
  const defaultAnchor: 'top' | 'center' | 'bottom' = isAutoshape ? 'center' : 'top';
  const anchor = effectiveBody.anchor ?? defaultAnchor;
  const margins = effectiveBody.margins;
  const lIns = margins.left ?? DEFAULT_INSET_X;
  const tIns = margins.top ?? DEFAULT_INSET_Y;
  const rIns = margins.right ?? DEFAULT_INSET_X;
  const bIns = margins.bottom ?? DEFAULT_INSET_Y;

  const {
    x: innerX,
    y: innerY,
    w: innerW,
    h: innerH,
  } = resolveTextBodyRect(
    getShapePreset(shape),
    bounds,
    { left: lIns, top: tIns, right: rIns, bottom: bIns },
    // A custom shape states where its text goes; only a preset has to be
    // approximated from a table. The rect is in the same EMU space as the
    // shape's own `<a:ext>`, which is what turns it into fractions here —
    // `bounds` carries a group's scale and would divide it away twice.
    shapeCustomTextRect(getShapeCustomGeometry(shape), getShapeBounds(shape)),
  );
  if (innerW <= 0 || innerH <= 0) return null;

  // The rect the pure-SVG path lays text into for a given vertical layout.
  // Horizontal and `upright` (wordArtVert) text use the shared inner rect; the
  // true ±90° rotations (cw90/cw270) swap the box extents and rotate the insets
  // with the glyphs — the reading/wrap dimension takes the L/R insets, the
  // column-stack dimension the T/B insets. Shared by the normAutofit shrink
  // search and the SVG render so the shrink is measured against the SAME box the
  // text is rendered into (with PowerPoint's asymmetric default insets the two
  // rects differ by ~9.6px/axis, which would otherwise mis-size the fit).
  const svgTextRect = (v: VerticalLayout): { x: number; y: number; w: number; h: number } =>
    v === 'none' || v === 'upright'
      ? { x: innerX, y: innerY, w: innerW, h: innerH }
      : {
          x: bounds.x + tIns,
          y: bounds.y + lIns,
          w: Math.max(0, bounds.w - tIns - bIns),
          h: Math.max(0, bounds.h - lIns - rIns),
        };

  // First pass — collect every run's text + format so we can both
  // (a) compute an autofit scale and (b) emit each run with the
  // adjusted size. RunData / ParaData are module-scoped (above) so the
  // pure-SVG path can consume the same resolved model.
  const paraData: ParaData[] = [];
  let hasAnyText = false;
  for (let p = 0; p < paragraphCount; p++) {
    // Resolve paragraph properties through the layout/master cascade so
    // placeholders inherit their default alignment / line-spacing / indent
    // without each slide having to repeat them. Fall back to the literal
    // readers if the cascade throws (very defensive — should never happen).
    let effective: ReturnType<typeof getParagraphPropertiesEffective>;
    try {
      effective = getParagraphPropertiesEffective(pres, shape, p);
    } catch {
      effective = {
        align: getParagraphAlignment(shape, p),
        level: getParagraphLevel(shape, p),
        marL: null,
        marR: null,
        indent: null,
        lineSpacing: null,
        spcBefPts: null,
        spcAftPts: null,
        rtl: null,
        bullet: null,
      };
    }
    const align = effective.align ?? defaultAlign;
    const level = effective.level;
    // The literal slide-level bullet wins; otherwise the cascade supplies the
    // inherited one (master bodyStyle authors "•" for body placeholders).
    const bulletStyle = getParagraphBullet(shape, p) ?? effective.bullet;
    let bulletDetail: ReturnType<typeof getParagraphBulletStyle> = {
      color: null,
      sizePct: null,
      sizePts: null,
      font: null,
    };
    try {
      bulletDetail = getParagraphBulletStyle(pres, shape, p);
    } catch {}
    let bulletIsPicture = false;
    try {
      bulletIsPicture = isParagraphBulletPicture(shape, p);
    } catch {}
    let bulletImageHref: string | null = null;
    if (bulletIsPicture) {
      try {
        const bytes = getParagraphBulletImageBytes(shape, p);
        if (bytes) bulletImageHref = bytesToDataUrl(bytes);
      } catch {}
    }
    const lineSpacing: ReturnType<typeof getParagraphLineSpacing> = effective.lineSpacing;
    let spcBefPts: number | null = effective.spcBefPts;
    let spcAftPts: number | null = effective.spcAftPts;
    const indent: ReturnType<typeof getParagraphIndent> = {
      leftEmu: effective.marL,
      rightEmu: effective.marR,
      firstLineEmu: effective.indent,
    };
    // Literal pPr values override the cascade so any per-slide override
    // wins on this paragraph.
    try {
      const spacing = getParagraphSpacing(shape, p);
      if (spacing.beforePts !== null) spcBefPts = spacing.beforePts;
      if (spacing.afterPts !== null) spcAftPts = spacing.afterPts;
    } catch {}
    try {
      const lit = getParagraphIndent(shape, p);
      if (lit.leftEmu !== null) indent.leftEmu = lit.leftEmu;
      if (lit.rightEmu !== null) indent.rightEmu = lit.rightEmu;
      if (lit.firstLineEmu !== null) indent.firstLineEmu = lit.firstLineEmu;
    } catch {}
    const runs: RunData[] = [];
    // Walk the paragraph's inline elements — runs, field placeholders,
    // and explicit line breaks — in document order. The strict
    // <a:r>-only `getShapeRunCount` would silently skip footer / date /
    // slide-number fields, which is exactly what real templates emit.
    let elements: ReadonlyArray<ReturnType<typeof getShapeParagraphElements>[number]> = [];
    try {
      elements = getShapeParagraphElements(shape, p);
    } catch {
      elements = [];
    }
    let rIdx = 0;
    for (const el of elements) {
      if (el.kind === 'br') {
        runs.push({ text: '\n', fmt: null, sizePt: defaultPt });
        continue;
      }
      // A `slidenum` field shows the slide's own number, which PowerPoint
      // recomputes on open; the cached `<a:t>` is whatever it last wrote, and
      // is empty for a field the deck just gained. Every other field type
      // keeps its cached text — `datetime` in particular has thirteen
      // locale-dependent variants that a preview should not guess at.
      const txt = el.kind === 'fld' && el.type === 'slidenum' ? activeSlideNumber : el.text;
      let fmt: ReadTextFormat | null = el.format;
      let href: string | undefined;
      let hrefTip: string | undefined;
      if (el.kind === 'r') {
        // The cascade only makes sense for actual <a:r> runs; field
        // text is opaque cached content and shouldn't pretend to be a
        // specific run index.
        try {
          fmt = getShapeRunFormatEffective(pres, shape, p, rIdx);
        } catch {
          fmt = getShapeRunFormat(shape, p, rIdx);
        }
        try {
          href = getShapeRunHyperlink(shape, p, rIdx) ?? undefined;
          // Per-run slide-jump actions (`<a:hlinkClick action=
          // "ppaction://hlinksldjump"/>`) resolve to an in-page anchor.
          // Fall back to them only when no external URL was authored.
          if (!href) {
            const act = getShapeRunClickAction(shape, p, rIdx);
            href = clickActionHref(pres, shape, act);
          }
          if (href) hrefTip = getShapeRunHyperlinkTooltip(shape, p, rIdx) ?? undefined;
        } catch {
          href = undefined;
        }
        rIdx++;
      }
      const sizePt = fmt?.size ?? defaultPt;
      if (txt) hasAnyText = true;
      runs.push({
        text: txt,
        fmt,
        sizePt,
        ...(href !== undefined ? { href } : {}),
        ...(hrefTip !== undefined ? { hrefTip } : {}),
      });
    }
    paraData.push({
      align,
      level,
      bulletStyle,
      bulletDetail,
      bulletIsPicture,
      bulletImageHref,
      runs,
      lineSpacing,
      spcBefPts,
      spcAftPts,
      indent,
    });
  }
  if (!hasAnyText) return null;

  // Prefer the *authored* autofit factor when PowerPoint already
  // computed one (`<a:normAutofit fontScale=…/>`). That's the same
  // multiplier PowerPoint applies on-screen, so honouring it is what
  // brings the rendered size into 1:1 agreement with the deck.
  //
  // No estimation happens for shapes WITHOUT `<a:normAutofit>`: PowerPoint
  // never shrinks those — `<a:noAutofit>` (and no autofit at all) simply
  // overflows the box, and `<a:spAutoFit>` grows the box instead of the text.
  // An earlier heuristic shrank such shapes to fit their authored box, which
  // rendered template placeholders (size inherited from layout/master, box
  // sized by the template author) at up to 0.4× of their PowerPoint size.
  const authoredAutofit = effectiveBody.autoFitParams;
  let autoFitScale = authoredAutofit?.fontScale ?? 1;
  const lineHeightScale = 1 - (authoredAutofit?.lnSpcReduction ?? 0);

  // Apply line-height reduction by tightening per-line spacing. We
  // pass it through to renderRun via a closed-over factor.
  const effectiveLineHeight = LINE_HEIGHT * lineHeightScale;
  void effectiveLineHeight; // currently unused — kept for forward compat

  const numberLabels = paragraphNumberLabels(paraData);

  // A bare `<a:normAutofit/>` (no baked `fontScale`, so it defaults to 1) means
  // "shrink text to fit the box" — PowerPoint computes that reduction at display
  // time. Compute it ONCE here, before either render path runs, so the
  // foreignObject (browser) and pure-SVG (server) paths shrink by the SAME
  // factor: the server==browser parity this preview maintains. An explicit baked
  // fontScale (autoFitScale !== 1) is honoured unchanged; noAutofit / spAutoFit
  // yield no authoredAutofit and never shrink here. The shrink is measured with
  // the SVG layout engine (our only real line-breaker); the foreignObject path
  // then reuses the resulting scale rather than computing its own.
  // Mac PowerPoint opens placeholders with an inherited bare normAutofit at
  // their authored size (for example the two-line title in 01-title-only).
  // Inheritance exposes the editing policy, but is not a saved shrink request.
  // Keep inherited baked scales above; estimate only a shape-local request.
  if (authoredAutofit && autoFitScale === 1 && getShapeTextAutoFit(shape) === 'normal') {
    const fitVert = verticalLayoutOf(effectiveBody.vert ?? getShapeTextDirection(shape));
    const fitCols = effectiveBody.columns;
    const fitColumns: ColumnLayout | null =
      fitVert === 'none' && fitCols && fitCols.count >= 2
        ? {
            count: fitCols.count,
            gapPx: fitCols.gapEmu !== undefined ? fitCols.gapEmu / EMU_PER_PX : 0,
          }
        : null;
    // Measure against the SAME rect the SVG path renders into (rotated text uses
    // the inset-swapped box). For cw90/cw270 the engine swaps the box extents, so
    // the measured `requiredH` (stacking dimension) fills the box WIDTH, not its
    // height — compare against the rect's w there, its h otherwise.
    const fitRect = svgTextRect(fitVert);
    const fitRotated = fitVert === 'cw90' || fitVert === 'cw270';
    const fitBoxPx = (fitRotated ? fitRect.w : fitRect.h) / EMU_PER_PX;
    const fitArgsBase: Omit<SvgTextArgs, 'autoFitScale'> = {
      pres,
      shape,
      theme,
      paraData,
      numberLabels,
      lineHeightScale,
      defaultPt,
      themeFace,
      defaultColor,
      anchor: anchor === 'center' || anchor === 'bottom' ? anchor : 'top',
      anchorCentered: effectiveBody.anchorCentered ?? false,
      wrap: effectiveBody.wrap !== 'none',
      innerX: fitRect.x,
      innerY: fitRect.y,
      innerW: fitRect.w,
      innerH: fitRect.h,
      measure,
      vert: fitVert,
      columns: fitColumns,
    };
    let s = 1;
    while (s > AUTOFIT_FLOOR) {
      if (
        measureTextBodyHeight(buildSvgTextInput({ ...fitArgsBase, autoFitScale: s }), measure) <=
        fitBoxPx
      )
        break;
      s -= AUTOFIT_STEP;
    }
    autoFitScale = Math.max(AUTOFIT_FLOOR, s);
  }

  let anchorOffset = { x: 0, y: 0 };
  if (effectiveBody.anchorCentered) {
    const vert = verticalLayoutOf(effectiveBody.vert ?? getShapeTextDirection(shape));
    const rect = svgTextRect(vert);
    const cols = effectiveBody.columns;
    const input = buildSvgTextInput({
      pres,
      shape,
      theme,
      paraData,
      numberLabels,
      lineHeightScale,
      defaultPt,
      themeFace,
      defaultColor,
      anchor,
      anchorCentered: true,
      wrap: effectiveBody.wrap !== 'none',
      innerX: rect.x,
      innerY: rect.y,
      innerW: rect.w,
      innerH: rect.h,
      measure,
      vert,
      autoFitScale,
      columns:
        vert === 'none' && cols && cols.count >= 2
          ? { count: cols.count, gapPx: cols.gapEmu !== undefined ? cols.gapEmu / EMU_PER_PX : 0 }
          : null,
    });
    const { anchorShift } = layoutCore(input, measure);
    anchorOffset =
      vert === 'cw90'
        ? { x: 0, y: anchorShift }
        : vert === 'cw270'
          ? { x: 0, y: -anchorShift }
          : { x: anchorShift, y: 0 };
  }

  return {
    paraData,
    numberLabels,
    authoredAutofit,
    anchorOffset,
    autoFitScale,
    lineHeightScale,
    defaultPt,
    themeFace,
    effectiveDefaultFont,
    effectiveBody,
    anchor,
    innerX,
    innerY,
    innerW,
    innerH,
    svgTextRect,
  };
};

const renderHtmlParagraphs = (
  paraData: ReadonlyArray<ParaData>,
  numberLabels: ReadonlyArray<string | null>,
  theme: PresentationTheme | null,
  autoFitScale: number,
  defaultPt: number,
  defaultColor: string,
): string[] => {
  // Second pass — emit runs with scaled sizes.
  const paragraphs: string[] = [];
  for (let pi = 0; pi < paraData.length; pi++) {
    const para = paraData[pi]!;
    const runHtmls = para.runs.map((run) => {
      // Per-run hyperlinks render the text in the theme's hyperlink
      // color (with underline) and wrap the span in an <a href> so the
      // preview is clickable.
      let runFmt = run.fmt;
      if (run.href) {
        const hlinkColor = theme ? normalizeHex(theme.hyperlink) : '#0563C1';
        runFmt = {
          ...runFmt,
          // Theme hlink color overrides a hyperlink run's direct fill (see
          // the SVG path above) — match PowerPoint / LibreOffice.
          color: hlinkColor,
          underline: runFmt?.underline ?? true,
        };
      }
      const span = renderRun(
        run.text,
        runFmt,
        theme,
        run.sizePt * autoFitScale,
        run.fmt?.size === undefined,
      );
      if (!run.href) return span;
      const isInPage = run.href.startsWith('#');
      const targetAttrs = isInPage ? '' : ' target="_blank" rel="noopener noreferrer"';
      const titleAttr = run.hrefTip ? ` title="${escapeXml(run.hrefTip)}"` : '';
      return `<a href="${escapeXml(run.href)}"${targetAttrs}${titleAttr} style="color:inherit;text-decoration:inherit">${span}</a>`;
    });
    // <a:lnSpc> — paragraph line spacing. spcPct multiplies, spcPts
    // sets a fixed point value. CSS line-height accepts both forms;
    // we project pts to px at the run's authored size.
    let lineHeightCss = '';
    if (para.lineSpacing?.kind === 'pct') {
      lineHeightCss = `line-height:${para.lineSpacing.value.toFixed(3)}`;
    } else if (para.lineSpacing?.kind === 'pts') {
      lineHeightCss = `line-height:${(para.lineSpacing.value * PX_PER_PT * autoFitScale).toFixed(2)}px`;
    }
    // <a:spcBef> / <a:spcAft> map to CSS margin-top / margin-bottom.
    const marginTopCss =
      para.spcBefPts !== null && para.spcBefPts > 0
        ? `margin-top:${(para.spcBefPts * PX_PER_PT * autoFitScale).toFixed(2)}px`
        : '';
    const marginBottomCss =
      para.spcAftPts !== null && para.spcAftPts > 0
        ? `margin-bottom:${(para.spcAftPts * PX_PER_PT * autoFitScale).toFixed(2)}px`
        : '';
    // <a:pPr marL marR indent> → CSS padding-left / padding-right /
    // text-indent. Authored indents override the level-based default
    // so paragraphs with explicit marL don't get doubled.
    const leftPx =
      para.indent.leftEmu !== null
        ? (para.indent.leftEmu / EMU_PER_PX) * autoFitScale
        : para.level > 0
          ? para.level * 24 * PX_PER_PT * autoFitScale
          : 0;
    const rightPx =
      para.indent.rightEmu !== null ? (para.indent.rightEmu / EMU_PER_PX) * autoFitScale : 0;
    const firstLinePx =
      para.indent.firstLineEmu !== null
        ? (para.indent.firstLineEmu / EMU_PER_PX) * autoFitScale
        : 0;
    const pStyles: string[] = [
      'margin:0',
      marginTopCss,
      marginBottomCss,
      'padding:0',
      `text-align:${ALIGNMENT_TO_CSS[para.align] ?? 'left'}`,
      lineHeightCss,
      leftPx > 0 ? `padding-left:${leftPx.toFixed(2)}px` : '',
      rightPx > 0 ? `padding-right:${rightPx.toFixed(2)}px` : '',
      firstLinePx !== 0 ? `text-indent:${firstLinePx.toFixed(2)}px` : '',
    ].filter(Boolean);
    let prefix = '';
    const explicitChar =
      para.bulletStyle !== null &&
      typeof para.bulletStyle === 'object' &&
      'char' in para.bulletStyle
        ? para.bulletStyle.char
        : null;
    const numberLabel = numberLabels[pi];
    const showBullet =
      para.bulletStyle === 'bullet' ||
      explicitChar !== null ||
      numberLabel !== null ||
      para.bulletIsPicture ||
      (para.bulletStyle !== 'none' && para.level > 0);
    // An un-sized bullet is 100% of the paragraph's first-run size (not the
    // placeholder default) — must match buildBullet on the SVG path so the
    // browser preview and the rasterized SVG agree.
    const firstRunBulletPt =
      para.runs.find((r) => r.text !== '\n' && r.text !== '')?.sizePt ?? defaultPt;
    const baseBulletPx = firstRunBulletPt * PX_PER_PT * autoFitScale;
    if (showBullet && para.bulletImageHref) {
      // Picture bullet with resolved bytes — inline it as an <img> sized
      // to the bullet's font box so it sits where the glyph would.
      const bulletPx =
        para.bulletDetail.sizePct !== null
          ? baseBulletPx * para.bulletDetail.sizePct
          : para.bulletDetail.sizePts !== null
            ? para.bulletDetail.sizePts * PX_PER_PT * autoFitScale
            : baseBulletPx;
      const imgStyles = [
        `width:${bulletPx.toFixed(2)}px`,
        `height:${bulletPx.toFixed(2)}px`,
        'display:inline-block',
        'vertical-align:baseline',
        `margin-right:${(0.4 * defaultPt * PX_PER_PT * autoFitScale).toFixed(2)}px`,
      ];
      prefix = `<img src="${para.bulletImageHref}" alt="" style="${imgStyles.join(';')}"/>`;
    } else if (showBullet) {
      // Picture bullet whose bytes are unavailable falls back to a filled
      // square so users see a distinct glyph rather than the round bullet.
      const char = para.bulletIsPicture
        ? '■'
        : (numberLabel ?? explicitChar ?? bulletChar(para.level));
      // Bullet style overrides: color, size %, fixed pt size, font face.
      const bulletStyles: string[] = [
        `margin-right:${(0.4 * defaultPt * PX_PER_PT * autoFitScale).toFixed(2)}px`,
      ];
      bulletStyles.push(`color:${bulletFillOf(para, theme, defaultColor)}`);
      if (para.bulletDetail.sizePct !== null) {
        bulletStyles.push(`font-size:${(para.bulletDetail.sizePct * 100).toFixed(1)}%`);
      } else if (para.bulletDetail.sizePts !== null) {
        bulletStyles.push(
          `font-size:${(para.bulletDetail.sizePts * PX_PER_PT * autoFitScale).toFixed(2)}px`,
        );
      } else {
        // No authored size → the first-run size, matching the SVG path.
        bulletStyles.push(`font-size:${baseBulletPx.toFixed(2)}px`);
      }
      bulletStyles.push(
        para.bulletDetail.font
          ? `font-family:${escapeXml(para.bulletDetail.font)}, ${DEFAULT_FONT}`
          : `font-family:${DEFAULT_BULLET_FONT}, ${DEFAULT_FONT}`,
      );
      prefix = `<span style="${bulletStyles.join(';')}">${escapeXml(char)}</span>`;
    }
    // The index is the one `<p:bldP build="p">` counts in: a paragraph build
    // reveals `<a:p>` number N, and a player needs to find it in either text
    // path without re-reading the deck.
    paragraphs.push(
      `<p data-pptx-paragraph="${pi}" style="${pStyles.join(';')}">${prefix}${runHtmls.join('') || '&#8203;'}</p>`,
    );
  }

  return paragraphs;
};

const renderTextBody = (
  pres: PresentationData,
  shape: SlideShapeData,
  bounds: { x: number; y: number; w: number; h: number },
  theme: PresentationTheme | null,
  phType: string | null,
  ctx: LayoutCtx,
): string => {
  const model = resolveTextBodyModel(
    pres,
    shape,
    bounds,
    theme,
    phType,
    ctx.measure,
    activeDeckTextColor,
  );
  if (model === null) return '';
  const {
    paraData,
    numberLabels,
    authoredAutofit,
    autoFitScale,
    defaultPt,
    themeFace,
    effectiveDefaultFont,
    effectiveBody,
    anchor,
    innerX,
    innerY,
    innerW,
    innerH,
    svgTextRect,
  } = model;

  // Fallback color for runs with no authored color — the deck's body-text color
  // (master bodyStyle), not the `tx1` token, which an inverted map paints white.
  // Bullets fall back through it too, so it has to be in hand before the loop.
  const defaultColor = activeDeckTextColor;

  const paragraphs = renderHtmlParagraphs(
    paraData,
    numberLabels,
    theme,
    autoFitScale,
    defaultPt,
    defaultColor,
  );

  const justify = ANCHOR_TO_CSS[anchor] ?? 'flex-start';

  // Pure-SVG text path (browser-free rasterization). Rebuild the px-native
  // layout model from the already-resolved paraData and hand it to the engine,
  // matching the foreignObject path's vertical-text and multi-column handling so
  // server-side rendering agrees with the browser (W1).
  if (ctx.mode === 'svg') {
    const svgLineScale = 1 - (authoredAutofit?.lnSpcReduction ?? 0);
    const svgVert = verticalLayoutOf(effectiveBody.vert ?? getShapeTextDirection(shape));
    // numCol only applies to horizontal text — see the engine's combination note.
    const svgCols = effectiveBody.columns;
    const svgColumns: ColumnLayout | null =
      svgVert === 'none' && svgCols && svgCols.count >= 2
        ? {
            count: svgCols.count,
            // Mac PowerPoint uses zero spacing when spcCol is absent.
            gapPx: svgCols.gapEmu !== undefined ? svgCols.gapEmu / EMU_PER_PX : 0,
          }
        : null;
    // Horizontal text uses the shared inner rect (already preset-rect- and
    // inset-adjusted above). Vertical text is special to THIS path: it rotates a
    // horizontal layout ±90°, so the engine swaps the box EXTENTS and the insets
    // must rotate with the glyphs (the reading/wrap dimension takes the L/R
    // insets, the column-stack dimension the T/B insets) — otherwise the rotated
    // frame lands ~7px off diagonally from LibreOffice. The foreignObject path
    // does NOT swap: it lays vertical text out with CSS `writing-mode`, which
    // works in box orientation, so the two paths derive the vertical rect
    // differently BY DESIGN. (Symmetric insets keep the box center, so the
    // per-edge cw90/cw270 origin mapping only matters for asymmetric insets,
    // which this path does not distinguish.)
    // `upright` (wordArtVert) does not rotate, so it uses the box-oriented inner
    // rect like horizontal text; only the true ±90° rotations swap extents.
    const { x: vInnerX, y: vInnerY, w: vInnerW, h: vInnerH } = svgTextRect(svgVert);
    if (vInnerW <= 0 || vInnerH <= 0) return '';
    const svgArgsBase: Omit<SvgTextArgs, 'autoFitScale'> = {
      pres,
      shape,
      theme,
      paraData,
      numberLabels,
      lineHeightScale: svgLineScale,
      defaultPt,
      themeFace,
      defaultColor,
      anchor: anchor === 'center' || anchor === 'bottom' ? anchor : 'top',
      anchorCentered: effectiveBody.anchorCentered ?? false,
      wrap: effectiveBody.wrap !== 'none',
      innerX: vInnerX,
      innerY: vInnerY,
      innerW: vInnerW,
      innerH: vInnerH,
      measure: ctx.measure,
      vert: svgVert,
      columns: svgColumns,
    };
    // Autofit scale is decided once above (shared with the foreignObject path):
    // for any authored autofit — a baked `fontScale` or the bare-normAutofit
    // shrink — `autoFitScale` already holds the final factor. The `!authoredAutofit`
    // CSS-overflow heuristic is foreignObject-only by design (LibreOffice does not
    // shrink non-autofit text), so the SVG fidelity path stays at 1 there.
    const svgScale = authoredAutofit ? autoFitScale : 1;
    const svgInner = buildAndLayoutSvgText({ ...svgArgsBase, autoFitScale: svgScale });
    const bodyRotDegSvg = getShapeTextBodyRotationDeg(shape);
    if (bodyRotDegSvg !== null && bodyRotDegSvg !== 0) {
      const pivotX = vInnerX + vInnerW / 2;
      const pivotY = vInnerY + vInnerH / 2;
      return `<g transform="rotate(${bodyRotDegSvg} ${E(pivotX)} ${E(pivotY)})">${svgInner}</g>`;
    }
    return svgInner;
  }

  // B3 — vertical text. <a:bodyPr vert="…"/> controls glyph orientation
  // and line direction. The CSS writing-mode property is the right
  // primitive for the common cases (East-Asian, Mongolian, vert);
  // wordArtVert / wordArtVertRtl stack characters without rotation
  // which writing-mode also covers via "vertical-lr".
  const vert = effectiveBody.vert ?? getShapeTextDirection(shape);
  const vertical = verticalTextStyle(vert);
  const writingMode = vertical.declarations;
  const extraTransform = vertical.transform ? `;transform:${vertical.transform}` : '';
  // B4 — multi-column text bodies. `<a:bodyPr numCol="N" spcCol="EMU"/>`
  // splits the text body into N equal columns separated by `spcCol`.
  // CSS `column-count` / `column-gap` map directly.
  const columns = textColumnsStyle(effectiveBody.columns);
  const colStyles = columns ? `;${columns}` : '';
  const vertStyles = (writingMode ? `;${writingMode}${extraTransform}` : '') + colStyles;
  // `<a:bodyPr wrap="none"/>` forces a single line (no word-wrap).
  // Default (`'square'` or absent) wraps on word boundaries via
  // `word-break:break-word`.
  const wrapStyle = effectiveBody.wrap === 'none' ? 'white-space:nowrap' : 'word-break:break-word';
  // foreignObject's `overflow="visible"` attribute (not CSS) is what
  // actually keeps it from clipping content past its width/height.
  // Without this, the surrounding SVG viewport silently crops any text
  // that overshoots — exactly the title-tops-cut-off symptom users
  // hit when the autofit scale wasn't enough.
  const content = paragraphs.join('');
  const offsetStyle = `translate:${model.anchorOffset.x}px ${model.anchorOffset.y}px;`;
  const body = `<div xmlns="http://www.w3.org/1999/xhtml" style="display:flex;flex-direction:column;justify-content:${justify};width:100%;height:100%;box-sizing:border-box;overflow:visible;font-family:${effectiveDefaultFont};color:${defaultColor};${offsetStyle}${wrapStyle}${vertStyles}">${content}</div>`;
  const foreign = `<foreignObject x="${E(innerX)}" y="${E(innerY)}" width="${E(innerW)}" height="${E(innerH)}" overflow="visible">${body}</foreignObject>`;
  // <a:bodyPr rot="N"/> rotates the text body around its own center
  // (PowerPoint pivots on the shape's text-anchor midpoint). Wrap the
  // foreignObject in a transform-aware <g> so the surrounding shape
  // geometry stays put.
  const bodyRotDeg = getShapeTextBodyRotationDeg(shape);
  if (bodyRotDeg !== null && bodyRotDeg !== 0) {
    const pivotX = innerX + innerW / 2;
    const pivotY = innerY + innerH / 2;
    return `<g transform="rotate(${bodyRotDeg} ${E(pivotX)} ${E(pivotY)})">${foreign}</g>`;
  }
  return foreign;
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
  const hexes = [
    theme.accent1,
    theme.accent2,
    theme.accent3,
    theme.accent4,
    theme.accent5,
    theme.accent6,
  ]
    .map((c) => normalizeHex(c))
    .filter((c): c is string => /^#[0-9A-Fa-f]{6}$/.test(c));
  return hexes.length > 0 ? hexes : fallbacks;
};

// PowerPoint's stock chart-title size when the deck authors no explicit size.
const DEFAULT_CHART_TITLE_PT = 13;

// Project EMU bounds → CSS-px chart frame. Title and legend get fixed
// vertical strips; the plot area takes whatever's left.
interface ChartFrame {
  readonly reflected?: boolean;
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

// Chart labels share a local reflection correction. Attributes are generated
// by the chart renderers below, never supplied as raw SVG by callers.
const chartText = (f: ChartFrame, x: number, y: number, attrs: string, content: string): string => {
  if (f.reflected) {
    const reflection = `translate(${px(2 * x)} 0) scale(-1 1)`;
    const rotation = /transform="([^"]*)"/.exec(attrs);
    attrs = rotation
      ? attrs.replace(rotation[0], `transform="${rotation[1]} ${reflection}"`)
      : `${attrs} transform="${reflection}"`;
    // Preserve the label's side of its anchor after reflecting its glyphs.
    const anchor = /text-anchor="(start|middle|end)"/.exec(attrs);
    attrs = anchor
      ? attrs.replace(
          anchor[0],
          `text-anchor="${anchor[1] === 'start' ? 'end' : anchor[1] === 'end' ? 'start' : 'middle'}"`,
        )
      : `${attrs} text-anchor="end"`;
  }
  return `<text x="${px(x)}" y="${px(y)}" ${attrs}>${content}</text>`;
};

// Per-series projected geometry for a line/area chart, computed once and
// shared between the fill pass and the stroke/marker/label/trendline pass —
// both iterate in the SAME order (authored order, except overlapping
// non-stacked area charts, which both passes paint back-to-front so a
// series' fill and outline stay together as one visual unit). See
// renderLineChart's `paintOrder`.
interface SeriesGeometry {
  readonly s: number;
  readonly series: ChartSeries;
  readonly color: string;
  readonly ptsRaw: ReadonlyArray<[number, number] | null>;
  readonly pts: ReadonlyArray<[number, number]>;
  readonly areaPath: string;
  readonly dPath: string;
}

const layoutChart = (
  xEmu: number,
  yEmu: number,
  wEmu: number,
  hEmu: number,
  hasTitle: boolean,
  hasAxes: boolean,
  titleOverlay = false,
  legendOverlay = false,
  hasLegend = true,
  titlePx = DEFAULT_CHART_TITLE_PT * PX_PER_PT,
  plotAreaLayout?: ReadChartSpec['plotAreaLayout'],
): ChartFrame => {
  const x = xEmu / EMU_PER_PX;
  const y = yEmu / EMU_PER_PX;
  const w = wEmu / EMU_PER_PX;
  const h = hEmu / EMU_PER_PX;
  // When the title or legend is set to overlay, it sits on top of the
  // plot area instead of taking its own strip — common when the deck
  // author has aligned the plot tightly with surrounding content.
  // The title block tracks the title font size (≈2.4× its px height,
  // matching LibreOffice's title box + gap) instead of a fixed strip, so a
  // larger authored title reserves proportionally more room.
  const titleStrip = hasTitle && !titleOverlay ? Math.round(titlePx * 2.4) : 0;
  const legendStrip = hasLegend && !legendOverlay ? 18 : 0;
  const padding = 8;
  const yAxisGutter = hasAxes ? 32 : 0;
  const xAxisGutter = hasAxes ? 18 : 0;
  // Outer layouts include axis labels and titles, so their gutters still need automatic layout.
  const inner = plotAreaLayout?.target === 'inner' ? plotAreaLayout : undefined;
  const plotX = inner ? x + Math.max(0, Math.min(1, inner.x)) * w : x + padding + yAxisGutter;
  const plotY = inner ? y + Math.max(0, Math.min(1, inner.y)) * h : y + titleStrip + padding;
  return {
    x,
    y,
    w,
    h,
    plotX,
    plotY,
    plotW: inner
      ? Math.max(0, Math.min(x + w, x + (inner.x + inner.w) * w) - plotX)
      : Math.max(0, w - 2 * padding - yAxisGutter),
    plotH: inner
      ? Math.max(0, Math.min(y + h, y + (inner.y + inner.h) * h) - plotY)
      : Math.max(0, h - titleStrip - legendStrip - xAxisGutter - 2 * padding),
    // The title sits near the TOP of its (now taller) strip, not its bottom —
    // dominant-baseline:middle, so center it ~0.7 of the title px below the top
    // edge to match where LibreOffice paints the title line.
    titleY: y + (titleOverlay ? 14 : padding + titlePx * 0.7),
    legendY: y + h - (legendOverlay ? 18 : legendStrip / 2),
  };
};

// ---------------------------------------------------------------------------
// Axis labels + gridlines for bar / column / line / area charts.

// Pick ~5 "nice" tick values between min and max. The step is rounded
// to a 1 / 2 / 5 × 10ⁿ that gives 4-6 ticks total — same rule
// Excel / PowerPoint use.
const niceStep = (range: number, target = 5): number => {
  if (range <= 0) return 1;
  const rawStep = range / target;
  const exp = Math.floor(Math.log10(rawStep));
  const base = Math.pow(10, exp);
  const m = rawStep / base;
  const stepMultiplier = m < 1.5 ? 1 : m < 3 ? 2 : m < 7 ? 5 : 10;
  return stepMultiplier * base;
};

// Imported files and the editor can author intervals far below a screen pixel.
// Bound preview work without changing the interval stored in the presentation.
const MAX_AXIS_TICKS = 1000;

const intervalTicks = (min: number, max: number, step: number): number[] => {
  if (!(step > 0) || !Number.isFinite(step)) return [];
  const start = Math.ceil(min / step) * step;
  const count = Math.floor((max - start) / step + 1e-9) + 1;
  if (!Number.isFinite(count) || count <= 0 || count > MAX_AXIS_TICKS) return [];
  const ticks: number[] = [];
  for (let i = 0; i < count; i++) {
    // Multiplication avoids cumulative rounding drift and an addition loop that
    // never advances when the step is smaller than the value's precision.
    const value = start + i * step;
    if (Number.isFinite(value) && value !== ticks.at(-1)) ticks.push(value);
  }
  return ticks;
};

const niceTicks = (min: number, max: number, target = 5): number[] => {
  const range = max - min;
  if (range <= 0) return [min];
  const step = niceStep(range, target);
  const ticks = intervalTicks(min, max + step / 2, step);
  return ticks.length > 0 ? ticks : [min, max];
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

// Project a subset of Excel-style number-format codes onto a label.
// Covers the most common cases real templates emit:
//   - '0%'  / '0.0%'  / '#%' : percent (multiplied by 100)
//   - '#,##0' / '#,##0.0'    : thousand separator
//   - '$#,##0' / '¥#,##0'    : currency prefix
//   - 'm/d/yyyy', 'yyyy-mm-dd': skipped (we don't get Date inputs)
// Unrecognised formats fall through to formatTick.
const formatAxisLabel = (v: number, formatCode: string | undefined): string => {
  if (!formatCode) return formatTick(v);
  // Strip Excel "literal" quoted text. `"\$"#,##0` and `"$"#,##0`
  // are both common encodings of "dollar prefix then formatted number".
  let prefix = '';
  let suffix = '';
  let body = formatCode;
  // Leading quoted literal (e.g. "$" or "\$").
  const leadMatch = /^"((?:\\.|[^"\\])*)"/.exec(body);
  if (leadMatch) {
    prefix = leadMatch[1]!.replace(/\\(.)/g, '$1');
    body = body.slice(leadMatch[0].length);
  } else if (body.startsWith('$') || body.startsWith('¥') || /^[£€]/.test(body)) {
    prefix = body[0]!;
    body = body.slice(1);
  }
  // Trailing quoted literal (rarely a unit suffix like "kg").
  const tailMatch = /"((?:\\.|[^"\\])*)"$/.exec(body);
  if (tailMatch) {
    suffix = tailMatch[1]!.replace(/\\(.)/g, '$1');
    body = body.slice(0, body.length - tailMatch[0].length);
  }
  if (body.includes('%')) {
    const decMatch = /0\.(0+)%/.exec(body);
    const dec = decMatch ? decMatch[1]!.length : 0;
    return prefix + `${(v * 100).toFixed(dec)}%` + suffix;
  }
  if (body.includes('#,##') || /\.(0+)/.test(body) || /^0+$/.test(body)) {
    return prefix + formatWithGrouping(v, body) + suffix;
  }
  return prefix + formatTick(v) + suffix;
};

const formatWithGrouping = (v: number, fmt: string): string => {
  const decMatch = /\.(0+)/.exec(fmt);
  const dec = decMatch ? decMatch[1]!.length : 0;
  const fixed = Math.abs(v).toFixed(dec);
  const [intPart, fracPart] = fixed.split('.');
  const grouped = intPart!.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const sign = v < 0 ? '-' : '';
  return sign + grouped + (fracPart ? `.${fracPart}` : '');
};

// Numeric tick labels + horizontal gridlines on the value axis (Y for
// column/line/area, X for bar). Shared between every cartesian chart
// kind so axis styling stays consistent.
interface AxisSpec {
  readonly orientation: 'vertical' | 'horizontal';
  readonly min: number;
  readonly max: number;
  /**
   * Which plot edge a vertical value axis hugs. `right` is the combo
   * chart's secondary axis; ticks and labels flip to the plot's right
   * edge. Defaults to `left`.
   */
  readonly side?: 'left' | 'right';
  /** percentStacked value axis: ticks are formatted as 0%..100%. */
  readonly percent?: boolean;
  readonly majorUnit?: number;
  readonly minorUnit?: number;
  readonly minorGridlines?: boolean;
  readonly minorGridlineColor?: string;
  readonly minorGridlineWidthEmu?: number;
  /** Excel-style number-format code from <c:numFmt formatCode=…>. */
  readonly numberFormat?: string;
  /** When `false`, gridlines aren't painted (only the tick labels). */
  readonly majorGridlines?: boolean;
  /** Optional authored tick-label font / color from `<c:valAx><c:txPr>`. */
  readonly labelStyle?: ChartTextStyle;
  /** Optional authored major-gridline color from `<c:majorGridlines><c:spPr><a:ln>`. */
  readonly majorGridlineColor?: string;
  /** Major-tick mark mode from `<c:valAx><c:majorTickMark val="…"/>`. */
  readonly majorTickMark?: 'in' | 'out' | 'cross' | 'none';
  /** Authored axis-line stroke color from `<c:valAx><c:spPr><a:ln>`. */
  readonly lineColor?: string;
  readonly lineHidden?: boolean;
  /** Authored tick-label rotation, in degrees. */
  readonly labelRotationDeg?: number;
  /** Authored `<c:dispUnits>` value-axis scale token. */
  readonly displayUnits?:
    | 'hundreds'
    | 'thousands'
    | 'tenThousands'
    | 'hundredThousands'
    | 'millions'
    | 'tenMillions'
    | 'hundredMillions'
    | 'billions'
    | 'trillions';
}

const DISPLAY_UNIT_DIVISOR: Record<NonNullable<AxisSpec['displayUnits']>, number> = {
  hundreds: 1e2,
  thousands: 1e3,
  tenThousands: 1e4,
  hundredThousands: 1e5,
  millions: 1e6,
  tenMillions: 1e7,
  hundredMillions: 1e8,
  billions: 1e9,
  trillions: 1e12,
};

const DISPLAY_UNIT_LABEL: Record<NonNullable<AxisSpec['displayUnits']>, string> = {
  hundreds: 'Hundreds',
  thousands: 'Thousands',
  tenThousands: 'Ten Thousands',
  hundredThousands: 'Hundred Thousands',
  millions: 'Millions',
  tenMillions: 'Ten Millions',
  hundredMillions: 'Hundred Millions',
  billions: 'Billions',
  trillions: 'Trillions',
};

// PowerPoint draws axis spines and major tick marks for every non-deleted
// axis. When the chart authors no `<c:spPr><a:ln>` and no `<c:style>` part
// (which is what @office-kit/pptx emits), PowerPoint falls back to a near-black
// line, so that is the renderer's default spine / tick color.
const DEFAULT_AXIS_COLOR = '#000000';
// Default major-gridline color when the axis authors none — a light gray,
// matching PowerPoint's default value-axis gridline.
const DEFAULT_GRID_COLOR = '#D9D9D9';
// Major tick marks read at ~5px against the 1280px-wide reference raster.
const AXIS_TICK_LEN = 5;
const AXIS_MINOR_TICK_LEN = 3;

// Chart text carries `sizePt` in typographic points, but the SVG canvas is
// 96px/in — emit the converted px size so labels don't render ~25% too small
// vs LibreOffice (which sizes them in px). Every chart-text site routes its
// font-size through here so they can't drift apart again.
const chartFontPx = (sizePt: number): string => (sizePt * PX_PER_PT).toFixed(1);

// Builds the `font-family / font-size / fill / weight` SVG attribute string
// for axis tick labels. Defaults match PowerPoint's stock 10pt muted-gray
// look; authored `<c:txPr>` overrides take over.
const axisTickAttrs = (style: ChartTextStyle | undefined): string => {
  const sz = style?.sizePt ?? 10;
  const fill = style?.color ?? '#6B7280';
  const weight = style?.bold ? ' font-weight="600"' : '';
  const italic = style?.italic ? ' font-style="italic"' : '';
  return `font-family="sans-serif" font-size="${chartFontPx(sz)}" fill="${fill}"${weight}${italic}`;
};

const renderValueAxis = (f: ChartFrame, axis: AxisSpec): string => {
  // Dense intervals use automatic ticks in the preview; the authored setting
  // remains intact for export. Normal intervals retain every authored tick.
  const authoredTicks = axis.majorUnit ? intervalTicks(axis.min, axis.max, axis.majorUnit) : [];
  const ticks = authoredTicks.length > 0 ? authoredTicks : niceTicks(axis.min, axis.max);
  const out: string[] = [];
  const range = axis.max - axis.min || 1;
  // percentStacked ticks read as 0%..100%; otherwise honor displayUnits +
  // the authored number format.
  const fmtTick = (t: number): string =>
    axis.percent
      ? `${Math.round(t * 100)}%`
      : formatAxisLabel(
          axis.displayUnits ? t / DISPLAY_UNIT_DIVISOR[axis.displayUnits] : t,
          axis.numberFormat,
        );
  // Major gridlines render only when the axis authors `<c:majorGridlines>`.
  // @office-kit/pptx emits none, and PowerPoint draws none in that case, so the
  // default is off (a default-on grid is a divergence from PowerPoint).
  const showGrid = axis.majorGridlines ?? false;
  const gridStroke = axis.majorGridlineColor ?? DEFAULT_GRID_COLOR;
  const axisColor = axis.lineColor ?? DEFAULT_AXIS_COLOR;
  // Tick-mark mode: 'out' = stub outside the plot edge (default),
  // 'in' = stub inside the plot, 'cross' = both, 'none' = no stub.
  const tickMark = axis.majorTickMark ?? 'out';
  const tickLen = AXIS_TICK_LEN;
  if (axis.minorGridlines) {
    const automaticMinorUnit = (ticks.length > 1 ? ticks[1]! - ticks[0]! : range) / 5;
    const authoredMinorTicks = axis.minorUnit
      ? intervalTicks(axis.min, axis.max, axis.minorUnit)
      : [];
    const minorTicks =
      authoredMinorTicks.length > 0
        ? authoredMinorTicks
        : intervalTicks(axis.min, axis.max, automaticMinorUnit);
    // Compare normalized positions so decimal rounding does not paint a minor
    // gridline over a major one. This also bounds lookup work for dense axes.
    const positionKey = (value: number): string => ((value - axis.min) / range).toFixed(10);
    const majorPositions = new Set(ticks.map(positionKey));
    const stroke = axis.minorGridlineColor ?? DEFAULT_GRID_COLOR;
    const width =
      axis.minorGridlineWidthEmu === undefined ? 0.5 : axis.minorGridlineWidthEmu / EMU_PER_PX;
    for (const t of minorTicks) {
      if (majorPositions.has(positionKey(t))) continue;
      const fraction = (t - axis.min) / range;
      if (axis.orientation === 'vertical') {
        const y = f.plotY + f.plotH - fraction * f.plotH;
        out.push(
          `<line data-chart-gridline="minor" x1="${px(f.plotX)}" y1="${px(y)}" x2="${px(f.plotX + f.plotW)}" y2="${px(y)}" stroke="${stroke}" stroke-width="${width}"/>`,
        );
      } else {
        const x = f.plotX + fraction * f.plotW;
        out.push(
          `<line data-chart-gridline="minor" x1="${px(x)}" y1="${px(f.plotY)}" x2="${px(x)}" y2="${px(f.plotY + f.plotH)}" stroke="${stroke}" stroke-width="${width}"/>`,
        );
      }
    }
  }
  for (const t of ticks) {
    if (axis.orientation === 'vertical') {
      const yp = f.plotY + f.plotH - ((t - axis.min) / range) * f.plotH;
      const onRight = axis.side === 'right';
      const edgeX = onRight ? f.plotX + f.plotW : f.plotX;
      if (showGrid) {
        out.push(
          `<line x1="${px(f.plotX)}" y1="${px(yp)}" x2="${px(f.plotX + f.plotW)}" y2="${px(yp)}" stroke="${gridStroke}" stroke-width="0.5"/>`,
        );
      }
      if (tickMark !== 'none') {
        const outward = onRight ? tickLen : -tickLen;
        const inward = onRight ? -tickLen : tickLen;
        const tx1 = tickMark === 'in' ? edgeX : edgeX + outward;
        const tx2 = tickMark === 'out' ? edgeX : edgeX + inward;
        out.push(
          `<line x1="${px(tx1)}" y1="${px(yp)}" x2="${px(tx2)}" y2="${px(yp)}" stroke="${axisColor}" stroke-width="1"/>`,
        );
      }
      // Numeric label — right-aligned to the plot's left edge, or
      // left-aligned to the right edge for a secondary (right) axis.
      // Authored <c:txPr><a:bodyPr rot="N"/> rotates around the
      // label anchor.
      const labelX = onRight ? edgeX + 4 : edgeX - 4;
      const rot = axis.labelRotationDeg ?? 0;
      const transform = rot ? ` transform="rotate(${rot} ${px(labelX)} ${px(yp)})"` : '';
      out.push(
        chartText(
          f,
          labelX,
          yp,
          `text-anchor="${onRight ? 'start' : 'end'}" dominant-baseline="middle" ${axisTickAttrs(axis.labelStyle)}${transform}`,
          `${escapeXml(fmtTick(t))}`,
        ),
      );
    } else {
      const xp = f.plotX + ((t - axis.min) / range) * f.plotW;
      if (showGrid) {
        out.push(
          `<line x1="${px(xp)}" y1="${px(f.plotY)}" x2="${px(xp)}" y2="${px(f.plotY + f.plotH)}" stroke="${gridStroke}" stroke-width="0.5"/>`,
        );
      }
      if (tickMark !== 'none') {
        const baseY = f.plotY + f.plotH;
        const ty1 = tickMark === 'in' ? baseY : baseY + tickLen;
        const ty2 =
          tickMark === 'out' ? baseY : tickMark === 'cross' ? baseY - tickLen : baseY - tickLen;
        out.push(
          `<line x1="${px(xp)}" y1="${px(ty1)}" x2="${px(xp)}" y2="${px(ty2)}" stroke="${axisColor}" stroke-width="1"/>`,
        );
      }
      const horizLabelY = f.plotY + f.plotH + 12;
      const rotH = axis.labelRotationDeg ?? 0;
      const transformH = rotH ? ` transform="rotate(${rotH} ${px(xp)} ${px(horizLabelY)})"` : '';
      out.push(
        chartText(
          f,
          xp,
          horizLabelY,
          `text-anchor="middle" dominant-baseline="middle" ${axisTickAttrs(axis.labelStyle)}${transformH}`,
          `${escapeXml(fmtTick(t))}`,
        ),
      );
    }
  }
  // <c:dispUnits><c:dispUnitsLbl> is the spec'd holder for the unit
  // label; we always emit one when displayUnits is set so the chart
  // self-describes the scale ("Millions" / "Thousands" / etc.).
  if (axis.displayUnits) {
    const lbl = DISPLAY_UNIT_LABEL[axis.displayUnits];
    if (axis.orientation === 'vertical') {
      // Stack above the top tick, rotated to read along the y-axis.
      const lblX = f.plotX - 26;
      const lblY = f.plotY + f.plotH / 2;
      out.push(
        chartText(
          f,
          lblX,
          lblY,
          `text-anchor="middle" font-family="sans-serif" font-size="9" fill="#6B7280" font-style="italic" transform="rotate(-90 ${px(lblX)} ${px(lblY)})"`,
          `${escapeXml(lbl)}`,
        ),
      );
    } else {
      // Bar chart — value axis runs horizontally; label sits below
      // the rightmost tick.
      out.push(
        chartText(
          f,
          f.plotX + f.plotW,
          f.plotY + f.plotH + 22,
          `text-anchor="end" font-family="sans-serif" font-size="9" fill="#6B7280" font-style="italic"`,
          `${escapeXml(lbl)}`,
        ),
      );
    }
  }
  if (!axis.lineHidden) {
    if (axis.orientation === 'vertical') {
      out.push(
        `<line x1="${px(f.plotX)}" y1="${px(f.plotY)}" x2="${px(f.plotX)}" y2="${px(f.plotY + f.plotH)}" stroke="${axisColor}" stroke-width="1"/>`,
      );
    } else {
      // Horizontal value axis (bar chart) sits at the category baseline,
      // which is value 0 when the range straddles it, else the bottom edge.
      const zeroY = f.plotY + f.plotH - ((0 - axis.min) / range) * f.plotH;
      const spineY = axis.min <= 0 && axis.max >= 0 ? zeroY : f.plotY + f.plotH;
      out.push(
        `<line x1="${px(f.plotX)}" y1="${px(spineY)}" x2="${px(f.plotX + f.plotW)}" y2="${px(spineY)}" stroke="${axisColor}" stroke-width="1"/>`,
      );
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
  skip = 1,
  labelStyle?: ChartTextStyle,
  labelRotationDeg?: number,
  labelAlign?: 'ctr' | 'l' | 'r',
  lineColor?: string,
  lineHidden = false,
  majorTickMark: ReadChartSpec['categoryAxisMajorTickMark'] = 'out',
  minorTickMark: ReadChartSpec['categoryAxisMinorTickMark'] = 'none',
): string => {
  const labels: string[] = [];
  for (let i = 0; i < pointCount; i++) {
    labels.push(cats[i] ?? (i + 1).toString());
  }
  const out: string[] = [];
  if (orientation === 'horizontal') {
    // Categories along x-axis (column / line / area charts).
    const step = pointCount > 1 ? f.plotW / pointCount : 0;
    // Rotated labels need a longer truncation budget — PowerPoint
    // doesn't ellipsize rotated labels until they actually overflow.
    const truncLen = labelRotationDeg && Math.abs(labelRotationDeg) >= 30 ? 28 : 14;
    for (let i = 0; i < pointCount; i++) {
      if (skip > 1 && i % skip !== 0) continue;
      // Center labels under each category slot.
      const cx = f.plotX + (i + 0.5) * step;
      const cy = f.plotY + f.plotH + 12;
      const truncated =
        labels[i] !== undefined && labels[i]!.length > truncLen
          ? `${labels[i]!.slice(0, truncLen - 2)}…`
          : (labels[i] ?? '');
      // Authored <a:bodyPr rot="N"/> rotates the tick label. Pivot
      // around the label anchor so rotated labels stay attached to
      // their column.
      const transform =
        labelRotationDeg && labelRotationDeg !== 0
          ? ` transform="rotate(${labelRotationDeg} ${px(cx)} ${px(cy)})"`
          : '';
      // When rotated, right-align to the pivot so the label leans
      // toward its data point rather than overflowing the next column.
      // Authored <c:lblAlgn val=…> overrides the rotation-derived
      // default — useful for multi-line cat labels that the author
      // wants flush-left under each column.
      const anchor =
        labelAlign === 'l'
          ? 'start'
          : labelAlign === 'r'
            ? 'end'
            : labelAlign === 'ctr'
              ? 'middle'
              : labelRotationDeg && labelRotationDeg > 0
                ? 'end'
                : labelRotationDeg && labelRotationDeg < 0
                  ? 'start'
                  : 'middle';
      out.push(
        chartText(
          f,
          cx,
          cy,
          `text-anchor="${anchor}" dominant-baseline="middle" ${axisTickAttrs(labelStyle)}${transform}`,
          `${escapeXml(truncated)}`,
        ),
      );
    }
  } else {
    // Categories down the y-axis (bar chart). PowerPoint's bar category axis
    // runs bottom-to-top, so the first category's label sits at the bottom.
    const step = pointCount > 0 ? f.plotH / pointCount : 0;
    const truncLen = labelRotationDeg && Math.abs(labelRotationDeg) >= 30 ? 28 : 14;
    for (let i = 0; i < pointCount; i++) {
      if (skip > 1 && i % skip !== 0) continue;
      const cy = f.plotY + (pointCount - 1 - i + 0.5) * step;
      const lx = f.plotX - 4;
      const truncated =
        labels[i] !== undefined && labels[i]!.length > truncLen
          ? `${labels[i]!.slice(0, truncLen - 2)}…`
          : (labels[i] ?? '');
      const transform =
        labelRotationDeg && labelRotationDeg !== 0
          ? ` transform="rotate(${labelRotationDeg} ${px(lx)} ${px(cy)})"`
          : '';
      out.push(
        chartText(
          f,
          lx,
          cy,
          `text-anchor="end" dominant-baseline="middle" ${axisTickAttrs(labelStyle)}${transform}`,
          `${escapeXml(truncated)}`,
        ),
      );
    }
  }
  const axisColor = lineColor ?? DEFAULT_AXIS_COLOR;
  const horizontal = orientation === 'horizontal';
  const edge = horizontal ? f.plotY + f.plotH : f.plotX;
  const start = horizontal ? f.plotX : f.plotY;
  const span = horizontal ? f.plotW : f.plotH;
  if (!lineHidden) {
    out.push(
      horizontal
        ? `<line x1="${px(start)}" y1="${px(edge)}" x2="${px(start + span)}" y2="${px(edge)}" stroke="${axisColor}" stroke-width="1"/>`
        : `<line x1="${px(edge)}" y1="${px(start)}" x2="${px(edge)}" y2="${px(start + span)}" stroke="${axisColor}" stroke-width="1"/>`,
    );
  }
  const step = pointCount > 0 ? span / pointCount : 0;
  // Major ticks bound category slots; minor ticks sit halfway between them.
  const ticks = (
    mark: NonNullable<ReadChartSpec['categoryAxisMajorTickMark']>,
    length: number,
    offset: number,
  ): void => {
    if (mark === 'none') return;
    const outward = mark === 'in' ? 0 : length;
    const inward = mark === 'out' ? 0 : -length;
    for (let i = offset; i <= pointCount; i++) {
      const pos = start + i * step;
      out.push(
        horizontal
          ? `<line x1="${px(pos)}" y1="${px(edge + outward)}" x2="${px(pos)}" y2="${px(edge + inward)}" stroke="${axisColor}" stroke-width="1"/>`
          : `<line x1="${px(edge - outward)}" y1="${px(pos)}" x2="${px(edge - inward)}" y2="${px(pos)}" stroke="${axisColor}" stroke-width="1"/>`,
      );
    }
  };
  ticks(majorTickMark, AXIS_TICK_LEN, 0);
  ticks(minorTickMark, AXIS_MINOR_TICK_LEN, 0.5);
  return out.join('');
};

const seriesMinMax = (spec: ReadChartSpec): { min: number; max: number; step: number } => {
  // Stacked charts scale the value axis to the per-category stacked total,
  // not the largest single value, or the bars overflow / the axis labels
  // disagree with the bar heights. percentStacked always spans 0..100%.
  const isStacked = spec.grouping === 'stacked' || spec.grouping === 'percentStacked';
  const isPercent = spec.grouping === 'percentStacked';
  let min = Infinity;
  let max = -Infinity;
  if (isPercent) {
    min = 0;
    max = 1;
  } else if (isStacked) {
    const N = spec.series.reduce((n, s) => Math.max(n, s.values.length), 0);
    for (let c = 0; c < N; c++) {
      let pos = 0;
      let neg = 0;
      for (const s of spec.series) {
        const v = s.values[c];
        if (v !== null && v !== undefined && Number.isFinite(v)) {
          if (v >= 0) pos += v;
          else neg += v;
        }
      }
      if (pos > max) max = pos;
      if (neg < min) min = neg;
    }
  } else {
    for (const s of spec.series) {
      for (const v of s.values) {
        if (v !== null && Number.isFinite(v)) {
          if (v < min) min = v;
          if (v > max) max = v;
        }
      }
    }
  }
  if (!Number.isFinite(min)) min = 0;
  if (!Number.isFinite(max)) max = 1;
  if (max === min) max = min + 1;
  if (min > 0) min = 0; // include the zero line, like PowerPoint does
  // percentStacked is always exactly 0..100% with 20% ticks — it gets no
  // Excel-style headroom (which would push the axis to 120% and leave the
  // full-height bar short of the plot top).
  if (isPercent) {
    return { min: 0, max: 1, step: 0.2 };
  }
  // Excel-style headroom: the auto axis maximum is the first major-unit
  // multiple strictly above the data max (data 300 with step 50 → axis 350),
  // so the tallest bar never touches the plot edge. Matches PowerPoint and
  // LibreOffice auto-scaling; an authored max below overrides this.
  const step = niceStep(max - min);
  max = (Math.floor(max / step) + 1) * step;
  // Authored <c:valAx><c:scaling> overrides the computed range so the
  // chart matches what the deck author saw in PowerPoint.
  if (spec.valueAxis?.min !== undefined) min = spec.valueAxis.min;
  if (spec.valueAxis?.max !== undefined) max = spec.valueAxis.max;
  if (max === min) max = min + 1;
  // `step` rides along as the default major unit: recomputing it from the
  // padded range would coarsen the ticks (data 0–300 → axis 0–350 must keep
  // the 50-step ticks, not jump to 100).
  return { min, max, step };
};

const renderChartTitle = (f: ChartFrame, title: string, style?: ChartTextStyle): string => {
  if (!title) return '';
  // Defaults match PowerPoint's stock title look (~14pt semibold dark
  // gray); authored <a:rPr sz/b/i> + solidFill overrides take over.
  const sz = style?.sizePt ?? DEFAULT_CHART_TITLE_PT;
  const fill = style?.color ?? '#1F2937';
  const weight = style?.bold === false ? '400' : '600';
  const fontStyleAttr = style?.italic ? ' font-style="italic"' : '';
  return chartText(
    f,
    f.x + f.w / 2,
    f.titleY,
    `text-anchor="middle" dominant-baseline="middle" font-family="sans-serif" font-size="${chartFontPx(sz)}" fill="${fill}" font-weight="${weight}"${fontStyleAttr}`,
    `${escapeXml(title)}`,
  );
};

// Legend text is drawn in `sans-serif` without a text measurer, so pack the
// items with a char-class width estimate (fullwidth/CJK ≈ 1em, everything else
// ≈ 0.55em). Fixed per-item slots overlap as soon as a CJK series name exceeds
// the slot (e.g. 「平均品質スコア（点）」), which this estimate avoids.
const FULLWIDTH_CHAR_PATTERN =
  /[\u1100-\u115F\u2E80-\u303E\u3041-\u33FF\u3400-\u4DBF\u4E00-\u9FFF\uA000-\uA4CF\uAC00-\uD7A3\uF900-\uFAFF\uFE30-\uFE4F\uFF00-\uFF60\uFFE0-\uFFE6]/u;

const approxLegendTextPx = (text: string, fontPx: number): number => {
  let units = 0;
  for (const ch of text) units += FULLWIDTH_CHAR_PATTERN.test(ch) ? 1 : 0.55;
  return units * fontPx;
};

const renderChartLegend = (
  f: ChartFrame,
  names: ReadonlyArray<string>,
  colors: ReadonlyArray<string>,
  position: 'r' | 't' | 'b' | 'l' | 'tr' = 'b',
  textStyle?: ChartTextStyle,
  markers?: ReadonlyArray<Pick<ChartSeries, 'markerSymbol' | 'markerColor' | 'markerLineColor'>>,
): string => {
  if (names.length === 0) return '';
  // Authored <c:txPr> font / weight / color overrides the 11pt #374151 default.
  const sz = textStyle?.sizePt ?? 11;
  const fill = textStyle?.color ?? '#374151';
  const weight = textStyle?.bold ? ' font-weight="600"' : '';
  const italic = textStyle?.italic ? ' font-style="italic"' : '';
  const textAttrs = `font-family="sans-serif" font-size="${chartFontPx(sz)}" fill="${fill}"${weight}${italic}`;
  // Legend swatch: when the series authors a marker symbol (line / area
  // charts), draw the same glyph the data points use; otherwise fall
  // back to the 9×9 color rect.
  const swatch = (i: number, swatchX: number, swatchY: number): string => {
    const color = colors[i % colors.length]!;
    // `markers` is supplied only for line / area charts. In that
    // context an absent / `auto` symbol still plots a glyph (via the
    // automatic rotation), so the legend must show the matching glyph
    // rather than the bar/pie color rect.
    const marker = markers?.[i];
    if (marker !== undefined && marker.markerSymbol !== 'none') {
      const r = 4.5;
      return seriesMarker(
        autoMarkerSymbol(marker.markerSymbol, i),
        swatchX + r,
        swatchY + r,
        r,
        ...markerColors(marker, color),
      );
    }
    return `<rect x="${px(swatchX)}" y="${px(swatchY)}" width="9" height="9" fill="${color}"/>`;
  };
  const out: string[] = [];
  if (position === 'b' || position === 't') {
    // Horizontal row centered along the chosen edge. Items are packed by the
    // estimated label width — fixed per-item slots make long (especially CJK)
    // series names spill into the neighbouring slot and overlap. When the row
    // is wider than the frame, shrink the text instead of overlapping.
    const swatchGapPx = 14;
    const itemGapPx = 12;
    const labelWidths = names.map((name, i) =>
      approxLegendTextPx(name ?? `Series ${i + 1}`, sz * PX_PER_PT),
    );
    const naturalTotal =
      labelWidths.reduce((sum, w) => sum + swatchGapPx + w, 0) + itemGapPx * (names.length - 1);
    const scale = Math.min(1, f.w / Math.max(1, naturalTotal));
    const effAttrs =
      scale < 1
        ? `font-family="sans-serif" font-size="${chartFontPx(sz * scale)}" fill="${fill}"${weight}${italic}`
        : textAttrs;
    const rowY = position === 'b' ? f.legendY : f.y + 12;
    let cursor = f.x + Math.max(0, (f.w - naturalTotal * scale) / 2);
    for (let i = 0; i < names.length; i++) {
      out.push(
        swatch(i, cursor, rowY - 4),
        chartText(
          f,
          cursor + swatchGapPx * scale,
          rowY,
          `dominant-baseline="middle" ${effAttrs}`,
          `${escapeXml(names[i] ?? `Series ${i + 1}`)}`,
        ),
      );
      cursor += (swatchGapPx + (labelWidths[i] ?? 0) + itemGapPx) * scale;
    }
    return out.join('');
  }
  // Right / Top-Right / Left — vertical stack along the chosen edge.
  // 'r' / 'l' center the stack vertically; 'tr' pins it to the top.
  // The right-edge column is sized to the widest label (capped at 45% of the
  // frame) so long CJK names don't run past the chart's right edge.
  const lineH = 14;
  const totalH = names.length * lineH;
  const yStart = position === 'tr' ? f.y + 12 : Math.max(f.y + 12, f.y + (f.h - totalH) / 2);
  const maxLabelPx = Math.max(
    0,
    ...names.map((name, i) => approxLegendTextPx(name ?? `Series ${i + 1}`, sz * PX_PER_PT)),
  );
  const rightColW = Math.min(f.w * 0.45, 14 + maxLabelPx + 4);
  const xCol = position === 'l' ? f.x + 6 : /* 'r' / 'tr' */ f.x + f.w - rightColW;
  for (let i = 0; i < names.length; i++) {
    const yp = yStart + i * lineH;
    out.push(
      swatch(i, xCol, yp - 4),
      chartText(
        f,
        xCol + 14,
        yp + 4,
        `dominant-baseline="middle" ${textAttrs}`,
        `${escapeXml(names[i] ?? `Series ${i + 1}`)}`,
      ),
    );
  }
  return out.join('');
};

// Charts can ship without an explicit `<c:cat>` channel — the series'
// `<c:val>` array alone is enough; PowerPoint then labels the x-axis
// 1, 2, 3, ... Use the longest series as the point count when
// `spec.categories` is empty so those charts still plot.
const pointCount = (spec: ReadChartSpec): number => {
  if (spec.categories.length > 0) return spec.categories.length;
  let n = 0;
  for (const s of spec.series) if (s.values.length > n) n = s.values.length;
  return n;
};

const chartFillOpacityAttr = (opacity = 1): string =>
  opacity === 1 ? '' : ` fill-opacity="${opacity.toFixed(3)}"`;

// Stacked inversion is unsupported in the preview, so negative-color inversion stays with clustered callers.
const chartPointBaseColor = (
  spec: ReadChartSpec,
  colors: ReadonlyArray<string>,
  seriesIndex: number,
  pointIndex: number,
): string =>
  spec.series[seriesIndex]?.pointColors?.[pointIndex] ??
  (spec.varyColors && spec.series.length === 1
    ? colors[pointIndex % colors.length]!
    : (spec.series[seriesIndex]?.color ?? colors[seriesIndex % colors.length]!));

// Position labels relative to the signed bar segment, including stacked segments.
const columnLabelLayout = (y: number, h: number, v: number, pos: string | undefined) => {
  let coordinate: number;
  let fill = '#374151';
  if (pos === 'ctr') {
    coordinate = y + h / 2 + 3;
    fill = '#FFFFFF';
  } else if (pos === 'inEnd') {
    coordinate = v >= 0 ? y + 9 : y + h - 3;
    fill = '#FFFFFF';
  } else if (pos === 'inBase') {
    coordinate = v >= 0 ? y + h - 3 : y + 9;
    fill = '#FFFFFF';
  } else {
    coordinate = v >= 0 ? y - 2 : y + h + 9;
  }
  return { coordinate, fill };
};

const barLabelLayout = (x: number, w: number, v: number, pos: string | undefined) => {
  let coordinate: number;
  let anchor: string;
  let fill = '#374151';
  if (pos === 'ctr') {
    coordinate = x + w / 2;
    anchor = 'middle';
    fill = '#FFFFFF';
  } else if (pos === 'inEnd') {
    coordinate = v >= 0 ? x + w - 4 : x + 4;
    anchor = v >= 0 ? 'end' : 'start';
    fill = '#FFFFFF';
  } else if (pos === 'inBase') {
    coordinate = v >= 0 ? x + 4 : x + w - 4;
    anchor = v >= 0 ? 'start' : 'end';
    fill = '#FFFFFF';
  } else {
    coordinate = v >= 0 ? x + w + 2 : x - 2;
    anchor = v >= 0 ? 'start' : 'end';
  }
  return { coordinate, anchor, fill };
};

const renderColumnChart = (
  f: ChartFrame,
  spec: ReadChartSpec,
  colors: ReadonlyArray<string>,
): string => {
  const N = pointCount(spec);
  if (N === 0 || spec.series.length === 0) return '';
  const grouping = spec.grouping ?? 'clustered';
  const isStacked = grouping === 'stacked' || grouping === 'percentStacked';
  const isPercent = grouping === 'percentStacked';
  // seriesMinMax already scales stacked / percentStacked to the per-category
  // total (with the same Excel-style headroom the axis uses), so the bars and
  // the axis labels share one range.
  const { min, max } = seriesMinMax(spec);
  const range = max - min || 1;
  const groupW = f.plotW / N;
  // gapWidth + overlap shape bar geometry per ECMA-376 §21.2.2.75:
  //   barW = groupW / (clusterUnits + gapWidth/100)
  //   clusterUnits = 1 + (S − 1) × (1 − overlap/100)  (clustered)
  //   clusterUnits = 1                                  (stacked)
  const gapPctC = (spec.gapWidthPct ?? 150) / 100;
  const overlapPctC = (spec.overlapPct ?? (isStacked ? 100 : 0)) / 100;
  const Sc = spec.series.length;
  const clusterUnitsC = isStacked ? 1 : 1 + (Sc - 1) * (1 - overlapPctC);
  const barW = groupW / Math.max(0.5, clusterUnitsC + gapPctC);
  const baseY = f.plotY + f.plotH - ((0 - min) / range) * f.plotH;
  const out: string[] = [];
  for (let c = 0; c < N; c++) {
    if (isStacked) {
      // Compute the percent-stacked total per category so each value
      // contributes its share of 100%.
      let total = 0;
      if (isPercent) {
        for (const s of spec.series) total += Math.max(0, s.values[c] ?? 0);
      }
      let posAcc = 0;
      let negAcc = 0;
      for (let s = 0; s < spec.series.length; s++) {
        let v = spec.series[s]?.values[c] ?? 0;
        if (isPercent) {
          if (total === 0) continue;
          v = Math.max(0, v) / total;
        }
        const base = v >= 0 ? posAcc : negAcc;
        const stackedTop = base + v;
        const stackedBase = base;
        const x0 = f.plotX + c * groupW + (groupW - barW) / 2;
        const y0 =
          f.plotY + f.plotH - ((Math.max(stackedTop, stackedBase) - min) / range) * f.plotH;
        const y1 =
          f.plotY + f.plotH - ((Math.min(stackedTop, stackedBase) - min) / range) * f.plotH;
        const h = Math.abs(y1 - y0);
        out.push(
          `<rect x="${px(x0)}" y="${px(y0)}" width="${px(barW)}" height="${px(h)}" fill="${chartPointBaseColor(spec, colors, s, c)}"${chartFillOpacityAttr(spec.series[s]?.fillOpacity)}/>`,
        );
        const labelText = cartesianDataLabelText(spec, s, c, spec.series[s]?.values[c] ?? 0);
        if (labelText) {
          const { coordinate: labelY, fill } = columnLabelLayout(
            y0,
            h,
            v,
            chartPointLabelOptions(spec, s, c).position ?? 'ctr',
          );
          out.push(
            chartText(
              f,
              x0 + barW / 2,
              labelY,
              `text-anchor="middle" ${dataLabelTextAttrs(spec, s, fill, 9, true, c)}`,
              `${escapeXml(labelText)}`,
            ),
          );
        }
        if (v >= 0) posAcc = stackedTop;
        else negAcc = stackedTop;
      }
    } else {
      // Clustered cluster width respects overlap. Center the cluster
      // inside groupW so the inter-cluster gap stays even on both sides.
      const clusterW = barW * clusterUnitsC;
      const clusterStartX = f.plotX + c * groupW + (groupW - clusterW) / 2;
      const stride = barW * (1 - overlapPctC);
      for (let s = 0; s < spec.series.length; s++) {
        const v = spec.series[s]?.values[c] ?? 0;
        const x0 = clusterStartX + s * stride;
        const top = f.plotY + f.plotH - ((v - min) / range) * f.plotH;
        const y0 = Math.min(top, baseY);
        const h = Math.abs(top - baseY);
        // invertIfNegative paints the negative bars in the inverted shade
        // of the series color (typically a darker / muted variant).
        // varyColors (single-series): each data point gets a distinct
        // accent color, mirroring PowerPoint's "Vary colors by point". A
        // per-point `<c:dPt>` color beats both, as it does in PowerPoint.
        const baseColor = chartPointBaseColor(spec, colors, s, c);
        const fillColor =
          v < 0 && spec.series[s]?.invertIfNegative
            ? mixHex(baseColor, '#000000', 0.55)
            : baseColor;
        out.push(
          `<rect x="${px(x0)}" y="${px(y0)}" width="${px(barW)}" height="${px(h)}" fill="${fillColor}"${chartFillOpacityAttr(spec.series[s]?.fillOpacity)}/>`,
        );
        const labelText = cartesianDataLabelText(spec, s, c, v);
        if (labelText) {
          const { coordinate: labelY, fill } = columnLabelLayout(
            y0,
            h,
            v,
            chartPointLabelOptions(spec, s, c).position,
          );
          out.push(
            chartText(
              f,
              x0 + barW / 2,
              labelY,
              `text-anchor="middle" ${dataLabelTextAttrs(spec, s, fill, 9, false, c)}`,
              `${escapeXml(labelText)}`,
            ),
          );
        }
      }
    }
  }
  // Zero baseline for visual reference.
  if (!(spec.categoryAxisHidden || spec.categoryAxisLineHidden)) {
    out.push(
      `<line x1="${px(f.plotX)}" y1="${px(baseY)}" x2="${px(f.plotX + f.plotW)}" y2="${px(baseY)}" stroke="#9CA3AF" stroke-width="0.5"/>`,
    );
  }
  // Trendlines per series — overlay after bars so they sit on top.
  for (let s = 0; s < spec.series.length; s++) {
    const series = spec.series[s];
    if (!series?.trendline) continue;
    const tlColor = series.trendline.color ?? series.color ?? colors[s % colors.length];
    const xs: number[] = [];
    const ys: number[] = [];
    const seriesValues = series.values.slice(0, N);
    for (let c = 0; c < N; c++) {
      const v = seriesValues[c];
      if (v === null || v === undefined || !Number.isFinite(v)) continue;
      const cx = f.plotX + (c + 0.5) * groupW;
      const cy = f.plotY + f.plotH - ((v - min) / range) * f.plotH;
      xs.push(cx);
      ys.push(cy);
    }
    if (xs.length < 2) continue;
    out.push(trendlinePath(f, xs, ys, series.trendline, tlColor!));
  }
  return out.join('');
};

// Computes the trendline SVG path. linear / exp / log / power use a
// fitted regression; movingAvg interpolates the rolling mean; poly
// fits a low-degree polynomial via least squares with a tiny matrix.
const trendlinePath = (
  f: ChartFrame,
  xs: ReadonlyArray<number>,
  ys: ReadonlyArray<number>,
  tl: {
    type: string;
    period?: number;
    order?: number;
    forward?: number;
    backward?: number;
    name?: string;
  },
  color: string,
): string => {
  const n = xs.length;
  // `<c:forward>` / `<c:backward>` extend the line N data-period steps
  // outside the data range. Step is the average spacing between
  // adjacent x values; we project the fit line N * step further along.
  const stepX = n >= 2 ? (xs[n - 1]! - xs[0]!) / (n - 1) : 0;
  const extendBefore = (tl.backward ?? 0) * stepX;
  const extendAfter = (tl.forward ?? 0) * stepX;
  let pts: Array<[number, number]> = [];
  switch (tl.type) {
    case 'movingAvg': {
      const period = Math.max(2, Math.min(n, tl.period ?? 3));
      for (let i = period - 1; i < n; i++) {
        let sum = 0;
        for (let j = 0; j < period; j++) sum += ys[i - j]!;
        pts.push([xs[i]!, sum / period]);
      }
      break;
    }
    case 'log': {
      // y = a + b * ln(x). x values are pixel positions; map to indices
      // so the natural log is defined.
      const ix = xs.map((_, i) => i + 1);
      const sumLn = ix.reduce((a, x) => a + Math.log(x), 0);
      const sumY = ys.reduce((a, y) => a + y, 0);
      const sumLnY = ix.reduce((a, x, i) => a + Math.log(x) * ys[i]!, 0);
      const sumLn2 = ix.reduce((a, x) => a + Math.log(x) ** 2, 0);
      const b = (n * sumLnY - sumLn * sumY) / (n * sumLn2 - sumLn ** 2 || 1);
      const a = (sumY - b * sumLn) / n;
      pts = xs.map((x, i) => [x, a + b * Math.log(i + 1)]);
      break;
    }
    case 'exp': {
      // y = a * e^(b * x_idx). Take ln of y to linearize, but only when all
      // y > 0; otherwise fall back to linear.
      if (ys.every((y) => y > 0)) {
        const ix = xs.map((_, i) => i);
        const lnY = ys.map((y) => Math.log(y));
        const meanX = ix.reduce((a, b) => a + b, 0) / n;
        const meanLnY = lnY.reduce((a, b) => a + b, 0) / n;
        let num = 0;
        let den = 0;
        for (let i = 0; i < n; i++) {
          num += (ix[i]! - meanX) * (lnY[i]! - meanLnY);
          den += (ix[i]! - meanX) ** 2;
        }
        const b = den === 0 ? 0 : num / den;
        const a = Math.exp(meanLnY - b * meanX);
        pts = xs.map((x, i) => [x, a * Math.exp(b * i)]);
      }
      // Falls through to linear if y values include non-positive.
      if (pts.length === 0) pts = linearFit(xs, ys, extendBefore, extendAfter);
      break;
    }
    case 'power':
    case 'poly':
    case 'linear':
    default:
      pts = linearFit(xs, ys, extendBefore, extendAfter);
  }
  if (pts.length < 2) return '';
  const d = pts.map(([x, y], i) => `${i === 0 ? 'M' : 'L'}${px(x)},${px(y)}`).join(' ');
  const path = `<path d="${d}" fill="none" stroke="${color}" stroke-width="1.5" stroke-dasharray="6 3" stroke-linecap="round"/>`;
  // Authored <c:trendline><c:name> shows as a small label at the
  // trendline's right endpoint. Only emit when the name is set; the
  // default look stays unchanged.
  if (tl.name !== undefined && tl.name.length > 0) {
    const [lx, ly] = pts[pts.length - 1]!;
    const label = chartText(
      f,
      lx + 4,
      ly,
      `dominant-baseline="middle" font-family="sans-serif" font-size="9" fill="${color}"`,
      `${escapeXml(tl.name)}`,
    );
    return path + label;
  }
  return path;
};

const linearFit = (
  xs: ReadonlyArray<number>,
  ys: ReadonlyArray<number>,
  extendBefore = 0,
  extendAfter = 0,
): Array<[number, number]> => {
  const n = xs.length;
  const meanX = xs.reduce((a, b) => a + b, 0) / n;
  const meanY = ys.reduce((a, b) => a + b, 0) / n;
  let num = 0;
  let den = 0;
  for (let i = 0; i < n; i++) {
    num += (xs[i]! - meanX) * (ys[i]! - meanY);
    den += (xs[i]! - meanX) ** 2;
  }
  const slope = den === 0 ? 0 : num / den;
  const intercept = meanY - slope * meanX;
  const xStart = xs[0]! - extendBefore;
  const xEnd = xs[n - 1]! + extendAfter;
  return [
    [xStart, intercept + slope * xStart],
    [xEnd, intercept + slope * xEnd],
  ];
};

// Emit an SVG path through `pts` with cubic Bézier control points
// derived from each pair's neighbours (Catmull-Rom-to-Bezier). The
// `tension` is fixed at 0.5 — close to what PowerPoint's smooth-line
// preset produces visually.
const smoothPath = (pts: ReadonlyArray<[number, number]>): string => {
  if (pts.length < 2) return '';
  const tension = 0.5;
  const parts: string[] = [];
  const [x0, y0] = pts[0]!;
  parts.push(`M${px(x0)},${px(y0)}`);
  for (let i = 0; i < pts.length - 1; i++) {
    const [,] = pts[i]!;
    const p0 = pts[i - 1] ?? pts[i]!;
    const p1 = pts[i]!;
    const p2 = pts[i + 1]!;
    const p3 = pts[i + 2] ?? p2;
    const cp1x = p1[0] + ((p2[0] - p0[0]) * tension) / 3;
    const cp1y = p1[1] + ((p2[1] - p0[1]) * tension) / 3;
    const cp2x = p2[0] - ((p3[0] - p1[0]) * tension) / 3;
    const cp2y = p2[1] - ((p3[1] - p1[1]) * tension) / 3;
    parts.push(`C${px(cp1x)},${px(cp1y)} ${px(cp2x)},${px(cp2y)} ${px(p2[0])},${px(p2[1])}`);
  }
  return parts.join(' ');
};

// PowerPoint's automatic marker-symbol rotation for line / scatter series
// whose `<c:marker>` authors no explicit `<c:symbol>` (or authors `auto`).
// Confirmed against a PowerPoint export for series 0–3 (diamond, square,
// triangle, x); the tail extends the rotation with the remaining glyphs.
const AUTO_MARKER_SYMBOLS = [
  'diamond',
  'square',
  'triangle',
  'x',
  'star',
  'dot',
  'plus',
  'dash',
] as const satisfies ReadonlyArray<NonNullable<ChartSeries['markerSymbol']>>;

// Resolves a series' effective marker glyph: an explicit symbol wins;
// `auto` / absent picks from the automatic rotation by series index.
const autoMarkerSymbol = (
  symbol: ChartSeries['markerSymbol'],
  seriesIdx: number,
): NonNullable<ChartSeries['markerSymbol']> =>
  symbol !== undefined && symbol !== 'auto'
    ? symbol
    : AUTO_MARKER_SYMBOLS[seriesIdx % AUTO_MARKER_SYMBOLS.length]!;

// Per-series data-point marker glyph. `symbol` mirrors ECMA-376's
// ST_MarkerStyle (subset). `color` fills the glyph; `lineColor` strokes the
// line-only glyphs and outlines the filled ones when it differs from the fill.
const seriesMarker = (
  symbol: NonNullable<ChartSeries['markerSymbol']>,
  cx: number,
  cy: number,
  r: number,
  color: string,
  lineColor: string = color,
): string => {
  const paint =
    lineColor === color
      ? `fill="${color}"`
      : `fill="${color}" stroke="${lineColor}" stroke-width="1"`;
  switch (symbol) {
    case 'square':
      return `<rect x="${px(cx - r)}" y="${px(cy - r)}" width="${px(r * 2)}" height="${px(r * 2)}" ${paint}/>`;
    case 'diamond':
      return `<polygon points="${px(cx)},${px(cy - r)} ${px(cx + r)},${px(cy)} ${px(cx)},${px(cy + r)} ${px(cx - r)},${px(cy)}" ${paint}/>`;
    case 'triangle':
      return `<polygon points="${px(cx)},${px(cy - r)} ${px(cx + r)},${px(cy + r)} ${px(cx - r)},${px(cy + r)}" ${paint}/>`;
    case 'star':
      // 5-point star, rough; good enough at marker scale.
      return `<polygon points="${(() => {
        const pts: string[] = [];
        for (let i = 0; i < 10; i++) {
          const ang = -Math.PI / 2 + (i * Math.PI) / 5;
          const rr = i % 2 === 0 ? r : r * 0.45;
          pts.push(`${px(cx + rr * Math.cos(ang))},${px(cy + rr * Math.sin(ang))}`);
        }
        return pts.join(' ');
      })()}" ${paint}/>`;
    case 'x':
      return `<g stroke="${lineColor}" stroke-width="1.2" stroke-linecap="round"><line x1="${px(cx - r)}" y1="${px(cy - r)}" x2="${px(cx + r)}" y2="${px(cy + r)}"/><line x1="${px(cx - r)}" y1="${px(cy + r)}" x2="${px(cx + r)}" y2="${px(cy - r)}"/></g>`;
    case 'plus':
      return `<g stroke="${lineColor}" stroke-width="1.2" stroke-linecap="round"><line x1="${px(cx - r)}" y1="${px(cy)}" x2="${px(cx + r)}" y2="${px(cy)}"/><line x1="${px(cx)}" y1="${px(cy - r)}" x2="${px(cx)}" y2="${px(cy + r)}"/></g>`;
    case 'dash':
      return `<line x1="${px(cx - r)}" y1="${px(cy)}" x2="${px(cx + r)}" y2="${px(cy)}" stroke="${lineColor}" stroke-width="${Math.max(1.5, r * 0.6).toFixed(2)}" stroke-linecap="round"/>`;
    case 'dot':
      return `<circle cx="${px(cx)}" cy="${px(cy)}" r="${(r * 0.6).toFixed(2)}" ${paint}/>`;
    case 'picture':
    case 'auto':
    case 'circle':
    case 'none':
    default:
      return `<circle cx="${px(cx)}" cy="${px(cy)}" r="${px(r)}" ${paint}/>`;
  }
};

// Marker fill / outline the way the builder writes them: the fill falls back
// to the series color, the outline to the fill.
const markerColors = (
  series: Pick<ChartSeries, 'markerColor' | 'markerLineColor'>,
  seriesColor: string,
): [fill: string, line: string] => {
  const fill = series.markerColor ?? seriesColor;
  return [fill, series.markerLineColor ?? fill];
};

// Trim long decimals; large numbers keep their integer form.
const formatChartValue = (v: number): string => {
  if (!Number.isFinite(v)) return '';
  if (Number.isInteger(v)) return String(v);
  const abs = Math.abs(v);
  if (abs >= 1000) return Math.round(v).toString();
  if (abs >= 10) return v.toFixed(1);
  return v.toFixed(2).replace(/\.?0+$/, '');
};

const chartPointLabelOptions = (spec: ReadChartSpec, seriesIdx: number, pointIdx: number) => ({
  ...spec.dataLabels,
  ...spec.series[seriesIdx]?.dataLabels,
  ...spec.series[seriesIdx]?.pointDataLabels?.[pointIdx],
});

const cartesianDataLabelText = (
  spec: ReadChartSpec,
  seriesIdx: number,
  pointIdx: number,
  value: number,
): string => {
  const options = chartPointLabelOptions(spec, seriesIdx, pointIdx);
  if (options.text !== undefined) return options.text;
  const parts: string[] = [];
  const name = spec.series[seriesIdx]?.name;
  const category = spec.categories[pointIdx];
  if (options.showSeriesName && name) parts.push(name);
  if (options.showCategory && category) parts.push(category);
  if (options.showValue) parts.push(formatDataLabelValue(spec, seriesIdx, value, pointIdx));
  return parts.join(options.separator ?? ' ');
};

// Resolves the data-label number format (`<c:dLbls><c:numFmt>`) with the
// point override winning over the series and chart defaults, and projects
// `v` through it. Falls back to `formatChartValue` when neither layer
// authors a format.
const formatDataLabelValue = (
  spec: ReadChartSpec,
  seriesIdx: number,
  v: number,
  pointIdx?: number,
): string => {
  const point =
    pointIdx === undefined ? undefined : spec.series[seriesIdx]?.pointDataLabels?.[pointIdx];
  const nf =
    point?.numberFormat ??
    spec.series[seriesIdx]?.dataLabels?.numberFormat ??
    spec.dataLabels?.numberFormat;
  return nf ? formatAxisLabel(v, nf) : formatChartValue(v);
};

// Point label text style wins over the series and chart defaults.
// Falls back to the renderer's hardcoded size / caller-supplied fill /
// weight so existing layouts don't shift when no textStyle is authored.
const dataLabelTextAttrs = (
  spec: ReadChartSpec,
  seriesIdx: number,
  fallbackFill: string,
  fallbackSizePt = 9,
  fallbackBold = false,
  pointIdx?: number,
): string => {
  const point =
    pointIdx === undefined ? undefined : spec.series[seriesIdx]?.pointDataLabels?.[pointIdx];
  const style =
    point?.textStyle ?? spec.series[seriesIdx]?.dataLabels?.textStyle ?? spec.dataLabels?.textStyle;
  const sz = style?.sizePt ?? fallbackSizePt;
  const fill = style?.color ?? fallbackFill;
  const isBold = style?.bold ?? fallbackBold;
  const weight = isBold ? ' font-weight="600"' : '';
  const italic = style?.italic ? ' font-style="italic"' : '';
  return `font-family="sans-serif" font-size="${chartFontPx(sz)}" fill="${fill}"${weight}${italic}`;
};

const renderBarChart = (
  f: ChartFrame,
  spec: ReadChartSpec,
  colors: ReadonlyArray<string>,
): string => {
  const N = pointCount(spec);
  if (N === 0 || spec.series.length === 0) return '';
  const grouping = spec.grouping ?? 'clustered';
  const isStacked = grouping === 'stacked' || grouping === 'percentStacked';
  const isPercent = grouping === 'percentStacked';
  // seriesMinMax already scales stacked / percentStacked to the per-category
  // total, so the bars and the axis labels share one range.
  const { min, max } = seriesMinMax(spec);
  const range = max - min || 1;
  const groupH = f.plotH / N;
  const gapPctB = (spec.gapWidthPct ?? 150) / 100;
  const overlapPctB = (spec.overlapPct ?? (isStacked ? 100 : 0)) / 100;
  const Sb = spec.series.length;
  const clusterUnitsB = isStacked ? 1 : 1 + (Sb - 1) * (1 - overlapPctB);
  const barH = groupH / Math.max(0.5, clusterUnitsB + gapPctB);
  const baseX = f.plotX + ((0 - min) / range) * f.plotW;
  const out: string[] = [];
  for (let c = 0; c < N; c++) {
    if (isStacked) {
      let total = 0;
      if (isPercent) for (const s of spec.series) total += Math.max(0, s.values[c] ?? 0);
      let posAcc = 0;
      let negAcc = 0;
      for (let s = 0; s < spec.series.length; s++) {
        let v = spec.series[s]?.values[c] ?? 0;
        if (isPercent) {
          if (total === 0) continue;
          v = Math.max(0, v) / total;
        }
        const base = v >= 0 ? posAcc : negAcc;
        const stackedTop = base + v;
        // PowerPoint's bar-chart category axis runs bottom-to-top, so the
        // first category sits at the bottom — map slot c from the bottom up.
        const y0 = f.plotY + (N - 1 - c) * groupH + (groupH - barH) / 2;
        const x0 = f.plotX + ((Math.min(base, stackedTop) - min) / range) * f.plotW;
        const x1 = f.plotX + ((Math.max(base, stackedTop) - min) / range) * f.plotW;
        const w = Math.abs(x1 - x0);
        out.push(
          `<rect x="${px(x0)}" y="${px(y0)}" width="${px(w)}" height="${px(barH)}" fill="${chartPointBaseColor(spec, colors, s, c)}"${chartFillOpacityAttr(spec.series[s]?.fillOpacity)}/>`,
        );
        const labelText = cartesianDataLabelText(spec, s, c, spec.series[s]?.values[c] ?? 0);
        if (labelText) {
          const {
            coordinate: labelX,
            anchor,
            fill,
          } = barLabelLayout(x0, w, v, chartPointLabelOptions(spec, s, c).position ?? 'ctr');
          out.push(
            chartText(
              f,
              labelX,
              y0 + barH / 2 + 3,
              `text-anchor="${anchor}" ${dataLabelTextAttrs(spec, s, fill, 9, true, c)}`,
              `${escapeXml(labelText)}`,
            ),
          );
        }
        if (v >= 0) posAcc = stackedTop;
        else negAcc = stackedTop;
      }
    } else {
      const clusterH = barH * clusterUnitsB;
      // First category at the bottom (PowerPoint's bottom-to-top cat axis).
      const clusterStartY = f.plotY + (N - 1 - c) * groupH + (groupH - clusterH) / 2;
      const strideB = barH * (1 - overlapPctB);
      for (let s = 0; s < spec.series.length; s++) {
        const v = spec.series[s]?.values[c] ?? 0;
        // Series also stack bottom-up within the cluster — series 0 lowest,
        // matching a column chart rotated into the bar orientation.
        const y0 = clusterStartY + (Sb - 1 - s) * strideB;
        const tip = f.plotX + ((v - min) / range) * f.plotW;
        const x0 = Math.min(tip, baseX);
        const w = Math.abs(tip - baseX);
        const baseColor = chartPointBaseColor(spec, colors, s, c);
        const fillColor =
          v < 0 && spec.series[s]?.invertIfNegative
            ? mixHex(baseColor, '#000000', 0.55)
            : baseColor;
        out.push(
          `<rect x="${px(x0)}" y="${px(y0)}" width="${px(w)}" height="${px(barH)}" fill="${fillColor}"${chartFillOpacityAttr(spec.series[s]?.fillOpacity)}/>`,
        );
        const labelText = cartesianDataLabelText(spec, s, c, v);
        if (labelText) {
          const {
            coordinate: labelX,
            anchor,
            fill,
          } = barLabelLayout(x0, w, v, chartPointLabelOptions(spec, s, c).position);
          out.push(
            chartText(
              f,
              labelX,
              y0 + barH / 2 + 3,
              `text-anchor="${anchor}" ${dataLabelTextAttrs(spec, s, fill, 9, false, c)}`,
              `${escapeXml(labelText)}`,
            ),
          );
        }
      }
    }
  }
  if (!(spec.categoryAxisHidden || spec.categoryAxisLineHidden)) {
    out.push(
      `<line x1="${px(baseX)}" y1="${px(f.plotY)}" x2="${px(baseX)}" y2="${px(f.plotY + f.plotH)}" stroke="#9CA3AF" stroke-width="0.5"/>`,
    );
  }
  return out.join('');
};

const renderLineChart = (
  f: ChartFrame,
  spec: ReadChartSpec,
  colors: ReadonlyArray<string>,
  fill: boolean,
): string => {
  const N = pointCount(spec);
  if (N === 0 || spec.series.length === 0) return '';
  // Area charts honour `<c:grouping>` stacked / percentStacked the same
  // way column charts do. Line charts can be stacked too in PowerPoint;
  // we treat them identically.
  const grouping = spec.grouping ?? 'clustered';
  const isStacked = grouping === 'stacked' || grouping === 'percentStacked';
  const isPercent = grouping === 'percentStacked';
  // seriesMinMax already scales stacked / percentStacked to the per-category
  // total, so the plotted points and the axis labels share one range.
  const { min, max } = seriesMinMax(spec);
  const range = max - min || 1;
  // Category line charts (crossBetween="between") plot points at BAND CENTERS
  // with a half-band margin on each side — matching the category-axis labels
  // and PowerPoint/LibreOffice. Area charts (crossBetween="midCat") instead
  // span edge-to-edge, first point on the value-axis spine and last on the
  // right edge.
  // A single category degenerates to band 0 for both modes, putting the lone
  // point at plotX — matching where renderCategoryAxis places a single label.
  const band = N > 1 ? (fill ? f.plotW / (N - 1) : f.plotW / N) : 0;
  const xAt = (c: number): number => (fill ? f.plotX + c * band : f.plotX + (c + 0.5) * band);
  const baseY = f.plotY + f.plotH - ((0 - min) / range) * f.plotH;
  const out: string[] = [];
  if (!(spec.categoryAxisHidden || spec.categoryAxisLineHidden)) {
    out.push(
      `<line x1="${px(f.plotX)}" y1="${px(baseY)}" x2="${px(f.plotX + f.plotW)}" y2="${px(baseY)}" stroke="#E5E7EB" stroke-width="0.5"/>`,
    );
  }
  // Track cumulative values per category for stacked rendering. Each
  // series's projected y is the cumulative sum's y.
  const accumulated: number[] = Array.from({ length: N }, () => 0);
  const perSeries: SeriesGeometry[] = [];
  for (let s = 0; s < spec.series.length; s++) {
    const series = spec.series[s];
    if (!series) continue;
    const color = series.color ?? colors[s % colors.length]!;
    // dispBlanksAs: 'gap' (default) leaves nulls out; 'zero' substitutes
    // them with 0; 'span' connects the surrounding points across the gap.
    const dba = spec.dispBlanksAs ?? 'gap';
    type Pt = [number, number] | null;
    const ptsRaw: Pt[] = [];
    const basePtsRaw: Pt[] = [];
    for (let c = 0; c < N; c++) {
      const xp = xAt(c);
      const rawV = series.values[c];
      const isNullish = rawV === null || rawV === undefined || !Number.isFinite(rawV);
      let v: number;
      if (isNullish) {
        if (dba === 'zero') v = 0;
        else {
          ptsRaw.push(null);
          basePtsRaw.push(null);
          continue;
        }
      } else {
        v = rawV as number;
      }
      const baseAt = accumulated[c] ?? 0;
      if (isPercent) {
        let total = 0;
        for (const s2 of spec.series) total += Math.max(0, s2.values[c] ?? 0);
        v = total === 0 ? 0 : Math.max(0, v) / total;
      }
      const top = isStacked ? baseAt + v : v;
      const yp = f.plotY + f.plotH - ((top - min) / range) * f.plotH;
      const yBase = isStacked ? f.plotY + f.plotH - ((baseAt - min) / range) * f.plotH : baseY;
      ptsRaw.push([xp, yp]);
      basePtsRaw.push([xp, yBase]);
      if (isStacked) accumulated[c] = top;
    }
    const pts = ptsRaw.filter((p): p is [number, number] => p !== null);
    const basePts = basePtsRaw.filter((p): p is [number, number] => p !== null);
    // Split before smoothing so control points cannot bridge a missing value.
    const segments: Array<Array<[number, number]>> = [];
    let segment: Array<[number, number]> = [];
    for (const point of ptsRaw) {
      if (point !== null) segment.push(point);
      else if (dba === 'gap' && segment.length > 0) {
        segments.push(segment);
        segment = [];
      }
    }
    if (segment.length > 0) segments.push(segment);
    const paths = segments.map((points) =>
      series.smooth && points.length > 2
        ? smoothPath(points)
        : points.map(([x, y], i) => `${i === 0 ? 'M' : 'L'}${px(x)},${px(y)}`).join(' '),
    );
    const dPath = paths.join(' ');
    let offset = 0;
    // Close each area against its own baseline, including stacked baselines.
    const areaPath = fill
      ? segments
          .map((points, i) => {
            const back = basePts
              .slice(offset, offset + points.length)
              .reverse()
              .map(([x, y]) => `L${px(x)},${px(y)}`)
              .join(' ');
            offset += points.length;
            return `${paths[i]} ${back} Z`;
          })
          .join(' ')
      : '';
    perSeries.push({ s, series, color, ptsRaw, pts, areaPath, dPath });
  }
  // Area fills are opaque in PowerPoint (the authored solidFill at full
  // alpha). Overlapping (non-stacked) areas paint back-to-front — series and
  // stroke together, as PowerPoint treats each series as one atomic unit —
  // so the FIRST authored series ends up on top and fully occludes a shorter
  // series behind it, fill AND outline. Stacked areas are disjoint cumulative
  // bands — they never overlap, so authored order is fine (and is what the
  // accumulation math assumes); line charts (fill=false) have no occlusion
  // concern either, so both keep authored order.
  const paintOrder = fill && !isStacked ? perSeries.slice().reverse() : perSeries;
  if (fill) {
    for (const { color, areaPath } of paintOrder) {
      out.push(`<path d="${areaPath}" fill="${color}" stroke="none"/>`);
    }
  }
  for (const { s, series, color, ptsRaw, pts, dPath } of paintOrder) {
    // Authored line width / dash on the series (<c:ser><c:spPr><a:ln>).
    const lineWPx = series.lineWidthEmu ? Math.max(0.3, series.lineWidthEmu / EMU_PER_PX) : 1.5;
    const dashAttr = series.lineDash
      ? (() => {
          const sw = lineWPx;
          const pat = DASH_PATTERNS[series.lineDash!];
          if (!pat) return '';
          const arr = pat
            .split(' ')
            .map((n) => (Number.parseFloat(n) * sw).toFixed(2))
            .join(' ');
          return ` stroke-dasharray="${arr}"`;
        })()
      : '';
    out.push(
      `<path d="${dPath}" fill="none" stroke="${series.lineColor ?? color}" stroke-width="${lineWPx.toFixed(2)}" stroke-linejoin="round" stroke-linecap="round"${dashAttr}/>`,
    );
    // Markers show only on the "Line with Markers" subtype
    // (<c:lineChart><c:marker val="1"/> → spec.lineMarkers) or when the
    // series authors an explicit symbol. Area charts (`fill`) never show
    // them by default, and `markerSymbol='none'` always hides. This keeps
    // plain imported line charts marker-free, matching PowerPoint.
    const explicitSymbol =
      series.markerSymbol !== undefined &&
      series.markerSymbol !== 'auto' &&
      series.markerSymbol !== 'none';
    if (series.markerSymbol !== 'none' && !fill && (spec.lineMarkers === true || explicitSymbol)) {
      const symbol = autoMarkerSymbol(series.markerSymbol, s);
      const size = series.markerSizePt ?? 5;
      const r = Math.max(1, size * 0.5);
      for (const [xp, yp] of pts) {
        out.push(seriesMarker(symbol, xp, yp, r, ...markerColors(series, color)));
      }
    }
    // Resolve each point independently so imported point labels can override
    // the series visibility, text and placement.
    for (let c = 0; c < N; c++) {
      const p = ptsRaw[c];
      if (p == null) continue;
      const v = series.values[c];
      if (v == null || !Number.isFinite(v)) continue;
      const labelText = cartesianDataLabelText(spec, s, c, v);
      if (!labelText) continue;
      const [xp, yp] = p;
      const pos = chartPointLabelOptions(spec, s, c).position;
      const lx = pos === 'l' ? xp - 6 : pos === 'r' ? xp + 6 : xp;
      const ly =
        pos === 'b' ? yp + 13 : pos === 'ctr' || pos === 'l' || pos === 'r' ? yp + 3 : yp - 5;
      const anchor = pos === 'l' ? 'end' : pos === 'r' ? 'start' : 'middle';
      out.push(
        chartText(
          f,
          lx,
          ly,
          `text-anchor="${anchor}" ${dataLabelTextAttrs(spec, s, '#374151', 9, false, c)}`,
          `${escapeXml(labelText)}`,
        ),
      );
    }
    // Trendline overlay per series (only meaningful on the clustered
    // layout — stacked already shows the cumulative shape).
    if (!isStacked && series.trendline) {
      const finiteXs: number[] = [];
      const finiteYs: number[] = [];
      for (const [xp, yp] of pts) {
        if (Number.isFinite(yp)) {
          finiteXs.push(xp);
          finiteYs.push(yp);
        }
      }
      if (finiteXs.length >= 2) {
        const tlColor = series.trendline.color ?? color;
        out.push(trendlinePath(f, finiteXs, finiteYs, series.trendline, tlColor));
      }
    }
  }
  // Drop lines + hi-low lines. Drop lines go from each data point of
  // the first series down to the value-axis baseline. Hi-low lines
  // span the highest and lowest series values at each category.
  if (!isStacked && (spec.dropLines || spec.hiLowLines) && spec.series.length > 0) {
    for (let c = 0; c < N; c++) {
      const xp = xAt(c);
      if (spec.dropLines) {
        const firstVal = spec.series[0]?.values[c];
        if (firstVal !== null && firstVal !== undefined && Number.isFinite(firstVal)) {
          const yp = f.plotY + f.plotH - ((firstVal - min) / range) * f.plotH;
          out.push(
            `<line x1="${px(xp)}" y1="${px(yp)}" x2="${px(xp)}" y2="${px(baseY)}" stroke="#9CA3AF" stroke-width="0.5" stroke-dasharray="2 2"/>`,
          );
        }
      }
      if (spec.hiLowLines) {
        let hiV = -Infinity;
        let loV = Infinity;
        for (const s of spec.series) {
          const v = s.values[c];
          if (v === null || v === undefined || !Number.isFinite(v)) continue;
          if (v > hiV) hiV = v;
          if (v < loV) loV = v;
        }
        if (hiV > loV) {
          const yHi = f.plotY + f.plotH - ((hiV - min) / range) * f.plotH;
          const yLo = f.plotY + f.plotH - ((loV - min) / range) * f.plotH;
          out.push(
            `<line x1="${px(xp)}" y1="${px(yHi)}" x2="${px(xp)}" y2="${px(yLo)}" stroke="#4B5563" stroke-width="1"/>`,
          );
        }
      }
    }
  }
  return out.join('');
};

const renderPieChart = (
  f: ChartFrame,
  spec: ReadChartSpec,
  colors: ReadonlyArray<string>,
  doughnut: boolean,
): string => {
  const series = spec.series[0];
  if (!series) return '';
  const values = series.values.map((v) => Math.max(0, v ?? 0));
  const total = values.reduce((a, b) => a + b, 0);
  if (total === 0) return '';
  const radius = Math.min(f.plotW, f.plotH) / 2 - 2;
  const cx = f.plotX + f.plotW / 2;
  const cy = f.plotY + f.plotH / 2;
  const innerR = doughnut ? radius * ((spec.holeSizePct ?? 55) / 100) : 0;
  // <c:firstSliceAng> rotates the start position clockwise from 12
  // o'clock. SVG angles run counterclockwise from the +x axis, so we
  // start at -π/2 (12 o'clock) and add the user-authored degrees.
  const startDeg = spec.firstSliceAngleDeg ?? 0;
  let acc = -Math.PI / 2 + (startDeg * Math.PI) / 180;
  const out: string[] = [];
  for (let i = 0; i < values.length; i++) {
    const v = values[i] ?? 0;
    const angle = (v / total) * 2 * Math.PI;
    const start = acc;
    const end = acc + angle;
    acc = end;
    const largeArc = angle > Math.PI ? 1 : 0;
    // <c:dPt><c:explosion val="N"/> shoves the slice outward along the
    // mid-angle by N% of the radius (PowerPoint clamps to a sane max
    // around 400, but in practice authors stay 0–60).
    const explPct = series.pointExplosions?.[i] ?? 0;
    const explOffset = (explPct / 100) * radius;
    const midAngle = (start + end) / 2;
    const sx = cx + Math.cos(midAngle) * explOffset;
    const sy = cy + Math.sin(midAngle) * explOffset;
    const ox1 = sx + radius * Math.cos(start);
    const oy1 = sy + radius * Math.sin(start);
    const ox2 = sx + radius * Math.cos(end);
    const oy2 = sy + radius * Math.sin(end);
    // Per-slice color: <c:dPt> override wins, then series.color, then
    // accent palette fallback.
    const dptColor = series.pointColors?.[i];
    const color = dptColor ?? series.color ?? colors[i % colors.length]!;
    if (doughnut) {
      const ix1 = sx + innerR * Math.cos(start);
      const iy1 = sy + innerR * Math.sin(start);
      const ix2 = sx + innerR * Math.cos(end);
      const iy2 = sy + innerR * Math.sin(end);
      const d = `M${px(ox1)},${px(oy1)} A${px(radius)},${px(radius)} 0 ${largeArc} 1 ${px(ox2)},${px(oy2)} L${px(ix2)},${px(iy2)} A${px(innerR)},${px(innerR)} 0 ${largeArc} 0 ${px(ix1)},${px(iy1)} Z`;
      out.push(`<path d="${d}" fill="${color}" stroke="#FFFFFF" stroke-width="0.6"/>`);
    } else {
      const d = `M${px(sx)},${px(sy)} L${px(ox1)},${px(oy1)} A${px(radius)},${px(radius)} 0 ${largeArc} 1 ${px(ox2)},${px(oy2)} Z`;
      out.push(`<path d="${d}" fill="${color}" stroke="#FFFFFF" stroke-width="0.6"/>`);
    }
    // Pie / doughnut data labels — value or percent at the slice midpoint.
    // Track the explosion offset so labels move with their slice. dLblPos
    // chooses the radial position:
    //   - ctr (default for pie): radius * 0.6 (or annulus mid for doughnut)
    //   - inEnd: just inside the slice's outer edge
    //   - outEnd: outside the slice (with a darker fill so it shows on
    //     the chart-area background)
    const labelMid = (start + end) / 2;
    const labelOptions = chartPointLabelOptions(spec, 0, i);
    const pos = labelOptions.position;
    let labelR: number;
    let labelFill = '#FFFFFF';
    if (pos === 'inEnd') {
      labelR = radius - 12;
    } else if (pos === 'outEnd') {
      labelR = radius + 12;
      labelFill = '#374151';
    } else {
      // ctr / bestFit / undefined — slice midline.
      labelR = doughnut ? (radius + innerR) / 2 : radius * 0.6;
    }
    const labelX = sx + labelR * Math.cos(labelMid);
    const labelY = sy + labelR * Math.sin(labelMid);
    // PowerPoint/LibreOffice order a pie/doughnut label as category, then
    // value, then percent (e.g. "Web — 48%"), not the reverse.
    const labels: string[] = [];
    if (labelOptions.showSeriesName && series.name) labels.push(series.name);
    if (labelOptions.showCategory) {
      const catLabel = spec.categories[i];
      if (catLabel) labels.push(catLabel);
    }
    if (labelOptions.showValue) labels.push(formatDataLabelValue(spec, 0, v, i));
    if (labelOptions.showPercent) labels.push(`${((v / total) * 100).toFixed(0)}%`);
    const labelText = labelOptions.text ?? labels.join(labelOptions.separator ?? ' ');
    if (labelText) {
      out.push(
        chartText(
          f,
          labelX,
          labelY,
          `text-anchor="middle" dominant-baseline="middle" ${dataLabelTextAttrs(spec, 0, labelFill, 10, true, i)}`,
          `${escapeXml(labelText)}`,
        ),
      );
    }
  }
  return out.join('');
};

// ---------------------------------------------------------------------------
// Scatter / radar / bubble plotters. Scatter and bubble plot xy(z) tuples
// against two value axes (reusing `renderValueAxis`); radar is polar.

// Padded [min, max] for a scatter / bubble value axis so extreme points
// don't sit on the plot edge. Aligns the bounds to the niceTicks step.
const scatterAxisBounds = (vals: ReadonlyArray<number>): { min: number; max: number } => {
  let min = Infinity;
  let max = -Infinity;
  for (const v of vals) {
    if (Number.isFinite(v)) {
      if (v < min) min = v;
      if (v > max) max = v;
    }
  }
  if (!Number.isFinite(min)) return { min: 0, max: 1 };
  if (min === max) return { min: min - 1, max: max + 1 };
  const ticks = niceTicks(min, max);
  const step = ticks.length >= 2 ? ticks[1]! - ticks[0]! : (max - min) / 4 || 1;
  return { min: Math.floor(min / step) * step, max: Math.ceil(max / step) * step };
};

// Pairs a scatter / bubble series' x-channel (`xValues`, or the 1-based
// index when absent) with its y-channel (`values`) and bubble size,
// dropping points whose x or y is null / non-finite (PowerPoint skips
// those too).
const xyPoints = (series: ChartSeries): Array<{ x: number; y: number; size: number }> => {
  const hasX = series.xValues !== undefined && series.xValues.length > 0;
  const n = Math.max(series.xValues?.length ?? 0, series.values.length);
  const out: Array<{ x: number; y: number; size: number }> = [];
  for (let i = 0; i < n; i++) {
    let x: number;
    if (hasX) {
      const xv = series.xValues![i];
      if (xv === null || xv === undefined || !Number.isFinite(xv)) continue;
      x = xv;
    } else {
      x = i + 1;
    }
    const y = series.values[i];
    if (y === null || y === undefined || !Number.isFinite(y)) continue;
    const sz = series.bubbleSizes?.[i];
    out.push({ x, y, size: sz != null && Number.isFinite(sz) ? sz : 0 });
  }
  return out;
};

// Two value axes (x horizontal, y vertical) shared by scatter / bubble.
// Honours the authored `<c:valAx><c:scaling>` min / max + numberFormat on
// the y axis; the x axis is auto-scaled (our model carries only one
// value-axis scaling block, which PowerPoint applies to the primary
// (y) axis).
const renderScatterAxes = (
  f: ChartFrame,
  spec: ReadChartSpec,
  xB: { min: number; max: number },
  yB: { min: number; max: number },
): string => {
  const yAxis: AxisSpec = {
    orientation: 'vertical',
    min: yB.min,
    max: yB.max,
    ...(spec.valueAxisLineHidden !== undefined ? { lineHidden: spec.valueAxisLineHidden } : {}),
    ...(spec.valueAxis?.numberFormat !== undefined
      ? { numberFormat: spec.valueAxis.numberFormat }
      : {}),
  };
  const xAxis: AxisSpec = {
    orientation: 'horizontal',
    min: xB.min,
    max: xB.max,
    ...(spec.categoryAxisLineHidden !== undefined
      ? { lineHidden: spec.categoryAxisLineHidden }
      : {}),
  };
  return (
    (spec.valueAxisHidden ? '' : renderValueAxis(f, yAxis)) +
    (spec.categoryAxisHidden ? '' : renderValueAxis(f, xAxis))
  );
};

const renderScatterChart = (
  f: ChartFrame,
  spec: ReadChartSpec,
  colors: ReadonlyArray<string>,
): string => {
  const perSeries = spec.series.map((s) => xyPoints(s));
  const allX: number[] = [];
  const allY: number[] = [];
  for (const pts of perSeries) {
    for (const p of pts) {
      allX.push(p.x);
      allY.push(p.y);
    }
  }
  if (allX.length === 0) return '';
  const xB = scatterAxisBounds(allX);
  const yB = scatterAxisBounds(allY);
  if (spec.categoryAxisScaling?.min !== undefined) xB.min = spec.categoryAxisScaling.min;
  if (spec.categoryAxisScaling?.max !== undefined) xB.max = spec.categoryAxisScaling.max;
  if (spec.valueAxis?.min !== undefined) yB.min = spec.valueAxis.min;
  if (spec.valueAxis?.max !== undefined) yB.max = spec.valueAxis.max;
  const xRange = xB.max - xB.min || 1;
  const yRange = yB.max - yB.min || 1;
  const projX = (x: number): number => f.plotX + ((x - xB.min) / xRange) * f.plotW;
  const projY = (y: number): number => f.plotY + f.plotH - ((y - yB.min) / yRange) * f.plotH;
  const out: string[] = [renderScatterAxes(f, spec, xB, yB)];
  // scatterStyle governs lines vs markers. Default (absent) = markers
  // only, matching PowerPoint's stock "Scatter" subtype.
  const style = spec.scatterStyle;
  const showLine =
    style === 'line' || style === 'lineMarker' || style === 'smooth' || style === 'smoothMarker';
  const smooth = style === 'smooth' || style === 'smoothMarker';
  const showMarker =
    style === undefined || style === 'marker' || style === 'lineMarker' || style === 'smoothMarker';
  for (let s = 0; s < spec.series.length; s++) {
    const series = spec.series[s]!;
    const pts = perSeries[s]!;
    if (pts.length === 0) continue;
    const color = series.color ?? colors[s % colors.length]!;
    const proj: Array<[number, number]> = pts.map((p) => [projX(p.x), projY(p.y)]);
    if (showLine && proj.length >= 2) {
      const d =
        smooth && proj.length > 2
          ? smoothPath(proj)
          : proj.map(([xp, yp], i) => `${i === 0 ? 'M' : 'L'}${px(xp)},${px(yp)}`).join(' ');
      const lineWPx = series.lineWidthEmu ? Math.max(0.3, series.lineWidthEmu / EMU_PER_PX) : 1.5;
      out.push(
        `<path d="${d}" fill="none" stroke="${series.lineColor ?? color}" stroke-width="${lineWPx.toFixed(2)}" stroke-linejoin="round" stroke-linecap="round"/>`,
      );
    }
    // markerSymbol='none' always hides; an explicit symbol always shows
    // (even under the line-only styles).
    const sym = series.markerSymbol;
    const drawMarker = sym === 'none' ? false : showMarker || (sym !== undefined && sym !== 'auto');
    if (drawMarker) {
      const r = Math.max(1.5, (series.markerSizePt ?? 5) * 0.5);
      const glyph = autoMarkerSymbol(sym, s);
      const [markerFill, markerLine] = markerColors(series, color);
      for (const [xp, yp] of proj) {
        out.push(seriesMarker(glyph, xp, yp, r, markerFill, markerLine));
      }
    }
  }
  return out.join('');
};

const renderBubbleChart = (
  f: ChartFrame,
  spec: ReadChartSpec,
  colors: ReadonlyArray<string>,
): string => {
  const perSeries = spec.series.map((s) => xyPoints(s));
  const allX: number[] = [];
  const allY: number[] = [];
  let maxSize = 0;
  for (const pts of perSeries) {
    for (const p of pts) {
      allX.push(p.x);
      allY.push(p.y);
      if (p.size > maxSize) maxSize = p.size;
    }
  }
  if (allX.length === 0) return '';
  const xB = scatterAxisBounds(allX);
  const yB = scatterAxisBounds(allY);
  if (spec.categoryAxisScaling?.min !== undefined) xB.min = spec.categoryAxisScaling.min;
  if (spec.categoryAxisScaling?.max !== undefined) xB.max = spec.categoryAxisScaling.max;
  if (spec.valueAxis?.min !== undefined) yB.min = spec.valueAxis.min;
  if (spec.valueAxis?.max !== undefined) yB.max = spec.valueAxis.max;
  const xRange = xB.max - xB.min || 1;
  const yRange = yB.max - yB.min || 1;
  const projX = (x: number): number => f.plotX + ((x - xB.min) / xRange) * f.plotW;
  const projY = (y: number): number => f.plotY + f.plotH - ((y - yB.min) / yRange) * f.plotH;
  const out: string[] = [renderScatterAxes(f, spec, xB, yB)];
  // The largest bubble's diameter is bubbleScale% (default 25%) of the
  // plot's smaller dimension. Per-point radius scales with sqrt(size) for
  // the default area mode (sizeRepresents='area'), linearly for 'width'.
  const scaleFraction = spec.bubbleScale !== undefined ? spec.bubbleScale / 100 : 0.25;
  const maxRadiusPx = Math.max(2, scaleFraction * Math.min(f.plotW, f.plotH) * 0.5);
  const areaMode = (spec.bubbleSizeRepresents ?? 'area') === 'area';
  const radiusFor = (size: number): number => {
    if (maxSize <= 0 || size <= 0) return Math.max(1.5, maxRadiusPx * 0.15);
    const frac = areaMode ? Math.sqrt(size / maxSize) : size / maxSize;
    return Math.max(1.5, maxRadiusPx * frac);
  };
  for (let s = 0; s < spec.series.length; s++) {
    const series = spec.series[s]!;
    const pts = perSeries[s]!;
    const color = series.color ?? colors[s % colors.length]!;
    for (const p of pts) {
      const cx = projX(p.x);
      const cy = projY(p.y);
      const r = radiusFor(p.size);
      out.push(
        `<circle cx="${px(cx)}" cy="${px(cy)}" r="${r.toFixed(2)}" fill="${color}" fill-opacity="0.55" stroke="${color}" stroke-width="0.75"/>`,
      );
    }
  }
  return out.join('');
};

const renderRadarChart = (
  f: ChartFrame,
  spec: ReadChartSpec,
  colors: ReadonlyArray<string>,
): string => {
  const N = pointCount(spec);
  if (N === 0 || spec.series.length === 0) return '';
  // Rings start at 0 (or the authored min) and span to the data max (or
  // the authored max).
  let max = -Infinity;
  let min = 0;
  for (const sr of spec.series) {
    for (const v of sr.values) {
      if (v !== null && Number.isFinite(v) && v > max) max = v;
    }
  }
  if (!Number.isFinite(max)) max = 1;
  if (spec.valueAxis?.min !== undefined) min = spec.valueAxis.min;
  if (spec.valueAxis?.max !== undefined) max = spec.valueAxis.max;
  if (max <= min) max = min + 1;
  const range = max - min;
  const cx = f.plotX + f.plotW / 2;
  const cy = f.plotY + f.plotH / 2;
  // Leave a margin around the rim for the category labels.
  const R = Math.max(4, Math.min(f.plotW, f.plotH) / 2 - 14);
  const angleAt = (i: number): number => -Math.PI / 2 + (i * 2 * Math.PI) / N;
  const out: string[] = [];
  // Value rings (polygon gridlines) at each nice tick + a tick label on
  // the top spoke.
  for (const t of niceTicks(min, max)) {
    if (t < min || t > max + 1e-9) continue;
    const rr = ((t - min) / range) * R;
    if (rr <= 0) continue;
    const ring: string[] = [];
    for (let i = 0; i < N; i++) {
      const a = angleAt(i);
      ring.push(`${px(cx + rr * Math.cos(a))},${px(cy + rr * Math.sin(a))}`);
    }
    out.push(
      `<polygon points="${ring.join(' ')}" fill="none" stroke="#E5E7EB" stroke-width="0.5"/>`,
    );
    out.push(
      chartText(
        f,
        cx + 3,
        cy - rr,
        `dominant-baseline="middle" font-family="sans-serif" font-size="8" fill="#9CA3AF"`,
        `${escapeXml(formatTick(t))}`,
      ),
    );
  }
  // Spokes + category labels.
  for (let i = 0; i < N; i++) {
    const a = angleAt(i);
    out.push(
      `<line x1="${px(cx)}" y1="${px(cy)}" x2="${px(cx + R * Math.cos(a))}" y2="${px(cy + R * Math.sin(a))}" stroke="#D1D5DB" stroke-width="0.5"/>`,
    );
    const cat = spec.categories[i] ?? String(i + 1);
    const lx = cx + (R + 8) * Math.cos(a);
    const ly = cy + (R + 8) * Math.sin(a);
    const cosA = Math.cos(a);
    const anchor = Math.abs(cosA) < 0.3 ? 'middle' : cosA > 0 ? 'start' : 'end';
    const label = cat.length > 12 ? `${cat.slice(0, 11)}…` : cat;
    out.push(
      chartText(
        f,
        lx,
        ly,
        `text-anchor="${anchor}" dominant-baseline="middle" ${axisTickAttrs(spec.categoryAxisLabelStyle)}`,
        `${escapeXml(label)}`,
      ),
    );
  }
  // Series polygons (closed). 'filled' fills the polygon at reduced
  // opacity; 'marker' draws a point glyph at each vertex.
  const filled = spec.radarStyle === 'filled';
  const showMarker = spec.radarStyle === 'marker';
  for (let s = 0; s < spec.series.length; s++) {
    const series = spec.series[s]!;
    const color = series.color ?? colors[s % colors.length]!;
    const proj: Array<[number, number]> = [];
    for (let i = 0; i < N; i++) {
      const v = series.values[i];
      if (v === null || v === undefined || !Number.isFinite(v)) continue;
      const rr = ((v - min) / range) * R;
      const a = angleAt(i);
      proj.push([cx + rr * Math.cos(a), cy + rr * Math.sin(a)]);
    }
    if (proj.length < 2) continue;
    const ptsStr = proj.map(([xp, yp]) => `${px(xp)},${px(yp)}`).join(' ');
    const lineWPx = series.lineWidthEmu ? Math.max(0.3, series.lineWidthEmu / EMU_PER_PX) : 1.8;
    out.push(
      filled
        ? `<polygon points="${ptsStr}" fill="${color}" fill-opacity="0.3" stroke="${series.lineColor ?? color}" stroke-width="${lineWPx.toFixed(2)}" stroke-linejoin="round"/>`
        : `<polygon points="${ptsStr}" fill="none" stroke="${series.lineColor ?? color}" stroke-width="${lineWPx.toFixed(2)}" stroke-linejoin="round"/>`,
    );
    if (showMarker) {
      const r = Math.max(1.5, (series.markerSizePt ?? 5) * 0.5);
      const glyph = autoMarkerSymbol(series.markerSymbol, s);
      const [markerFill, markerLine] = markerColors(series, color);
      for (const [xp, yp] of proj) {
        out.push(seriesMarker(glyph, xp, yp, r, markerFill, markerLine));
      }
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
  groupReflected: boolean,
): string | null => {
  let spec: ReadChartSpec | null = null;
  try {
    spec = getShapeChartSpec(shape);
  } catch {
    return null;
  }
  if (!spec) return null;
  // There is no candlestick or surface plotter: a stock chart draws as its
  // [open,] high / low / close lines and a surface as columns, which keeps
  // the data legible instead of dropping the chart.
  if (spec.kind === 'stock') spec = { ...spec, kind: 'line' };
  else if (spec.kind === 'surface') spec = { ...spec, kind: 'column' };
  const colors = accentSequence(theme);
  const isCartesian =
    spec.kind === 'column' || spec.kind === 'bar' || spec.kind === 'line' || spec.kind === 'area';
  // Scatter / bubble draw two value axes (handled inside their plotters)
  // but still need the axis gutters; radar is polar and uses the full
  // plot box. The shared category-axis block below stays gated on
  // `isCartesian` so it doesn't run for the xy(z) kinds.
  const hasAxes = isCartesian || spec.kind === 'scatter' || spec.kind === 'bubble';
  // No `<c:legend>` in the chart XML means no legend — PowerPoint and
  // LibreOffice both render none. An authored legend with `position: null`
  // (`<c:delete/>`-style) is also hidden.
  const hasLegend = spec.legend !== undefined && spec.legend.position !== null;
  const flip = getShapeFlip(shape);
  const f: ChartFrame = {
    reflected: groupReflected !== Boolean(flip && flip.horizontal !== flip.vertical),
    ...layoutChart(
      x,
      y,
      w,
      h,
      !!spec.title,
      hasAxes,
      spec.titleOverlay ?? false,
      spec.legend?.overlay ?? false,
      hasLegend,
      (spec.titleStyle?.sizePt ?? DEFAULT_CHART_TITLE_PT) * PX_PER_PT,
      spec.plotAreaLayout,
    ),
  };
  const allNamesForLegend: string[] =
    spec.kind === 'pie' || spec.kind === 'doughnut'
      ? Array.from(spec.categories)
      : spec.series.map((s) => s.name);
  const allColorsForLegend: string[] =
    spec.kind === 'pie' || spec.kind === 'doughnut'
      ? spec.categories.map(
          (_, i) =>
            spec.series[0]?.pointColors?.[i] ??
            spec.series[0]?.color ??
            colors[i % colors.length] ??
            '#888',
        )
      : spec.series.map((s, i) => s.color ?? colors[i % colors.length] ?? '#888');
  // `<c:legendEntry><c:delete val="1"/>` hides specific series indices
  // from the legend — typically trendline series. Filter the parallel
  // arrays in lock-step so colors / names / markers stay aligned.
  const hiddenSet = new Set(spec.legend?.hiddenIndices ?? []);
  const seriesNamesForLegend = allNamesForLegend.filter((_, i) => !hiddenSet.has(i));
  const seriesColorsForLegend = allColorsForLegend.filter((_, i) => !hiddenSet.has(i));
  // The legend shows marker glyphs only when the chart actually plots them:
  // a line chart of the "with markers" subtype, or any series authoring an
  // explicit symbol. A plain line, an area chart, or bar/pie get the color
  // swatch instead (matching their plotted appearance).
  const explicitMarker = (sym: ChartSeries['markerSymbol']): boolean =>
    sym !== undefined && sym !== 'auto' && sym !== 'none';
  const showsMarkers =
    spec.kind === 'line' &&
    (spec.lineMarkers === true || spec.series.some((s) => explicitMarker(s.markerSymbol)));
  const markersForLegend = showsMarkers
    ? spec.series.filter((_, i) => !hiddenSet.has(i))
    : undefined;

  // Count finite values across all series — when zero, draw a hint
  // label so an empty chart isn't indistinguishable from a working
  // one whose data we failed to read.
  let finiteCount = 0;
  for (const s of spec.series) {
    for (const v of s.values) if (v !== null && Number.isFinite(v)) finiteCount++;
  }

  let plot = '';
  let axes = '';
  // Combo chart: per-series kind overrides and/or a secondary value
  // axis. Split the series into plot groups, bake palette colors by
  // original index (filtered specs would otherwise re-index), scale
  // each axis from its own series, and paint bars below lines. The
  // horizontal `bar` base kind is excluded — its value axis is
  // horizontal and a line overlay has no meaningful geometry.
  const isCombo =
    isCartesian &&
    spec.kind !== 'bar' &&
    spec.series.some(
      (s) => (s.chartKind !== undefined && s.chartKind !== spec.kind) || s.secondaryAxis === true,
    );
  if (isCombo) {
    const baked = spec.series.map((s, i) => ({
      ...s,
      color: s.color ?? colors[i % colors.length] ?? '#888',
    }));
    const primarySeries = baked.filter((s) => s.secondaryAxis !== true);
    const secondarySeries = baked.filter((s) => s.secondaryAxis === true);
    const primaryScale = seriesMinMax({ ...spec, series: primarySeries });
    // The authored valueAxis min/max targets the PRIMARY axis; the
    // secondary axis always auto-scales from its own series.
    const { valueAxis: _primaryOnlyAxis, ...specWithoutAxis } = spec;
    const secondaryScale =
      secondarySeries.length > 0
        ? seriesMinMax({ ...specWithoutAxis, series: secondarySeries })
        : null;

    const N = pointCount(spec);
    if (!spec.valueAxisHidden) {
      axes = renderValueAxis(f, {
        orientation: 'vertical',
        min: primaryScale.min,
        max: primaryScale.max,
        majorUnit: spec.valueAxis?.majorUnit ?? primaryScale.step,
        ...(spec.valueAxis?.minorUnit !== undefined ? { minorUnit: spec.valueAxis.minorUnit } : {}),
        ...(spec.valueAxisMinorGridlines !== undefined
          ? { minorGridlines: spec.valueAxisMinorGridlines }
          : {}),
        ...(spec.valueAxisMinorGridlineColor !== undefined
          ? { minorGridlineColor: spec.valueAxisMinorGridlineColor }
          : {}),
        ...(spec.valueAxisMinorGridlineWidthEmu !== undefined
          ? { minorGridlineWidthEmu: spec.valueAxisMinorGridlineWidthEmu }
          : {}),
        ...(spec.valueAxisLineHidden !== undefined ? { lineHidden: spec.valueAxisLineHidden } : {}),
        ...(spec.valueAxisLineColor !== undefined ? { lineColor: spec.valueAxisLineColor } : {}),
        ...(spec.valueAxisMajorTickMark !== undefined
          ? { majorTickMark: spec.valueAxisMajorTickMark }
          : {}),
        ...(spec.valueAxis?.numberFormat !== undefined
          ? { numberFormat: spec.valueAxis.numberFormat }
          : {}),
        ...(spec.valueAxisMajorGridlines !== undefined
          ? { majorGridlines: spec.valueAxisMajorGridlines }
          : {}),
        ...(spec.valueAxisLabelStyle !== undefined ? { labelStyle: spec.valueAxisLabelStyle } : {}),
      });
    }
    if (secondaryScale) {
      axes += renderValueAxis(f, {
        orientation: 'vertical',
        side: 'right',
        min: secondaryScale.min,
        max: secondaryScale.max,
        majorUnit: secondaryScale.step,
        // Gridlines stay on the primary axis only — a second lattice
        // with a different pitch reads as noise.
        majorGridlines: false,
        ...(spec.valueAxisLabelStyle !== undefined ? { labelStyle: spec.valueAxisLabelStyle } : {}),
      });
    }
    if (N > 0 && !(spec.categoryAxisHidden || spec.categoryAxisTickLabelPos === 'none')) {
      axes += renderCategoryAxis(
        f,
        'horizontal',
        spec.categories,
        N,
        spec.categoryAxisTickLabelSkip ?? 1,
        spec.categoryAxisLabelStyle,
        spec.categoryAxisLabelRotationDeg,
        spec.categoryAxisLabelAlign,
        spec.categoryAxisLineColor,
        spec.categoryAxisLineHidden,
        spec.categoryAxisMajorTickMark,
        spec.categoryAxisMinorTickMark,
      );
    }

    // Group by (effective kind, axis); bars first, then line/area overlays.
    const groups = new Map<
      string,
      { kind: ChartKind; secondary: boolean; series: ChartSeries[] }
    >();
    for (const s of baked) {
      const kind = s.chartKind ?? spec.kind;
      const secondary = s.secondaryAxis === true;
      const key = `${kind}|${secondary ? '1' : '0'}`;
      const group = groups.get(key);
      if (group) group.series.push(s);
      else groups.set(key, { kind, secondary, series: [s] });
    }
    const paintOrder = (g: { kind: ChartKind; secondary: boolean }): number =>
      (g.secondary ? 2 : 0) + (g.kind === 'line' || g.kind === 'area' ? 1 : 0);
    for (const group of [...groups.values()].sort((a, b) => paintOrder(a) - paintOrder(b))) {
      const scale = group.secondary && secondaryScale ? secondaryScale : primaryScale;
      const groupSpec: ReadChartSpec = {
        ...spec,
        series: group.series,
        valueAxis: { min: scale.min, max: scale.max },
      };
      if (group.kind === 'line' || group.kind === 'area') {
        plot += renderLineChart(f, groupSpec, colors, group.kind === 'area');
      } else {
        plot += renderColumnChart(f, groupSpec, colors);
      }
    }
  }
  if (!isCombo && isCartesian) {
    const { min, max, step } = seriesMinMax(spec);
    const N = pointCount(spec);
    const majorUnit = spec.valueAxis?.majorUnit ?? step;
    const numberFormat = spec.valueAxis?.numberFormat;
    const axisExtras = {
      ...(spec.valueAxis?.minorUnit !== undefined ? { minorUnit: spec.valueAxis.minorUnit } : {}),
      ...(spec.valueAxisMinorGridlines !== undefined
        ? { minorGridlines: spec.valueAxisMinorGridlines }
        : {}),
      ...(spec.valueAxisMinorGridlineColor !== undefined
        ? { minorGridlineColor: spec.valueAxisMinorGridlineColor }
        : {}),
      ...(spec.valueAxisMinorGridlineWidthEmu !== undefined
        ? { minorGridlineWidthEmu: spec.valueAxisMinorGridlineWidthEmu }
        : {}),
      ...(spec.valueAxisLineHidden !== undefined ? { lineHidden: spec.valueAxisLineHidden } : {}),
      ...(majorUnit !== undefined ? { majorUnit } : {}),
      ...(numberFormat !== undefined ? { numberFormat } : {}),
      ...(spec.valueAxisMajorGridlines !== undefined
        ? { majorGridlines: spec.valueAxisMajorGridlines }
        : {}),
      ...(spec.valueAxisLabelStyle !== undefined ? { labelStyle: spec.valueAxisLabelStyle } : {}),
      ...(spec.valueAxisMajorGridlineColor !== undefined
        ? { majorGridlineColor: spec.valueAxisMajorGridlineColor }
        : {}),
      ...(spec.valueAxisMajorTickMark !== undefined
        ? { majorTickMark: spec.valueAxisMajorTickMark }
        : {}),
      ...(spec.valueAxisLineColor !== undefined ? { lineColor: spec.valueAxisLineColor } : {}),
      ...(spec.valueAxisLabelRotationDeg !== undefined
        ? { labelRotationDeg: spec.valueAxisLabelRotationDeg }
        : {}),
      ...(spec.valueAxis?.displayUnits !== undefined
        ? { displayUnits: spec.valueAxis.displayUnits }
        : {}),
    };
    const isPercentStacked = spec.grouping === 'percentStacked';
    const valueAxis: AxisSpec =
      spec.kind === 'bar'
        ? { orientation: 'horizontal', min, max, percent: isPercentStacked, ...axisExtras }
        : { orientation: 'vertical', min, max, percent: isPercentStacked, ...axisExtras };
    if (!spec.valueAxisHidden) axes = renderValueAxis(f, valueAxis);
    // tickLblPos='none' hides the labels but keeps the gridline (we
    // already conditionally skip below). Use the explicit skip step
    // when authored.
    const labelsHidden = spec.categoryAxisHidden || spec.categoryAxisTickLabelPos === 'none';
    if (N > 0 && !labelsHidden) {
      axes += renderCategoryAxis(
        f,
        spec.kind === 'bar' ? 'vertical' : 'horizontal',
        spec.categories,
        N,
        spec.categoryAxisTickLabelSkip ?? 1,
        spec.categoryAxisLabelStyle,
        spec.categoryAxisLabelRotationDeg,
        spec.categoryAxisLabelAlign,
        spec.categoryAxisLineColor,
        spec.categoryAxisLineHidden,
        spec.categoryAxisMajorTickMark,
        spec.categoryAxisMinorTickMark,
      );
    }
  }
  if (!isCombo) {
    switch (spec.kind) {
      case 'column':
      case 'bar':
        // @office-kit/pptx reports both as `bar` / `column` via separate `kind`;
        // legacy `barDir` distinction. We branch on `kind`.
        plot =
          spec.kind === 'column'
            ? renderColumnChart(f, spec, colors)
            : renderBarChart(f, spec, colors);
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
      case 'scatter':
        plot = renderScatterChart(f, spec, colors);
        break;
      case 'radar':
        plot = renderRadarChart(f, spec, colors);
        break;
      case 'bubble':
        plot = renderBubbleChart(f, spec, colors);
        break;
      default:
        // stock / surface were mapped to a drawable kind above, and 3-D
        // variants read as their flat kind; truly unmodeled kinds (resolved
        // to `null` spec) are handled earlier. Anything left falls back.
        return null;
    }
  }

  const emptyHint =
    finiteCount === 0
      ? chartText(
          f,
          f.plotX + f.plotW / 2,
          f.plotY + f.plotH / 2,
          `text-anchor="middle" dominant-baseline="middle" font-family="sans-serif" font-size="12" fill="#9CA3AF"`,
          `${escapeXml(`chart (${spec.kind}) — no data`)}`,
        )
      : '';

  // Axis titles — value title is rotated -90° to read along the y-axis
  // tick stack; category title sits below the x-axis. Authored
  // <a:rPr sz/b/i> + solidFill override the 11pt semibold default.
  const axisTitleAttrs = (style: ChartTextStyle | undefined): string => {
    const sz = style?.sizePt ?? 11;
    const fill = style?.color ?? '#374151';
    const weight = style?.bold === false ? '400' : '600';
    const italicAttr = style?.italic ? ' font-style="italic"' : '';
    return `font-family="sans-serif" font-size="${chartFontPx(sz)}" fill="${fill}" font-weight="${weight}"${italicAttr}`;
  };
  // Authored <c:title><c:tx><c:rich><a:bodyPr rot> overrides the
  // renderer's defaults (-90 for the value-axis title, 0 for the
  // category-axis title). Pivot stays at the label anchor so the title
  // hugs its axis.
  const valueAxisTitleRot = spec.valueAxisTitleRotationDeg ?? -90;
  const valueAxisTitleSvg = spec.valueAxisTitle
    ? chartText(
        f,
        f.plotX - 26,
        f.plotY + f.plotH / 2,
        `text-anchor="middle" ${axisTitleAttrs(spec.valueAxisTitleStyle)} transform="rotate(${valueAxisTitleRot} ${px(f.plotX - 26)} ${px(f.plotY + f.plotH / 2)})"`,
        `${escapeXml(spec.valueAxisTitle)}`,
      )
    : '';
  const catTitleRot = spec.categoryAxisTitleRotationDeg ?? 0;
  const catTitleCx = f.plotX + f.plotW / 2;
  // The tick-label row's baseline sits at plotY + plotH + 12 (renderCategoryAxis);
  // its own ascent/descent plus the title's ascent need clearance below that so
  // the two text rows don't collide, hence +28 here rather than the tick row's
  // near-identical +16.
  const catTitleCy = f.plotY + f.plotH + 28;
  const catTitleTransform =
    catTitleRot !== 0
      ? ` transform="rotate(${catTitleRot} ${px(catTitleCx)} ${px(catTitleCy)})"`
      : '';
  const categoryAxisTitleSvg = spec.categoryAxisTitle
    ? chartText(
        f,
        catTitleCx,
        catTitleCy,
        `text-anchor="middle" ${axisTitleAttrs(spec.categoryAxisTitleStyle)}${catTitleTransform}`,
        `${escapeXml(spec.categoryAxisTitle)}`,
      )
    : '';
  return [
    `<g${transform}>`,
    // Chart-area backdrop honors <c:chartSpace><c:spPr><a:solidFill> /
    // <a:ln>. plot-area gets its own tinted rect + border when
    // <c:plotArea><c:spPr> authors them. PowerPoint draws no chart-area
    // border unless the chartSpace authors an <a:ln>, so the default is
    // `none` — an invented light-gray frame is the most visible single
    // divergence from PowerPoint's actual rendering.
    `<rect x="${px(f.x)}" y="${px(f.y)}" width="${px(f.w)}" height="${px(f.h)}" fill="${spec.chartAreaFill ?? '#FFFFFF'}" stroke="${spec.chartAreaStrokeColor ?? 'none'}" stroke-width="0.6"${spec.roundedCorners ? ' rx="6" ry="6"' : ''}/>`,
    spec.plotAreaFill || spec.plotAreaStrokeColor
      ? `<rect x="${px(f.plotX)}" y="${px(f.plotY)}" width="${px(f.plotW)}" height="${px(f.plotH)}" fill="${spec.plotAreaFill ?? 'none'}" stroke="${spec.plotAreaStrokeColor ?? 'none'}" stroke-width="0.6"/>`
      : '',
    renderChartTitle(f, spec.title ?? '', spec.titleStyle),
    axes,
    valueAxisTitleSvg,
    categoryAxisTitleSvg,
    plot,
    emptyHint,
    hasLegend
      ? renderChartLegend(
          f,
          seriesNamesForLegend,
          seriesColorsForLegend,
          spec.legend?.position ?? 'b',
          spec.legend?.textStyle,
          // Marker glyphs only carry visual meaning for line / area
          // charts; bar / column / pie use the swatch rect.
          markersForLegend,
        )
      : '',
    '</g>',
  ].join('');
};

// ---------------------------------------------------------------------------
// Table rendering. Draws the grid, cell fills, banded rows, header row, merged
// cells, and per-run cell text. Cell text reuses the same layout building
// blocks as shape text (buildAndLayoutSvgText / renderRun), so it inherits
// per-run size / bold / italic / color / font, paragraph alignment, line
// wrapping, and the svg ↔ foreignObject split for free.

// Builds the per-paragraph layout model the shared text engine consumes from a
// cell's structured paragraphs and shared DrawingML paragraph properties. A run's effective point
// size resolves to its explicit `<a:rPr sz>` when present, else the table-cell
// default: @office-kit/pptx doesn't model `<a:tblStyle>` text props, so unstyled cells
// fall to PowerPoint's authored default cell size (18 pt — what it writes for a
// freshly inserted table) in the theme's minor font and the cell's text color.
const cellParaData = (
  paragraphs: ReadonlyArray<TableCellParagraph>,
  cell: Parameters<typeof getTableCellParagraphs>[0],
  pres: PresentationData,
  shape: SlideShapeData,
): { paraData: ParaData[]; hasText: boolean } => {
  let hasText = false;
  const paraData = paragraphs.map((para, index): ParaData => {
    const properties = getParagraphPropertiesEffective(pres, cell, index);
    const bulletIsPicture = isParagraphBulletPicture(cell, index);
    const bulletImageBytes = bulletIsPicture ? getParagraphBulletImageBytes(cell, index) : null;
    const runs: RunData[] = [];
    for (const el of para.elements) {
      if (el.kind === 'br') {
        runs.push({ text: '\n', fmt: null, sizePt: DEFAULT_BODY_PT });
        continue;
      }
      if (el.text.trim()) hasText = true;
      const href = el.clickAction ? clickActionHref(pres, shape, el.clickAction) : null;
      runs.push({
        text: el.text,
        fmt: el.format,
        sizePt: el.format?.size ?? DEFAULT_BODY_PT,
        ...(href ? { href, ...(el.tooltip !== undefined ? { hrefTip: el.tooltip } : {}) } : {}),
      });
    }
    return {
      align: properties.align ?? 'left',
      level: properties.level,
      bulletStyle: properties.bullet,
      bulletDetail: getParagraphBulletStyle(pres, cell, index),
      bulletIsPicture,
      bulletImageHref: bulletImageBytes ? bytesToDataUrl(bulletImageBytes) : null,
      runs,
      lineSpacing: properties.lineSpacing,
      spcBefPts: properties.spcBefPts,
      spcAftPts: properties.spcAftPts,
      indent: {
        leftEmu: properties.marL,
        rightEmu: properties.marR,
        firstLineEmu: properties.indent,
      },
    };
  });
  return { paraData, hasText };
};

const renderTableCellText = (
  cell: Parameters<typeof getTableCellParagraphs>[0],
  paragraphs: ReadonlyArray<TableCellParagraph>,
  cx: number,
  cy: number,
  cw: number,
  ch: number,
  color: string,
  pres: PresentationData,
  shape: SlideShapeData,
  theme: PresentationTheme | null,
  themeFace: string | null,
  ctx: LayoutCtx,
  vAnchor: 'top' | 'center' | 'bottom',
  margins: {
    left: number | null;
    right: number | null;
    top: number | null;
    bottom: number | null;
  },
): string => {
  const { paraData, hasText } = cellParaData(paragraphs, cell, pres, shape);
  if (!hasText) return '';
  const numberLabels = paragraphNumberLabels(paraData);
  // CT_TableCellProperties defaults: 0.1 inch horizontally, 0.05 vertically.
  const padL = (margins.left ?? 91440) / EMU_PER_PX;
  const padR = (margins.right ?? 91440) / EMU_PER_PX;
  const padT = (margins.top ?? 45720) / EMU_PER_PX;
  const padB = (margins.bottom ?? 45720) / EMU_PER_PX;
  const innerX = cx + padL;
  const innerY = cy + padT;
  const innerW = Math.max(0, cw - padL - padR);
  const innerH = Math.max(0, ch - padT - padB);
  if (innerW <= 0 || innerH <= 0) return '';

  // Pure-SVG path: hand the layout model to the same engine shape text uses so
  // a browser-free rasterizer gets wrapped, per-run-styled lines. The engine
  // works in EMU (it divides by EMU_PER_PX internally), so project the px box
  // back to EMU.
  if (ctx.mode === 'svg') {
    return buildAndLayoutSvgText({
      pres,
      shape,
      theme,
      paraData,
      numberLabels,
      autoFitScale: 1,
      lineHeightScale: 1,
      defaultPt: DEFAULT_BODY_PT,
      themeFace,
      defaultColor: color,
      anchor: vAnchor,
      wrap: true,
      innerX: innerX * EMU_PER_PX,
      innerY: innerY * EMU_PER_PX,
      innerW: innerW * EMU_PER_PX,
      innerH: innerH * EMU_PER_PX,
      measure: ctx.measure,
      // Cell-level vertical text (<a:tcPr vert>) isn't modeled yet; cells lay
      // out horizontally, single-column.
      vert: 'none',
      columns: null,
    });
  }

  // foreignObject path (browser preview): one <p> per paragraph, each run
  // rendered through renderRun so the browser lays the styled text out.
  const justify = vAnchor === 'top' ? 'flex-start' : vAnchor === 'bottom' ? 'flex-end' : 'center';
  const familyFont = themeFace ? `${escapeXml(themeFace)}, ${DEFAULT_FONT}` : DEFAULT_FONT;
  const body = renderHtmlParagraphs(paraData, numberLabels, theme, 1, DEFAULT_BODY_PT, color).join(
    '',
  );
  return `<foreignObject x="${px(innerX)}" y="${px(innerY)}" width="${px(innerW)}" height="${px(innerH)}"><div xmlns="http://www.w3.org/1999/xhtml" style="display:flex;flex-direction:column;justify-content:${justify};width:100%;height:100%;box-sizing:border-box;overflow:hidden;line-height:1.2;font-family:${familyFont};color:${color};word-break:break-word">${body}</div></foreignObject>`;
};

const renderTable = (
  shape: SlideShapeData,
  pres: PresentationData,
  x: number,
  y: number,
  w: number,
  h: number,
  transform: string,
  theme: PresentationTheme | null,
  ctx: LayoutCtx,
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

  const flip = getShapeFlip(shape);
  const textReflected = ctx.groupReflected !== Boolean(flip && flip.horizontal !== flip.vertical);

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
  const wScale = wSum > 0 ? w / EMU_PER_PX / wSum : 1;
  const hScale = hSum > 0 ? h / EMU_PER_PX / hSum : 1;
  const colXs: number[] = [xPx];
  for (let c = 0; c < widthsPx.length; c++) {
    colXs.push((colXs[c] ?? xPx) + (widthsPx[c] ?? 0) * wScale);
  }
  const rowYs: number[] = [yPx];
  for (let r = 0; r < heightsPx.length; r++) {
    rowYs.push((rowYs[r] ?? yPx) + (heightsPx[r] ?? 0) * hScale);
  }

  // A10 table style — header / footer / first-col / last-col / banded
  // rows / banded columns. Project the boolean flags onto per-cell tints
  // that approximate the theme-driven look in PowerPoint.
  const flags = getTableStyleFlags(shape);
  const accent = theme ? normalizeHex(theme.accent1) : '#4472C4';
  const headerFill = accent;
  // Pale tints for banded rows/cols/first-col/last-col — `t` is the accent's
  // weight, so a *light* tint needs a *low* t (mostly white). PowerPoint/
  // LibreOffice's built-in styles alternate TWO tints across body rows (no
  // row is left unshaded), not one tint vs. no fill.
  const bandFill = mixHex(accent, '#FFFFFF', 0.12);
  const bandFill2 = mixHex(accent, '#FFFFFF', 0.27);
  // Fallback for cells with no authored color — the deck's body-text color
  // (an inverted map would paint the `tx1` token white-on-white).
  const textColor = activeDeckTextColor;
  // Default table-cell face is the theme's body (minor) font, the same default
  // PowerPoint applies when a cell run carries no explicit typeface.
  const tableThemeFace = getPresentationFonts(pres)?.minorLatin ?? null;
  const out: string[] = [];
  out.push(`<g${transform}>`);
  // Whole-table backdrop so cells with no explicit fill still
  // contrast against whatever's behind.
  out.push(
    `<rect x="${px(xPx)}" y="${px(yPx)}" width="${px((colXs[widthsPx.length] ?? xPx) - xPx)}" height="${px((rowYs[heightsPx.length] ?? yPx) - yPx)}" fill="#FFFFFF"/>`,
  );
  const borderEdges: string[] = [];
  for (let r = 0; r < dims.rows; r++) {
    for (let c = 0; c < dims.cols; c++) {
      const cell = cells[r]?.[c];
      if (!cell) continue;
      const typedCell = cell as Parameters<typeof getTableCellSpan>[0];
      const span = getTableCellSpan(typedCell);
      // Skip cells absorbed by a horizontal or vertical merge — their
      // visual area is painted by the cell that owns the span.
      if (span.hMerge || span.vMerge) continue;
      const cx = colXs[c] ?? xPx;
      const cy = rowYs[r] ?? yPx;
      const endCol = Math.min(dims.cols, c + span.gridSpan);
      const endRow = Math.min(dims.rows, r + span.rowSpan);
      const cw = (colXs[endCol] ?? cx) - cx;
      const ch = (rowYs[endRow] ?? cy) - cy;
      const fill = getTableCellFill(cell as Parameters<typeof getTableCellFill>[0]);
      let resolvedFill: string;
      if (fill) {
        resolvedFill = resolveColor(fill, theme, '#FFFFFF');
      } else if (flags.firstRow && r === 0) {
        resolvedFill = headerFill;
      } else if (flags.lastRow && r === dims.rows - 1) {
        resolvedFill = headerFill;
      } else if (flags.firstCol && c === 0) {
        resolvedFill = bandFill;
      } else if (flags.lastCol && c === dims.cols - 1) {
        resolvedFill = bandFill;
      } else if (flags.bandRow) {
        // Alternation starts at the first body row (right after a header),
        // not at the raw grid row index — otherwise a firstRow table shifts
        // the whole band pattern by one row.
        const bandIndex = r - (flags.firstRow ? 1 : 0);
        resolvedFill = bandIndex % 2 === 0 ? bandFill : bandFill2;
      } else if (flags.bandCol) {
        const bandIndex = c - (flags.firstCol ? 1 : 0);
        resolvedFill = bandIndex % 2 === 0 ? bandFill : bandFill2;
      } else {
        resolvedFill = 'none';
      }
      const cellTextColor = resolvedFill === headerFill ? '#FFFFFF' : textColor;
      out.push(
        `<rect x="${px(cx)}" y="${px(cy)}" width="${px(cw)}" height="${px(ch)}" fill="${resolvedFill}"/>`,
      );
      // Per-side borders override the default thin gray grid. Draw them
      // separately after the fills so they sit on top.
      const borders = getTableCellBorders(pres, typedCell);
      // Project the OOXML `<a:prstDash>` token onto an SVG stroke-dasharray.
      // Scaled by the border's width so the dash visually matches PowerPoint.
      const dashAttr = (dash: string | null | undefined, widthPx: number): string => {
        if (!dash || dash === 'solid') return '';
        const pat = DASH_PATTERNS[dash];
        if (!pat) return '';
        const arr = pat
          .split(' ')
          .map((n) => (Number.parseFloat(n) * widthPx).toFixed(2))
          .join(' ');
        return ` stroke-dasharray="${arr}"`;
      };
      const edge = (
        side: keyof Pick<typeof borders, 'left' | 'right' | 'top' | 'bottom'>,
        x1: number,
        y1: number,
        x2: number,
        y2: number,
      ): void => {
        const b = borders[side];
        if (!b) return;
        const sw = b.widthEmu ? Math.max(0.4, b.widthEmu / EMU_PER_PX) : 0.5;
        const col = b.color ?? '#9CA3AF';
        borderEdges.push(
          `<line x1="${px(x1)}" y1="${px(y1)}" x2="${px(x2)}" y2="${px(y2)}" stroke="${col}" stroke-width="${px(sw)}"${dashAttr(b.dash, sw)}/>`,
        );
      };
      edge('left', cx, cy, cx, cy + ch);
      edge('right', cx + cw, cy, cx + cw, cy + ch);
      edge('top', cx, cy, cx + cw, cy);
      edge('bottom', cx, cy + ch, cx + cw, cy + ch);
      if (borders.tlToBr) {
        const sw = borders.tlToBr.widthEmu
          ? Math.max(0.4, borders.tlToBr.widthEmu / EMU_PER_PX)
          : 0.5;
        borderEdges.push(
          `<line x1="${px(cx)}" y1="${px(cy)}" x2="${px(cx + cw)}" y2="${px(cy + ch)}" stroke="${borders.tlToBr.color ?? '#9CA3AF'}" stroke-width="${px(sw)}"${dashAttr(borders.tlToBr.dash, sw)}/>`,
        );
      }
      if (borders.blToTr) {
        const sw = borders.blToTr.widthEmu
          ? Math.max(0.4, borders.blToTr.widthEmu / EMU_PER_PX)
          : 0.5;
        borderEdges.push(
          `<line x1="${px(cx)}" y1="${px(cy + ch)}" x2="${px(cx + cw)}" y2="${px(cy)}" stroke="${borders.blToTr.color ?? '#9CA3AF'}" stroke-width="${px(sw)}"${dashAttr(borders.blToTr.dash, sw)}/>`,
        );
      }
      // Default thin grid for sides that didn't define a border.
      const defaultColor = '#9CA3AF';
      if (!borders.left)
        borderEdges.push(
          `<line x1="${px(cx)}" y1="${px(cy)}" x2="${px(cx)}" y2="${px(cy + ch)}" stroke="${defaultColor}" stroke-width="0.4" opacity="0.6"/>`,
        );
      if (!borders.right)
        borderEdges.push(
          `<line x1="${px(cx + cw)}" y1="${px(cy)}" x2="${px(cx + cw)}" y2="${px(cy + ch)}" stroke="${defaultColor}" stroke-width="0.4" opacity="0.6"/>`,
        );
      if (!borders.top)
        borderEdges.push(
          `<line x1="${px(cx)}" y1="${px(cy)}" x2="${px(cx + cw)}" y2="${px(cy)}" stroke="${defaultColor}" stroke-width="0.4" opacity="0.6"/>`,
        );
      if (!borders.bottom)
        borderEdges.push(
          `<line x1="${px(cx)}" y1="${px(cy + ch)}" x2="${px(cx + cw)}" y2="${px(cy + ch)}" stroke="${defaultColor}" stroke-width="0.4" opacity="0.6"/>`,
        );

      const cellParagraphs = getTableCellParagraphs(
        cell as Parameters<typeof getTableCellParagraphs>[0],
      );
      // ECMA-376 default cell anchor is top ('t'), which is what PowerPoint /
      // LibreOffice render when `<a:tcPr anchor>` is absent.
      const vAnchor = getTableCellAnchor(typedCell) ?? 'top';
      const cellMargins = getTableCellMargins(typedCell);
      const cellText = renderTableCellText(
        typedCell,
        cellParagraphs,
        cx,
        cy,
        cw,
        ch,
        cellTextColor,
        pres,
        shape,
        theme,
        tableThemeFace,
        ctx,
        vAnchor,
        cellMargins,
      );
      // The table transform moves cells and their borders. Cancel reflection
      // around each cell's center so glyphs retain their reading direction,
      // including when a parent group contributes another reflection.
      out.push(
        textReflected && cellText
          ? `<g transform="translate(${px(2 * cx + cw)} 0) scale(-1 1)">${cellText}</g>`
          : cellText,
      );
    }
  }
  out.push(borderEdges.join(''));
  out.push('</g>');
  return out.join('');
};

interface ArcCubic {
  readonly c1x: number;
  readonly c1y: number;
  readonly c2x: number;
  readonly c2y: number;
  readonly ex: number;
  readonly ey: number;
}

/**
 * Converts an OOXML `<a:arcTo>` (center-parameterized: start at the pen
 * point, angle `stAng`, sweep `swAng`, both in 60000ths of a degree, on a
 * `wR`×`hR` axis-aligned ellipse) into cubic-Bézier segments in the path's
 * own coordinate space.
 *
 * We approximate with cubics rather than emit an SVG `A` arc because the
 * endpoint-parameterized `A` command can't represent a sweep ≥ 360° (full
 * ellipses appear in custom geometry) and forces large-arc / sweep flag
 * bookkeeping for the signed `swAng`; splitting into ≤ 90° cubic segments
 * sidesteps both and stays accurate.
 */
const arcToCubicSegments = (
  startX: number,
  startY: number,
  wR: number,
  hR: number,
  stAng: number,
  swAng: number,
): ArcCubic[] => {
  const toRad = (a: number): number => (a / 60_000) * (Math.PI / 180);
  const st = toRad(stAng);
  const sw = toRad(swAng);
  // Place the ellipse center so the current pen point lies on it at `st`.
  const cx = startX - wR * Math.cos(st);
  const cy = startY - hR * Math.sin(st);
  const segCount = Math.max(1, Math.ceil(Math.abs(sw) / (Math.PI / 2)));
  const delta = sw / segCount;
  // Tangent-based control-handle length for a circular-arc cubic, applied
  // per-axis so it works for the anisotropic (wR ≠ hR) ellipse.
  const k = (4 / 3) * Math.tan(delta / 4);
  const segs: ArcCubic[] = [];
  let a0 = st;
  for (let i = 0; i < segCount; i++) {
    const a1 = a0 + delta;
    const cos0 = Math.cos(a0);
    const sin0 = Math.sin(a0);
    const cos1 = Math.cos(a1);
    const sin1 = Math.sin(a1);
    const p1x = cx + wR * cos1;
    const p1y = cy + hR * sin1;
    segs.push({
      c1x: cx + wR * cos0 - k * wR * sin0,
      c1y: cy + hR * sin0 + k * hR * cos0,
      c2x: p1x + k * wR * sin1,
      c2y: p1y - k * hR * cos1,
      ex: p1x,
      ey: p1y,
    });
    a0 = a1;
  }
  return segs;
};

/**
 * Renders an evaluated {@link CustomGeometry} as one `<path>` per
 * `GeomPath`, scaling each path's coordinate space (its own `w`/`h`, or
 * the shape extents when those are absent) onto the shape's slide box
 * `(x, y, w, h)` in EMU. Returns `''` when nothing drawable remains, so
 * the caller can fall back to the labelled rect.
 *
 * Per-path `fill="none"` paints stroke-only; `stroke="0"` paints
 * fill-only — mirroring `ST_PathFillMode` / the path `stroke` flag.
 */
const customGeometryToSvg = (
  geom: CustomGeometry,
  x: number,
  y: number,
  w: number,
  h: number,
  fill: string,
  fillExtra: string,
  stroke: string,
  strokeWidthEmu: number,
  strokeExtra: string,
  markerExtra: string,
): string => {
  const out: string[] = [];
  for (const path of geom.paths) {
    const cw = path.w ?? w;
    const ch = path.h ?? h;
    // A zero-extent coordinate space can't be scaled (would divide by 0).
    if (cw === 0 || ch === 0) continue;
    const sx = w / cw;
    const sy = h / ch;
    const fx = (gx: number): string => ((x + gx * sx) / EMU_PER_PX).toFixed(2);
    const fy = (gy: number): string => ((y + gy * sy) / EMU_PER_PX).toFixed(2);
    const pt = (gx: number, gy: number): string => `${fx(gx)},${fy(gy)}`;

    let curX = 0;
    let curY = 0;
    let startX = 0;
    let startY = 0;
    const d: string[] = [];
    for (const cmd of path.commands) {
      switch (cmd.kind) {
        case 'moveTo':
          curX = cmd.pt.x;
          curY = cmd.pt.y;
          startX = curX;
          startY = curY;
          d.push(`M${pt(curX, curY)}`);
          break;
        case 'lnTo':
          curX = cmd.pt.x;
          curY = cmd.pt.y;
          d.push(`L${pt(curX, curY)}`);
          break;
        case 'quadBezTo':
          d.push(`Q${pt(cmd.pts[0].x, cmd.pts[0].y)} ${pt(cmd.pts[1].x, cmd.pts[1].y)}`);
          curX = cmd.pts[1].x;
          curY = cmd.pts[1].y;
          break;
        case 'cubicBezTo':
          d.push(
            `C${pt(cmd.pts[0].x, cmd.pts[0].y)} ${pt(cmd.pts[1].x, cmd.pts[1].y)} ${pt(cmd.pts[2].x, cmd.pts[2].y)}`,
          );
          curX = cmd.pts[2].x;
          curY = cmd.pts[2].y;
          break;
        case 'arcTo': {
          const segs = arcToCubicSegments(curX, curY, cmd.wR, cmd.hR, cmd.stAng, cmd.swAng);
          for (const s of segs) {
            d.push(`C${pt(s.c1x, s.c1y)} ${pt(s.c2x, s.c2y)} ${pt(s.ex, s.ey)}`);
          }
          const last = segs[segs.length - 1];
          if (last) {
            curX = last.ex;
            curY = last.ey;
          }
          break;
        }
        case 'close':
          d.push('Z');
          curX = startX;
          curY = startY;
          break;
      }
    }
    if (d.length === 0) continue;
    const pathFill = path.fill === 'none' ? 'none' : fill;
    const fillAttrs = path.fill === 'none' ? '' : fillExtra;
    const strokeAttrs = path.stroke
      ? ` stroke="${stroke}" stroke-width="${E(strokeWidthEmu)}"${strokeExtra}${markerExtra}`
      : ' stroke="none"';
    out.push(
      `<path d="${d.join(' ')}" fill="${pathFill}"${fillAttrs}${strokeAttrs} fill-rule="evenodd"/>`,
    );
  }
  return out.join('');
};

const renderShapeContent = (
  shape: SlideShapeData,
  pres: PresentationData,
  theme: PresentationTheme | null,
  ctx: LayoutCtx,
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
  const kind = getShapeKind(shape);
  // Connectors (lines) are legitimately zero-height when horizontal or
  // zero-width when vertical — only a true point (both zero) is degenerate.
  // Every other shape with no area is invisible, so skip it.
  if (kind === 'connector' ? w <= 0 && h <= 0 : w <= 0 || h <= 0) return '';

  const fill = getShapeFillEffective(pres, shape);
  const stroke = getShapeStrokeEffective(pres, shape);
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
  // Text follows the shape's ROTATION but not its flips — PowerPoint mirrors a
  // shape's geometry on flip while keeping its text upright and readable.
  // flip.vertical alone (or combined with flip.horizontal) still needs an
  // extra half-turn to keep the label right-reading-up, though: flipping a
  // shape top-to-bottom turns its (unflipped) text upside down, so PowerPoint
  // adds 180° to the text's rotation specifically for a vertical flip.
  // flip.horizontal never needs this — mirroring left-right keeps reading
  // direction (and text-upright-ness) unchanged. Verified against real
  // LibreOffice output across rotation × {none, flipH, flipV, both}.
  const textRotation = (flip.vertical ? rotation + 180 : rotation) % 360;
  const textTransforms: string[] = [];
  if (textRotation !== 0) textTransforms.push(`rotate(${textRotation} ${E(cx)} ${E(cy)})`);
  // Cancel ancestor reflection in the text's local axes, before its rotation.
  // This preserves the transformed center and baseline direction, including
  // the half-turn introduced by a vertical group flip. Two reflections cancel
  // naturally, even when intervening groups have their own rotations.
  if (ctx.groupReflected) textTransforms.push(`translate(${E(2 * cx)} 0) scale(-1 1)`);
  const textTransform = textTransforms.length > 0 ? ` transform="${textTransforms.join(' ')}"` : '';

  // A group's scale moves and resizes its children but leaves their text at the
  // authored point size — resizing a group in PowerPoint never reflows the type,
  // and LibreOffice renders it the same way. Laying the text out inside the
  // group-scaled rect and cancelling the scale back out keeps the box where the
  // geometry is while the glyphs stay undistorted; a non-uniform group scale
  // would otherwise squash them (`<a:ext cy>` far under `<a:chExt cy>` is what
  // Google Slides writes for a hand-resized group).
  const { sx: gsx, sy: gsy } = ctx.groupScale;
  const groupScaled = gsx !== 1 || gsy !== 1;
  const textBounds = groupScaled
    ? { x: x * gsx, y: y * gsy, w: w * gsx, h: h * gsy }
    : { x, y, w, h };
  const rawTextOverlay =
    kind === 'shape' || kind === 'graphicFrame'
      ? renderTextBody(pres, shape, textBounds, theme, phType, ctx)
      : '';
  const textOverlay =
    rawTextOverlay !== '' && groupScaled
      ? `<g transform="scale(${(1 / gsx).toFixed(6)} ${(1 / gsy).toFixed(6)})">${rawTextOverlay}</g>`
      : rawTextOverlay;

  if (kind === 'picture') {
    return renderPicture(
      shape,
      pres,
      x,
      y,
      w,
      h,
      transform,
      textOverlay,
      getShapeImageBytes(shape),
      getShapeImageFormat(shape),
      paint(shape, fill, stroke, theme, false, pres),
    );
  }

  // Shapes with an image fill (`<p:sp>` + `<a:blipFill>` instead of a
  // solid / gradient / pattern). PowerPoint's "Insert Picture from
  // File" and several third-party tools emit pictures this way rather
  // than as top-level `<p:pic>`.
  if (kind === 'shape' && fill.kind === 'image') {
    return renderPicture(
      shape,
      pres,
      x,
      y,
      w,
      h,
      transform,
      textOverlay,
      getShapeImageFillBytes(shape),
      getShapeImageFormat(shape),
      paint(shape, fill, stroke, theme, false, pres),
    );
  }

  if (kind === 'connector') {
    const p = paint(shape, fill, stroke, theme, false, pres);
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
    // Connectors encode their flip as which endpoint is x1/y1 vs x2/y2
    // (above), not as a transform — `transform` (built earlier) still
    // carries the flip's translate+scale for shapes/pictures, which would
    // double-apply the flip here and reverse the line back to unflipped.
    // Rotation, unlike flip, isn't reflected in x1/y1/x2/y2, so it's the only
    // part of `transform` a connector still needs.
    const connectorTransform =
      rotation !== 0 ? ` transform="rotate(${rotation} ${E(cx)} ${E(cy)})"` : '';
    const strokeColor =
      p.stroke === 'none' ? resolveColor('scheme:tx1', theme, '#1F2937') : p.stroke;
    const sa = p.strokeAttrs ? ` ${p.strokeAttrs}` : '';
    const ma = p.markerAttrs ?? '';
    // Connectors default to a round cap/join, but an explicit cap/join from
    // <a:ln cap="…"> / <a:round/> already sits in strokeAttrs — emitting the
    // default alongside it would repeat the attribute and break SVG parsing.
    const capDefault = sa.includes('stroke-linecap') ? '' : ' stroke-linecap="round"';
    const joinDefault = sa.includes('stroke-linejoin') ? '' : ' stroke-linejoin="round"';
    // B8 — bent / curved connector routing. Per ECMA-376 §20.1.9.18,
    // bentConnector{2,3,4,5} are L-shaped, step, and double-step paths;
    // curvedConnector{2,3,4,5} are quadratic / cubic Bézier curves. We
    // route them between the bounding box's diagonal endpoints in
    // CSS-px so the cadence matches PowerPoint within visual tolerance.
    const preset = getShapePreset(shape) ?? 'line';
    if (preset === 'straightConnector1' || preset === 'line') {
      return `${p.defs}<line x1="${E(x1)}" y1="${E(y1)}" x2="${E(x2)}" y2="${E(y2)}" stroke="${strokeColor}" stroke-width="${E(sw)}"${capDefault}${sa}${ma}${connectorTransform}/>`;
    }
    // For bent / curved, we work in CSS px to keep the path math readable.
    const px1 = x1 / EMU_PER_PX;
    const py1 = y1 / EMU_PER_PX;
    const px2 = x2 / EMU_PER_PX;
    const py2 = y2 / EMU_PER_PX;
    let d = `M${px1.toFixed(2)} ${py1.toFixed(2)}`;
    if (preset === 'bentConnector2') {
      // Single L: horizontal then vertical.
      d += ` L${px2.toFixed(2)} ${py1.toFixed(2)} L${px2.toFixed(2)} ${py2.toFixed(2)}`;
    } else if (preset === 'bentConnector3') {
      // Z-bend: horizontal halfway, vertical, horizontal to endpoint.
      const midX = (px1 + px2) / 2;
      d += ` L${midX.toFixed(2)} ${py1.toFixed(2)} L${midX.toFixed(2)} ${py2.toFixed(2)} L${px2.toFixed(2)} ${py2.toFixed(2)}`;
    } else if (preset === 'bentConnector4') {
      // Two-step: H halfway, V halfway, H, V to endpoint.
      const mx = (px1 + px2) / 2;
      const my = (py1 + py2) / 2;
      d += ` L${mx.toFixed(2)} ${py1.toFixed(2)} L${mx.toFixed(2)} ${my.toFixed(2)} L${px2.toFixed(2)} ${my.toFixed(2)} L${px2.toFixed(2)} ${py2.toFixed(2)}`;
    } else if (preset === 'bentConnector5') {
      // Three-step: same pattern, one more segment.
      const q1x = px1 + (px2 - px1) / 3;
      const q2x = px1 + (2 * (px2 - px1)) / 3;
      const my = (py1 + py2) / 2;
      d += ` L${q1x.toFixed(2)} ${py1.toFixed(2)} L${q1x.toFixed(2)} ${my.toFixed(2)} L${q2x.toFixed(2)} ${my.toFixed(2)} L${q2x.toFixed(2)} ${py2.toFixed(2)} L${px2.toFixed(2)} ${py2.toFixed(2)}`;
    } else if (preset === 'curvedConnector2') {
      // Single quadratic curve through the corner.
      d += ` Q${px2.toFixed(2)} ${py1.toFixed(2)} ${px2.toFixed(2)} ${py2.toFixed(2)}`;
    } else if (
      preset === 'curvedConnector3' ||
      preset === 'curvedConnector4' ||
      preset === 'curvedConnector5'
    ) {
      // S-curve: cubic Bézier with control points at one-third / two-thirds.
      const c1x = px1 + (px2 - px1) / 3;
      const c2x = px1 + (2 * (px2 - px1)) / 3;
      d += ` C${c1x.toFixed(2)} ${py1.toFixed(2)} ${c2x.toFixed(2)} ${py2.toFixed(2)} ${px2.toFixed(2)} ${py2.toFixed(2)}`;
    } else {
      // Unknown connector preset — fall back to a straight line.
      d += ` L${px2.toFixed(2)} ${py2.toFixed(2)}`;
    }
    return `${p.defs}<path d="${d}" fill="none" stroke="${strokeColor}" stroke-width="${E(sw)}"${capDefault}${joinDefault}${sa}${ma}${connectorTransform}/>`;
  }

  if (kind === 'group') {
    // Recurse into the group's children. Their bounds live in the
    // group's internal coordinate system; an SVG transform maps that
    // onto the slide. Children are rendered the same way as
    // top-level shapes — nested groups recurse naturally.
    const xform = getGroupTransform(shape);
    const children = getGroupChildren(shape);
    if (children.length === 0) return '';
    const tParts: string[] = [];
    const inverseParts: string[] = [];
    let groupScaleX = 1;
    let groupScaleY = 1;
    // B7 — group-level rotation / flip. The group's <a:xfrm rot=…
    // flipH=… flipV=…> applies to the whole subtree, around the group's
    // outer-rect center. Compose those transforms first, then the
    // translate+scale that maps internal coords onto slide coords.
    if (xform && rotation !== 0) {
      const cxG = ((xform.outer.x as number) + (xform.outer.w as number) / 2) / EMU_PER_PX;
      const cyG = ((xform.outer.y as number) + (xform.outer.h as number) / 2) / EMU_PER_PX;
      tParts.push(`rotate(${rotation} ${cxG.toFixed(2)} ${cyG.toFixed(2)})`);
      inverseParts.unshift(`rotate(${-rotation} ${cxG.toFixed(2)} ${cyG.toFixed(2)})`);
    }
    if (xform && flip.horizontal) {
      const cxG = ((xform.outer.x as number) + (xform.outer.w as number) / 2) / EMU_PER_PX;
      tParts.push(`translate(${(2 * cxG).toFixed(2)} 0) scale(-1 1)`);
      inverseParts.unshift(`translate(${(2 * cxG).toFixed(2)} 0) scale(-1 1)`);
    }
    if (xform && flip.vertical) {
      const cyG = ((xform.outer.y as number) + (xform.outer.h as number) / 2) / EMU_PER_PX;
      tParts.push(`translate(0 ${(2 * cyG).toFixed(2)}) scale(1 -1)`);
      inverseParts.unshift(`translate(0 ${(2 * cyG).toFixed(2)}) scale(1 -1)`);
    }
    if (xform) {
      const ox = xform.outer.x as number;
      const oy = xform.outer.y as number;
      const ow = xform.outer.w as number;
      const oh = xform.outer.h as number;
      const ix = xform.inner.x as number;
      const iy = xform.inner.y as number;
      const iw = (xform.inner.w as number) || 1;
      const ih = (xform.inner.h as number) || 1;
      groupScaleX = ow / iw;
      groupScaleY = oh / ih;
      const sx = groupScaleX.toFixed(6);
      const sy = groupScaleY.toFixed(6);
      // translate first, then scale, so the child's natural coords
      // (px) project: ox/EMU_PER_PX + (cx - ix) * sx / EMU_PER_PX,
      // which factors as translate(ox/EMU - ix*sx/EMU) scale(sx).
      const tx = ((ox - ix * (ow / iw)) / EMU_PER_PX).toFixed(2);
      const ty = ((oy - iy * (oh / ih)) / EMU_PER_PX).toFixed(2);
      tParts.push(`translate(${tx} ${ty})`, `scale(${sx} ${sy})`);
      inverseParts.unshift(
        `scale(${1 / Number(sx)} ${1 / Number(sy)}) translate(${-Number(tx)} ${-Number(ty)})`,
      );
    }
    const groupTransform = tParts.length > 0 ? ` transform="${tParts.join(' ')}"` : '';
    const childCtx: LayoutCtx =
      tParts.length === 0
        ? ctx
        : {
            ...ctx,
            inverseGroupTransform: `${inverseParts.join(' ')} ${ctx.inverseGroupTransform}`,
            groupReflected:
              ctx.groupReflected !== Boolean(xform && flip.horizontal !== flip.vertical),
            groupScale: {
              sx: ctx.groupScale.sx * groupScaleX,
              sy: ctx.groupScale.sy * groupScaleY,
            },
          };
    const childrenSvg = children.map((c) => renderShape(c, pres, theme, childCtx)).join('');
    return `<g${groupTransform}>${childrenSvg}</g>`;
  }

  const p = paint(shape, fill, stroke, theme, phType !== null, pres);
  if (fill.kind === 'background') {
    ctx.background.id ||= mintId();
    const id = mintId();
    // Background paint stays in slide coordinates even as the shape and its
    // ancestor groups rotate, reflect or scale around it.
    const inverse: string[] = [];
    if (flip.vertical) inverse.push(`translate(0 ${E(2 * cy)}) scale(1 -1)`);
    if (flip.horizontal) inverse.push(`translate(${E(2 * cx)} 0) scale(-1 1)`);
    if (rotation !== 0) inverse.push(`rotate(${-rotation} ${E(cx)} ${E(cy)})`);
    inverse.push(ctx.inverseGroupTransform);
    p.fill = `url(#${id})`;
    p.fillAttrs = '';
    p.defs += `<defs><pattern id="${id}" patternUnits="userSpaceOnUse" width="${ctx.background.width}" height="${ctx.background.height}" patternTransform="${inverse.join(' ')}"><use href="#${ctx.background.id}" xlink:href="#${ctx.background.id}"/></pattern></defs>`;
  }

  if (kind === 'graphicFrame') {
    // Charts and tables get real renders. SmartArt and the
    // graphicFrame variants @office-kit/pptx doesn't model fall through to a
    // labelled placeholder.
    if (isChartShape(shape)) {
      const chartSvg = renderChart(shape, x, y, w, h, transform, theme, ctx.groupReflected);
      if (chartSvg) return chartSvg;
    }
    if (isTableShape(shape)) {
      const tableSvg = renderTable(shape, pres, x, y, w, h, transform, theme, ctx);
      if (tableSvg) return tableSvg;
    }
    let label = 'graphicFrame';
    try {
      if (isChartShape(shape)) {
        // `isChartShape` was true but renderChart returned null —
        // @office-kit/pptx couldn't model this chart kind. scatter / radar /
        // bubble now render; what's left here is stock / surface / 3D /
        // of-pie, which the reader still folds or drops.
        label = 'chart (unsupported kind)';
      } else if (isTableShape(shape)) {
        const t = getTableDimensions(shape);
        label = `table (${t.rows}×${t.cols})`;
      }
    } catch {
      // Fall through with the generic label.
    }
    const gfFallbackKind = label.startsWith('chart') ? 'chart' : 'graphicFrame';
    return `<g data-pptx-fallback="${gfFallbackKind}"${transform}><rect x="${E(x)}" y="${E(y)}" width="${E(w)}" height="${E(h)}" fill="${p.fill === 'none' ? '#F9FAFB' : p.fill}" stroke="#9CA3AF" stroke-width="${E(9_525)}" stroke-dasharray="${E(50_000)},${E(30_000)}"/>${renderPicturePlaceholderLabel(x, y, w, h, label)}${textOverlay}</g>`;
  }

  // kind === 'shape'
  const rawPreset = getShapePreset(shape);
  const preset = rawPreset ?? 'rect';

  const fa = p.fillAttrs;
  const sa = p.strokeAttrs ? ` ${p.strokeAttrs}` : '';
  const ma = p.markerAttrs ?? '';
  // Custom geometry (<a:custGeom>) overrides the preset path entirely.
  // getShapePreset returns null for it, so only the no-preset case probes.
  const customGeom = rawPreset === null ? getShapeCustomGeometry(shape) : null;
  let geomSvg =
    customGeom !== null
      ? customGeometryToSvg(customGeom, x, y, w, h, p.fill, fa, p.stroke, p.strokeWidth, sa, ma)
      : '';
  // No preset and no rendered custom geometry means either a custGeom that
  // failed to evaluate (a true fallback — marked) or no geometry at all
  // (placeholders inherit theirs from the layout, which is almost always a
  // rect — correct, not marked). The string probe distinguishes the two
  // because getShapeCustomGeometry returns null for both malformed custGeom
  // and absent custGeom.
  const isCustGeom =
    geomSvg === '' && rawPreset === null && getShapeXmlString(shape).includes('custGeom');

  if (geomSvg !== '') {
    // geomSvg already holds the rendered custom geometry.
  } else if (preset === 'rect') {
    geomSvg = `<rect x="${E(x)}" y="${E(y)}" width="${E(w)}" height="${E(h)}" fill="${p.fill}"${fa} stroke="${p.stroke}" stroke-width="${E(p.strokeWidth)}"${sa}${ma}/>`;
  } else if (preset === 'roundRect') {
    // A6 — adjust-handle aware corner radius. <a:gd name="adj"
    // fmla="val N"/> in [0, 50000] = ratio of corner-radius to
    // min(w,h)/2 × 100. Defaults to ~16.6% when no adj is authored.
    const adjusts = getShapeAdjustValues(shape);
    const adjVal = adjusts.adj ?? 16667;
    const ratio = Math.max(0, Math.min(0.5, adjVal / 100_000));
    const r = E(Math.min(w, h) * ratio);
    geomSvg = `<rect x="${E(x)}" y="${E(y)}" width="${E(w)}" height="${E(h)}" rx="${r}" ry="${r}" fill="${p.fill}"${fa} stroke="${p.stroke}" stroke-width="${E(p.strokeWidth)}"${sa}${ma}/>`;
  } else if (preset === 'ellipse' || preset === 'oval') {
    geomSvg = `<ellipse cx="${E(cx)}" cy="${E(cy)}" rx="${E(w / 2)}" ry="${E(h / 2)}" fill="${p.fill}"${fa} stroke="${p.stroke}" stroke-width="${E(p.strokeWidth)}"${sa}${ma}/>`;
  } else {
    const pathFn = PRESET_PATHS[preset];
    if (pathFn) {
      // The path generators output CSS-px coords directly (post-E).
      const d = pathFn(x / EMU_PER_PX, y / EMU_PER_PX, w / EMU_PER_PX, h / EMU_PER_PX);
      geomSvg = `<path d="${d}" fill="${p.fill}"${fa} stroke="${p.stroke}" stroke-width="${E(p.strokeWidth)}" fill-rule="evenodd"${sa}${ma}/>`;
    } else {
      const pointsFn = PRESET_POINTS[preset];
      if (pointsFn) {
        const points = pointsFn(w / EMU_PER_PX, h / EMU_PER_PX)
          .map(([nx, ny]) => `${E(x + nx * w)},${E(y + ny * h)}`)
          .join(' ');
        geomSvg = `<polygon points="${points}" fill="${p.fill}"${fa} stroke="${p.stroke}" stroke-width="${E(p.strokeWidth)}"${sa}${ma}/>`;
      } else {
        // Unrecognised preset — fall back to a rectangle, but tag it
        // with the preset name so users (and future-us) can see which
        // shape needs a renderer. The `<title>` shows on hover; the
        // `data-pptx-preset` attribute is for DevTools inspection.
        geomSvg = `<rect x="${E(x)}" y="${E(y)}" width="${E(w)}" height="${E(h)}" fill="${p.fill}"${fa} stroke="${p.stroke}" stroke-width="${E(p.strokeWidth)}"${sa}${ma} data-pptx-preset="${escapeXml(preset)}"><title>${escapeXml(`preset: ${preset}`)}</title></rect>`;
      }
    }
  }

  // Effects (`<a:effectLst>`): outerShdw / innerShdw / glow / softEdge
  // / reflection / blur. Build a single SVG <filter> chain so multiple
  // effects compose the way PowerPoint composes them.
  const fx = buildEffectsFilter(pres, shape);
  const filterAttr = fx ? ` filter="url(#${fx.id})"` : '';
  let fxDefs = fx ? fx.defs : '';
  // Reflection can't live in the `<filter>` chain (SVG has no flip-and-fade
  // primitive), so it's a duplicated, vertically mirrored copy of the raw
  // geometry masked by an opacity gradient — built from the un-filtered
  // geometry markup before the filter wraps it.
  const reflection = buildReflection(pres, shape, geomSvg, { x, y, w, h });
  // Apply the filter to the geometry only — text overlays use foreignObject
  // and react badly to feGaussianBlur (DOM gets rasterized).
  geomSvg = `<g${filterAttr}>${geomSvg}</g>`;
  if (reflection) {
    // Paint the reflection behind the shape; it sits below the bottom edge
    // so they don't overlap, but keeping it first matches PowerPoint's z-order.
    geomSvg = reflection.svg + geomSvg;
    fxDefs += reflection.defs;
  }

  // Expose the shape's authored name as a data attribute so DevTools /
  // Selenium / a11y inspections can identify a shape without having to
  // parse SVG geometry. The PowerPoint alt-title / alt-description feed
  // `aria-label` so screen readers announce decks the same way
  // PowerPoint's Accessibility Inspector reports them.
  const shapeName = (() => {
    try {
      return getShapeName(shape);
    } catch {
      return null;
    }
  })();
  const altTitle = (() => {
    try {
      return getShapeAltTitle(shape);
    } catch {
      return null;
    }
  })();
  const altDesc = (() => {
    try {
      return getShapeDescription(shape);
    } catch {
      return null;
    }
  })();
  const a11yLabel = altTitle ?? altDesc ?? null;
  const nameAttr = shapeName ? ` data-pptx-shape-name="${escapeXml(shapeName)}"` : '';
  const ariaAttr = a11yLabel ? ` role="img" aria-label="${escapeXml(a11yLabel)}"` : '';
  // Geometry carries rotation + flips; text carries rotation only so it stays
  // upright when the shape is flipped (matching PowerPoint).
  const placedText = textOverlay ? `<g${textTransform}>${textOverlay}</g>` : '';
  const custGeomAttr = isCustGeom ? ' data-pptx-fallback="custGeom"' : '';
  const inner = `${p.defs}${fxDefs}<g${nameAttr}${ariaAttr}${custGeomAttr}><g${transform}>${geomSvg}</g>${placedText}</g>`;
  return inner;
};

// All object kinds use the same link wrapper, including pictures, connectors,
// charts and tables. Run links remain scoped to their rendered text.
const renderShape = (
  shape: SlideShapeData,
  pres: PresentationData,
  theme: PresentationTheme | null,
  ctx: LayoutCtx,
): string => {
  if (isShapeHidden(shape)) return '';
  const inner = renderShapeContent(shape, pres, theme, ctx);
  if (!inner) return inner;
  const href = clickActionHref(pres, shape, getShapeClickAction(shape));
  const tooltip = href === undefined ? null : getShapeHyperlinkTooltip(shape);
  const titleEl = tooltip ? `<title>${escapeXml(tooltip)}</title>` : '';
  const targetAttrs =
    href === undefined || href.startsWith('#') ? '' : ' target="_blank" rel="noopener noreferrer"';
  const linked =
    href === undefined
      ? inner
      : `<a href="${escapeXml(href)}"${targetAttrs}>${titleEl}${inner}</a>`;
  return wrapWithShapeId(shape, linked, ctx);
};

/**
 * Names the object the slide's animation timing would target, so a player can
 * find it without knowing how the shape was drawn. One wrapper per object,
 * whatever its kind — a group carries its own id and its children theirs, so a
 * build on a group child is reachable inside a group that is animated itself.
 *
 * The id is the shape's `<p:cNvPr id>`, which is what `<p:spTgt spid>` names.
 * A deck may repeat one (nothing in the schema forbids it), and the attribute
 * repeats with it rather than inventing a unique number the timing could not
 * refer to.
 */
const wrapWithShapeId = (shape: SlideShapeData, svg: string, ctx: LayoutCtx): string => {
  if (!ctx.ownShapes) return svg;
  const id = (() => {
    try {
      return getShapeId(shape);
    } catch {
      return null;
    }
  })();
  return id === null ? svg : `<g data-pptx-shape-id="${id}">${svg}</g>`;
};

// ---------------------------------------------------------------------------
// Effects → SVG filter.

interface EffectsResult {
  readonly id: string;
  readonly defs: string;
}

interface ReflectionResult {
  readonly svg: string;
  readonly defs: string;
}

// Builds the reflection: a vertically mirrored copy of the shape's raw
// geometry placed below its bottom edge, faded by a vertical opacity mask.
// PowerPoint encodes the mirror as a negative `sy`; `stA`/`endA` give the
// alpha at the near (contact) and far edges, `dist` the gap below the shape.
const buildReflection = (
  pres: PresentationData,
  shape: SlideShapeData,
  geomRaw: string,
  box: { x: number; y: number; w: number; h: number },
): ReflectionResult | null => {
  let effects: readonly ReturnType<typeof getShapeEffects>[number][];
  try {
    effects = getShapeEffectsEffective(pres, shape);
  } catch {
    return null;
  }
  const refl = effects.find((e) => e.kind === 'reflection');
  if (!refl || refl.kind !== 'reflection') return null;

  // Honor the authored vertical scale (signed: negative = the flip);
  // default to a full-height mirror when the deck omits `sy`, since a
  // reflection with no flip isn't a reflection.
  const f = refl.scaleY ?? -1;
  const startA = refl.startOpacity ?? 1;
  const endA = refl.opacity ?? 0;
  // Geometry coords are emitted in px (see `E`), so the transform math
  // works in px too: contact edge at the shape's bottom, gap pushed down.
  const contactPx = (box.y + box.h) / EMU_PER_PX;
  const distPx = refl.distEmu / EMU_PER_PX;
  // scale(1,f) about y=contactPx, then translate the whole thing down by dist.
  const transform = `translate(0 ${(distPx + contactPx * (1 - f)).toFixed(2)}) scale(1 ${f})`;

  const maskId = mintId();
  const gradId = mintId();
  // objectBoundingBox: y=0 is the contact edge (top of the mirrored copy),
  // y=1 the far edge. White luminance × stop-opacity becomes the alpha.
  const defs =
    `<defs><linearGradient id="${gradId}" x1="0" y1="0" x2="0" y2="1">` +
    `<stop offset="0" stop-color="#fff" stop-opacity="${startA.toFixed(3)}"/>` +
    `<stop offset="1" stop-color="#fff" stop-opacity="${endA.toFixed(3)}"/>` +
    `</linearGradient>` +
    `<mask id="${maskId}" maskContentUnits="objectBoundingBox">` +
    `<rect width="1" height="1" fill="url(#${gradId})"/></mask></defs>`;
  const svg = `<g transform="${transform}" mask="url(#${maskId})" data-pptx-reflection="1">${geomRaw}</g>`;
  return { svg, defs };
};

const buildEffectsFilter = (
  pres: PresentationData,
  shape: SlideShapeData,
): EffectsResult | null => {
  let effects: readonly ReturnType<typeof getShapeEffects>[number][];
  try {
    effects = getShapeEffectsEffective(pres, shape);
  } catch {
    return null;
  }
  if (effects.length === 0) return null;

  const id = mintId();
  const primitives: string[] = [];
  // Chain primitives by passing each result as `in` to the next merge.
  // The shape's original alpha + RGB live in SourceGraphic / SourceAlpha.
  const layers: string[] = [];

  for (const e of effects) {
    if (e.kind === 'outerShdw') {
      // dist + angle → dx, dy in EMU → px.
      const rad = (e.angleDeg * Math.PI) / 180;
      const dx = (e.distEmu * Math.cos(rad)) / EMU_PER_PX;
      const dy = (e.distEmu * Math.sin(rad)) / EMU_PER_PX;
      const blurPx = e.blurEmu / EMU_PER_PX / 2;
      const opacity = e.opacity ?? 1;
      const color = e.color || '#000000';
      // feDropShadow handles the whole shadow primitive in one go.
      const out = `shdwOut${primitives.length}`;
      primitives.push(
        `<feDropShadow dx="${dx.toFixed(2)}" dy="${dy.toFixed(2)}" stdDeviation="${blurPx.toFixed(2)}" flood-color="${color}" flood-opacity="${opacity.toFixed(3)}" result="${out}"/>`,
      );
      layers.push(out);
    } else if (e.kind === 'innerShdw') {
      // SVG has no innerShadow primitive — synthesize via:
      //   inset = (sourceAlpha offset, blurred) - sourceAlpha (inverted)
      // and re-flood with the shadow color.
      const rad = (e.angleDeg * Math.PI) / 180;
      const dx = (e.distEmu * Math.cos(rad)) / EMU_PER_PX;
      const dy = (e.distEmu * Math.sin(rad)) / EMU_PER_PX;
      const blurPx = e.blurEmu / EMU_PER_PX / 2;
      const color = e.color || '#000000';
      const opacity = e.opacity ?? 1;
      const i = primitives.length;
      primitives.push(
        `<feGaussianBlur in="SourceAlpha" stdDeviation="${blurPx.toFixed(2)}" result="innerBlur${i}"/>`,
        `<feOffset in="innerBlur${i}" dx="${dx.toFixed(2)}" dy="${dy.toFixed(2)}" result="innerOff${i}"/>`,
        `<feComposite in="innerOff${i}" in2="SourceAlpha" operator="arithmetic" k2="-1" k3="1" result="innerMask${i}"/>`,
        `<feFlood flood-color="${color}" flood-opacity="${opacity.toFixed(3)}" result="innerCol${i}"/>`,
        `<feComposite in="innerCol${i}" in2="innerMask${i}" operator="in" result="innerOut${i}"/>`,
      );
      layers.push(`innerOut${i}`);
    } else if (e.kind === 'glow') {
      // PowerPoint / LibreOffice keep the glow color near-opaque for most of
      // the `rad` reach and feather only at the outer edge. Compositing flood
      // 'in' a single wide Gaussian caps peak alpha at ~0.5 and over-diffuses,
      // so decouple the saturated band (a large dilation) from the feather (a
      // small blur): dilate by half the full reach, then blur lightly.
      const radiusPx = e.radiusEmu / EMU_PER_PX;
      const dilatePx = radiusPx * 0.5;
      const featherPx = radiusPx * 0.3;
      const color = e.color || '#FFFFFF';
      const opacity = e.opacity ?? 1;
      const i = primitives.length;
      primitives.push(
        `<feMorphology in="SourceAlpha" operator="dilate" radius="${dilatePx.toFixed(2)}" result="glowExp${i}"/>`,
        `<feGaussianBlur in="glowExp${i}" stdDeviation="${featherPx.toFixed(2)}" result="glowBlur${i}"/>`,
        `<feFlood flood-color="${color}" flood-opacity="${opacity.toFixed(3)}" result="glowCol${i}"/>`,
        `<feComposite in="glowCol${i}" in2="glowBlur${i}" operator="in" result="glowOut${i}"/>`,
      );
      layers.push(`glowOut${i}`);
    } else if (e.kind === 'softEdge') {
      const blurPx = e.radiusEmu / EMU_PER_PX / 2;
      // Soft-edge feathers the shape's mask. Replace the source by a
      // blurred version of itself.
      const i = primitives.length;
      primitives.push(
        `<feGaussianBlur in="SourceGraphic" stdDeviation="${blurPx.toFixed(2)}" result="softOut${i}"/>`,
      );
      // softEdge replaces the source; we drop earlier layers and the
      // unmodified source is no longer painted on top.
      layers.length = 0;
      layers.push(`softOut${i}`);
    } else if (e.kind === 'blur') {
      const blurPx = e.radiusEmu / EMU_PER_PX / 2;
      const i = primitives.length;
      primitives.push(
        `<feGaussianBlur in="SourceGraphic" stdDeviation="${blurPx.toFixed(2)}" result="blurOut${i}"/>`,
      );
      layers.length = 0;
      layers.push(`blurOut${i}`);
    }
    // Reflection is handled outside the filter chain (see buildReflection):
    // SVG `<filter>` has no flip-and-fade primitive.
  }

  // Compose: paint each effect layer plus the original SourceGraphic.
  // Shadows want to sit behind the source; glow behind too; innerShdw
  // and softEdge already replace bits of the source. Doing the merge
  // in order produces reasonable layering for the common cases.
  if (layers.length === 0) return null;

  // Always paint the original source last so it sits on top of shadows /
  // glows. softEdge/blur replaced the source so we don't double-paint.
  const replacedSource = effects.some((e) => e.kind === 'softEdge' || e.kind === 'blur');
  const mergeChildren = layers.map((l) => `<feMergeNode in="${l}"/>`).join('');
  const sourceMerge = replacedSource ? '' : '<feMergeNode in="SourceGraphic"/>';
  primitives.push(`<feMerge>${mergeChildren}${sourceMerge}</feMerge>`);

  const defs = `<defs><filter id="${id}" x="-25%" y="-25%" width="150%" height="150%">${primitives.join('')}</filter></defs>`;
  return { id, defs };
};

// ---------------------------------------------------------------------------
// Slide composition.

// getSlideShapes / getSlideLayoutShapes / getSlideMasterShapes flatten group
// descendants into the returned list (document order: each group is followed
// immediately by its descendants), while renderShape already recurses into
// groups and draws every child with the group transform applied. Rendering
// the flat list verbatim would therefore paint each group child a second
// time, untransformed — visibly offset whenever the group's chOff differs
// from its off. Walk the flat list and skip each group's descendants.
//
// The decoration readers (layout / master) additionally drop placeholder
// shapes from their flat list, so placeholder descendants must not be
// counted when skipping.
const flattenedDescendantCount = (group: SlideShapeData, dropPlaceholders: boolean): number => {
  let count = 0;
  for (const child of getGroupChildren(group)) {
    if (!dropPlaceholders || !isShapePlaceholder(child)) count += 1;
    if (getShapeKind(child) === 'group') {
      count += flattenedDescendantCount(child, dropPlaceholders);
    }
  }
  return count;
};

const topLevelShapes = (
  shapes: ReadonlyArray<SlideShapeData>,
  opts: { dropPlaceholders: boolean },
): SlideShapeData[] => {
  const out: SlideShapeData[] = [];
  let skip = 0;
  for (const shape of shapes) {
    if (skip > 0) {
      skip -= 1;
      continue;
    }
    out.push(shape);
    if (getShapeKind(shape) === 'group') {
      skip = flattenedDescendantCount(shape, opts.dropPlaceholders);
    }
  }
  return out;
};

export const renderSlideSvg = (
  pres: PresentationData,
  slide: SlideData,
  opts: RenderSlideOptions = {},
): string => {
  const size = getSlideSize(pres) ?? DEFAULT_SIZE;
  const W = size.width as number;
  const H = size.height as number;
  const theme = getPresentationTheme(pres);
  activeColorMap = getEffectiveColorMap(slide);
  activeDeckTextColor = resolveDeckBodyTextColor(slide) ?? '#000000';
  const position = getSlides(pres).indexOf(slide);
  activeSlideNumber = String(getPresentationFirstSlideNumber(pres) + (position < 0 ? 0 : position));
  const ctx: LayoutCtx = {
    groupScale: { sx: 1, sy: 1 },
    groupReflected: false,
    mode: opts.textLayout ?? 'foreignObject',
    measure: opts.measureText ?? defaultMeasurer,
    ownShapes: true,
    background: { id: '', width: W / EMU_PER_PX, height: H / EMU_PER_PX },
    inverseGroupTransform: '',
  };

  let bg = getSlideBackground(slide);
  // B10 — when the slide reports inherit, walk to the layout; when the
  // layout also inherits, walk one more step to the master. Real brand
  // templates put the actual background fill on the master only.
  if (bg.kind === 'inherit') {
    const layout = getSlideLayout(slide);
    if (layout) {
      const layoutBg = getSlideLayoutBackground(layout);
      if (layoutBg.kind !== 'inherit') {
        bg = layoutBg;
      } else {
        const masterBg = getSlideMasterBackground(pres, layout);
        if (masterBg.kind !== 'inherit') bg = masterBg;
      }
    }
  }
  let bgColor = '#FFFFFF';
  const bgOpacity = bg.kind === 'solid' ? bg.opacity : undefined;
  const bgOpacityAttr = bgOpacity === undefined ? '' : ` fill-opacity="${bgOpacity}"`;
  let bgGradient = '';
  let bgGradientDefs = '';
  if (bg.kind === 'solid') {
    bgColor = resolveColor(bg.color, theme, '#FFFFFF');
  } else if (bg.kind === 'gradient') {
    // B11 — gradient slide backgrounds. Use the same projector as
    // shape fills so radial / rect / shape paths all behave.
    // Walk slide → layout → master for the actual gradient definition.
    let grad = getSlideBackgroundGradientFill(slide);
    if (!grad) {
      const layout = getSlideLayout(slide);
      if (layout) {
        grad = getSlideLayoutBackgroundGradientFill(layout);
        if (!grad) grad = getSlideMasterBackgroundGradientFill(pres, layout);
      }
    }
    if (grad) {
      const built = gradientDef(grad, theme);
      bgGradientDefs = built.defs;
      bgGradient = `<rect width="${E(W)}" height="${E(H)}" fill="${built.fillAttr}"/>`;
    }
  } else if (bg.kind === 'pattern') {
    let pat = getSlideBackgroundPatternFill(pres, slide);
    if (!pat) {
      const layout = getSlideLayout(slide);
      if (layout) {
        pat = getSlideLayoutBackgroundPatternFill(pres, layout);
        if (!pat) pat = getSlideMasterBackgroundPatternFill(pres, layout);
      }
    }
    if (pat) {
      const built = patternDef(pat);
      bgGradientDefs += built.defs;
      bgGradient = `<rect width="${E(W)}" height="${E(H)}" fill="${built.fillAttr}"/>`;
    }
  } else if (theme && bg.kind === 'inherit') {
    bgColor = normalizeHex(theme.light1);
  }

  // Slide background image (`<p:bgPr><a:blipFill>`). Paint it above the
  // solid bg-color rect so the bytes show through; shapes still draw
  // on top of the image.
  let bgImage = '';
  if (bg.kind === 'image') {
    let bytes = getSlideBackgroundImageBytes(slide);
    if (!bytes && pres) {
      const layout = getSlideLayout(slide);
      if (layout) {
        bytes = getSlideLayoutBackgroundImageBytes(pres, layout);
        if (!bytes) bytes = getSlideMasterBackgroundImageBytes(pres, layout);
      }
    }
    if (bytes) {
      const fmt = detectImageFormatLocal(bytes);
      const mime = fmt ? (imageMime[fmt] ?? 'image/png') : 'image/png';
      const dataUrl = `data:${mime};base64,${u8ToBase64(bytes)}`;
      bgImage = `<image x="0" y="0" width="${E(W)}" height="${E(H)}" href="${dataUrl}" xlink:href="${dataUrl}" preserveAspectRatio="xMidYMid slice"/>`;
    }
  }

  // Layout-level non-placeholder shapes — corner bars, divider lines,
  // template logos as solid rects. Painted *before* slide shapes so
  // slide content stays on top.
  // Template decoration: the non-placeholder shapes on the slide master and
  // its layout (corner bars, divider lines, logos, watermark text). Render them
  // through the full `renderShape` pipeline — master behind layout, both behind
  // the slide's own content — so logos (pictures), text and real geometry all
  // appear. Picture bytes resolve because the shapes are bound to their part.
  let layoutBgShapes = '';
  const layoutForBg = getSlideLayout(slide);
  if (layoutForBg && !isSlideBackgroundGraphicsHidden(slide)) {
    try {
      const masterShapes = topLevelShapes(
        isSlideLayoutBackgroundGraphicsHidden(layoutForBg)
          ? []
          : getSlideMasterShapes(pres, layoutForBg),
        {
          dropPlaceholders: true,
        },
      );
      const layoutShapes = topLevelShapes(getSlideLayoutShapes(pres, layoutForBg), {
        dropPlaceholders: true,
      });
      // Drawn as background: these come from the layout and the master, whose
      // shape ids mean nothing to this slide's timing.
      const bgCtx: LayoutCtx = { ...ctx, ownShapes: false };
      layoutBgShapes = [...masterShapes, ...layoutShapes]
        .map((s) => renderShape(s, pres, theme, bgCtx))
        .join('');
    } catch {
      layoutBgShapes = '';
    }
  }

  const shapesSvg = topLevelShapes(getSlideShapes(slide), { dropPlaceholders: false })
    .map((s) => renderShape(s, pres, theme, ctx))
    .join('');

  const backgroundSvg = `${bgOpacity === undefined ? '' : `<rect width="${E(W)}" height="${E(H)}" fill="#FFFFFF"/>`}<rect width="${E(W)}" height="${E(H)}" fill="${bgColor}"${bgOpacityAttr}/>${bgGradient}${bgImage}`;

  return [
    `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 ${E(W)} ${E(H)}" preserveAspectRatio="xMidYMid meet">`,
    bgGradientDefs,
    ctx.background.id
      ? `<defs><g id="${ctx.background.id}">${backgroundSvg}</g></defs><use href="#${ctx.background.id}" xlink:href="#${ctx.background.id}"/>`
      : backgroundSvg,
    layoutBgShapes,
    shapesSvg,
    '</svg>',
  ].join('');
};

// Minimal magic-byte sniffer covering the formats we already know how
// to MIME-type. Used for slide background images, which @office-kit/pptx
// returns as raw bytes without exposing the format.
const detectImageFormatLocal = (bytes: Uint8Array): string | null => {
  if (bytes.length < 4) return null;
  if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47)
    return 'png';
  if (bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return 'jpeg';
  if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) return 'gif';
  if (bytes[0] === 0x42 && bytes[1] === 0x4d) return 'bmp';
  if (
    bytes.length >= 12 &&
    bytes[0] === 0x52 &&
    bytes[1] === 0x49 &&
    bytes[2] === 0x46 &&
    bytes[3] === 0x46 &&
    bytes[8] === 0x57 &&
    bytes[9] === 0x45 &&
    bytes[10] === 0x42 &&
    bytes[11] === 0x50
  )
    return 'webp';
  return null;
};

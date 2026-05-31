// rPr and pPr cascade resolution.

import {
  ATTR_LVL,
  NAME_A_PPR,
  NAME_A_RPR,
  ensureRPr,
  requireParagraph,
  requireRun,
  runsOf,
} from './shape-runs.ts';
import { parseRPrLikeElement } from './shape-color.ts';
import {
  getShapePlaceholderIdx,
  getShapePlaceholderType,
  matchPlaceholderShape,
} from './shape-read-base.ts';
import { getSlideLayout } from './shape-slide-read.ts';
import {
  type ParagraphAlignment,
  type TextFormat,
  applyHyperlinkToAllRuns,
  applyRunFormat as applyRunFormatInternal,
} from '../../internal/drawingml/index.ts';
import { emptyRels, nextRelId, partName, resolveTarget } from '../../internal/opc/index.ts';
import { REL_TYPES, readShapeTreeFromCsldRoot } from '../../internal/presentationml/index.ts';
import {
  NS,
  type XmlElement,
  firstChildElement,
  getAttrValue,
  parseXml,
  qname,
} from '../../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  LAYOUT_PART,
  LAYOUT_PART_NAME,
  type PresentationData,
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  SHAPE_SNAPSHOT,
  SLIDE_PART_NAME,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import { commitAndRefresh, decode, requireTxBody } from './_helpers.ts';
import { getPresentationFonts, getPresentationTheme } from './theme.ts';
// -- Effective rPr cascade (ECMA-376 §21.1.2.4.7) ---------------------------
//
// A run's effective character properties are resolved by walking the
// inheritance chain — each level fills in fields that no earlier level
// supplied. First-wins per property:
//
//   1. The run's own `<a:rPr>`
//   2. The paragraph's `<a:endParaRPr>` (last run only)
//   3. The paragraph's `<a:pPr><a:defRPr>` (paragraph-level run defaults)
//   4. The text body's `<a:lstStyle><a:lvl{N+1}pPr><a:defRPr>` (N = paragraph level)
//   5. The same path on the matching placeholder in the slide's layout
//   6. The same path on the matching placeholder on the slide master,
//      then the master's `<p:txStyles>` (`titleStyle` / `bodyStyle` / `otherStyle`)
//   7. The theme's `<a:fontScheme>` — font typeface fallback only
//
// Placeholder matching: by `<p:ph/@idx>` first, then by `<p:ph/@type>`.

const NAME_A_DEF_RPR = qname('a', 'defRPr', NS.dml);
const NAME_A_END_PARA_RPR = qname('a', 'endParaRPr', NS.dml);
const NAME_A_LST_STYLE = qname('a', 'lstStyle', NS.dml);
const NAME_P_TX_BODY_PML = qname('p', 'txBody', NS.pml);
const NAME_P_TX_STYLES = qname('p', 'txStyles', NS.pml);
const NAME_P_TITLE_STYLE = qname('p', 'titleStyle', NS.pml);
const NAME_P_BODY_STYLE = qname('p', 'bodyStyle', NS.pml);
const NAME_P_OTHER_STYLE = qname('p', 'otherStyle', NS.pml);

const mergeRPrLayer = (base: Partial<TextFormat>, layer: Partial<TextFormat>): void => {
  if (base.font === undefined && layer.font !== undefined) base.font = layer.font;
  if (base.size === undefined && layer.size !== undefined) base.size = layer.size;
  if (base.color === undefined && layer.color !== undefined) base.color = layer.color;
  if (base.bold === undefined && layer.bold !== undefined) base.bold = layer.bold;
  if (base.italic === undefined && layer.italic !== undefined) base.italic = layer.italic;
  if (base.underline === undefined && layer.underline !== undefined) {
    base.underline = layer.underline;
  }
  if (base.strike === undefined && layer.strike !== undefined) base.strike = layer.strike;
  if (base.spc === undefined && layer.spc !== undefined) base.spc = layer.spc;
  if (base.kern === undefined && layer.kern !== undefined) base.kern = layer.kern;
  if (base.baseline === undefined && layer.baseline !== undefined) base.baseline = layer.baseline;
  if (base.cap === undefined && layer.cap !== undefined) base.cap = layer.cap;
  if (base.highlight === undefined && layer.highlight !== undefined) {
    base.highlight = layer.highlight;
  }
};

// `<a:lstStyle>` carries one `<a:lvl{N}pPr>` per outline level (1..9, plus
// `<a:defPPr>` for the level-0 default). Returns the inner `<a:defRPr>` for
// the requested zero-based level, or `null` if the level isn't authored.
const lstStyleLevelDefRPr = (lstStyle: XmlElement | null, level: number): XmlElement | null => {
  if (!lstStyle) return null;
  const localName = `lvl${Math.max(0, Math.min(8, level)) + 1}pPr`;
  const lvlPPr = firstChildElement(lstStyle, qname('a', localName, NS.dml));
  if (!lvlPPr) {
    // Fall back to `<a:defPPr>` only for level 0 — that's what the schema
    // declares as the "no explicit level" slot.
    if (level !== 0) return null;
    const defPPr = firstChildElement(lstStyle, qname('a', 'defPPr', NS.dml));
    if (!defPPr) return null;
    return firstChildElement(defPPr, NAME_A_DEF_RPR);
  }
  return firstChildElement(lvlPPr, NAME_A_DEF_RPR);
};

// Companion to `lstStyleLevelDefRPr` but returns the `<a:lvlNpPr>` (or
// `<a:defPPr>` for level 0) element itself — i.e. the paragraph-property
// container, not the run-default child. Used by the pPr cascade.
const lstStyleLevelPPr = (lstStyle: XmlElement | null, level: number): XmlElement | null => {
  if (!lstStyle) return null;
  const localName = `lvl${Math.max(0, Math.min(8, level)) + 1}pPr`;
  const lvlPPr = firstChildElement(lstStyle, qname('a', localName, NS.dml));
  if (lvlPPr) return lvlPPr;
  if (level !== 0) return null;
  return firstChildElement(lstStyle, qname('a', 'defPPr', NS.dml));
};

const findShapeLstStyleElement = (shape: SlideShapeData): XmlElement | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_P_TX_BODY_PML);
  if (!txBody) return null;
  return firstChildElement(txBody, NAME_A_LST_STYLE);
};

const findPlaceholderShapeIn = (
  shapes: ReadonlyArray<{
    placeholderIdx: number | null;
    placeholderType: string | null;
    element: XmlElement;
  }>,
  phIdx: number | null,
  phType: string | null,
): { element: XmlElement } | undefined => matchPlaceholderShape(shapes, phIdx, phType);

const extractPlaceholderLstStyle = (placeholderEl: XmlElement): XmlElement | null => {
  const txBody = firstChildElement(placeholderEl, NAME_P_TX_BODY_PML);
  if (!txBody) return null;
  return firstChildElement(txBody, NAME_A_LST_STYLE);
};

const NAME_P_NV_SP_PR = qname('p', 'nvSpPr', NS.pml);
const NAME_P_NV_PR = qname('p', 'nvPr', NS.pml);
const NAME_P_PH = qname('p', 'ph', NS.pml);

// A shape participates in placeholder inheritance only when it actually carries
// a `<p:ph>`. A plain text box (no `<p:ph>`) must NOT inherit the slide master's
// titleStyle / bodyStyle / otherStyle — PowerPoint resolves it against the
// presentation's default text style (≈18pt), not the master body style.
// `getShapePlaceholderType()` returns null both for "placeholder with no type"
// (which DOES default to body) and "not a placeholder at all" (which does not),
// so we must check for the element directly to tell them apart.
const shapeIsPlaceholder = (shape: SlideShapeData): boolean => {
  const nvSpPr = firstChildElement(shape[SHAPE_ELEMENT], NAME_P_NV_SP_PR);
  if (!nvSpPr) return false;
  const nvPr = firstChildElement(nvSpPr, NAME_P_NV_PR);
  return nvPr !== null && firstChildElement(nvPr, NAME_P_PH) !== null;
};

const masterTxStyleFor = (masterRoot: XmlElement, phType: string | null): XmlElement | null => {
  const txStyles = firstChildElement(masterRoot, NAME_P_TX_STYLES);
  if (!txStyles) return null;
  if (phType === 'title' || phType === 'ctrTitle') {
    return firstChildElement(txStyles, NAME_P_TITLE_STYLE);
  }
  // Body / null-typed (= body default) / subTitle all inherit from bodyStyle.
  if (phType === 'body' || phType === 'subTitle' || phType === null) {
    return firstChildElement(txStyles, NAME_P_BODY_STYLE);
  }
  // Footer / date / sldNum / etc. inherit from otherStyle.
  return firstChildElement(txStyles, NAME_P_OTHER_STYLE);
};

/**
 * Resolves a run's effective character properties by walking the
 * ECMA-376 §21.1.2.4.7 inheritance chain — run rPr → endParaRPr →
 * pPr defRPr → text-body lstStyle → layout placeholder lstStyle →
 * master placeholder lstStyle + master txStyles → theme fontScheme.
 *
 * Each property (font, size, color, bold, italic, underline) is
 * resolved independently: the innermost layer that supplies a value
 * wins for that one property.
 *
 * Returns a non-null `TextFormat`; fields the cascade couldn't
 * resolve are simply absent (the renderer falls back to placeholder
 * defaults).
 *
 * Use `getShapeRunFormat` if you only want the literal `<a:rPr>` on
 * the run without inheritance.
 */
const NAME_TX_BODY = qname('p', 'txBody', NS.pml);

export const getShapeRunFormatEffective = (
  pres: PresentationData,
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
): TextFormat => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const run = requireRun(shape, paragraphIndex, runIndex);
  const result: Partial<TextFormat> = {};

  // Theme is consulted (a) at each layer to resolve scheme tokens and
  // color transforms eagerly, so the cascade can pick the innermost layer
  // that produces a concrete color, and (b) for typeface fallback at
  // layer 7. Reading once up-front keeps the per-layer cost flat.
  const theme = getPresentationTheme(pres);
  const ctx = { theme } as const;

  // Paragraph level (0..8). `<a:pPr lvl="..">`; absent = 0.
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  let level = 0;
  if (pPr) {
    const lvlAttr = getAttrValue(pPr, ATTR_LVL);
    if (lvlAttr !== null) {
      const parsed = Number.parseInt(lvlAttr, 10);
      if (Number.isFinite(parsed)) level = parsed;
    }
  }

  // 1. Run's own rPr.
  const runRPr = firstChildElement(run, NAME_A_RPR);
  if (runRPr) mergeRPrLayer(result, parseRPrLikeElement(runRPr, ctx));

  // 2. endParaRPr — applies to the last run in the paragraph per the spec.
  const runs = runsOf(paragraph);
  if (runs.length > 0 && runs[runs.length - 1] === run) {
    const endRPr = firstChildElement(paragraph, NAME_A_END_PARA_RPR);
    if (endRPr) mergeRPrLayer(result, parseRPrLikeElement(endRPr, ctx));
  }

  // 3. Paragraph-level defaults (pPr/defRPr).
  if (pPr) {
    const defRPr = firstChildElement(pPr, NAME_A_DEF_RPR);
    if (defRPr) mergeRPrLayer(result, parseRPrLikeElement(defRPr, ctx));
  }

  // 4. Text-body lstStyle at the paragraph's level.
  const shapeLstStyle = findShapeLstStyleElement(shape);
  const shapeLvlDef = lstStyleLevelDefRPr(shapeLstStyle, level);
  if (shapeLvlDef) mergeRPrLayer(result, parseRPrLikeElement(shapeLvlDef, ctx));

  const phIdx = getShapePlaceholderIdx(shape);
  const phType = getShapePlaceholderType(shape);
  const isPlaceholder = shapeIsPlaceholder(shape);

  const slide = shape[SHAPE_SLIDE];
  const layout = getSlideLayout(slide);

  // Steps 5-6 are placeholder inheritance: skip them entirely for non-
  // placeholder shapes (plain text boxes), which do not read the master's
  // txStyles. Without this guard an unsized text-box run wrongly inherits the
  // master body size (e.g. 32pt) instead of the ~18pt text-box default.
  if (layout && isPlaceholder) {
    // 5. Matching placeholder on the layout — both its inline rPr-bearing
    //    paragraph children (if the layout authored prompt text) and its
    //    own lstStyle.
    const layoutPh = findPlaceholderShapeIn(layout[LAYOUT_PART].shapes, phIdx, phType);
    if (layoutPh) {
      const layoutLst = extractPlaceholderLstStyle(layoutPh.element);
      const layoutLvlDef = lstStyleLevelDefRPr(layoutLst, level);
      if (layoutLvlDef) mergeRPrLayer(result, parseRPrLikeElement(layoutLvlDef, ctx));
    }

    // 6. Walk one rel up to the slide master.
    const pkg = pres[INTERNAL_PACKAGE];
    const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
    const layoutRels = pkg.getRels(layoutPartName);
    if (layoutRels) {
      const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
      if (masterRel) {
        const masterPart = pkg.getPart(resolveTarget(layoutPartName, masterRel.target));
        if (masterPart) {
          const masterRoot = parseXml(decode(masterPart.data)).root;
          const { shapes: masterShapes } = readShapeTreeFromCsldRoot(masterRoot, 'sldMaster');
          const masterPh = findPlaceholderShapeIn(masterShapes, phIdx, phType);
          if (masterPh) {
            const masterLst = extractPlaceholderLstStyle(masterPh.element);
            const masterLvlDef = lstStyleLevelDefRPr(masterLst, level);
            if (masterLvlDef) mergeRPrLayer(result, parseRPrLikeElement(masterLvlDef, ctx));
          }
          // Master text-style defaults (title / body / other).
          const txStyle = masterTxStyleFor(masterRoot, phType);
          const txLvlDef = lstStyleLevelDefRPr(txStyle, level);
          if (txLvlDef) mergeRPrLayer(result, parseRPrLikeElement(txLvlDef, ctx));
        }
      }
    }
  }

  // 7. Theme fontScheme — typeface resolution.
  //
  // The master often writes its `<a:latin typeface="+mj-lt"/>` /
  // `+mn-lt` placeholder tokens instead of a concrete face. Those
  // tokens must be resolved against the theme to produce a real
  // typeface; otherwise renderers see literal `+mj-lt` and fall
  // back to a generic font.
  //
  // When no layer in the cascade supplied a font at all, pick the
  // major font for title-class placeholders and the minor font for
  // everything else, matching PowerPoint's defaults.
  const fonts = getPresentationFonts(pres);
  if (fonts) {
    const resolveThemeToken = (token: string): string | undefined => {
      switch (token) {
        case '+mj-lt':
          return fonts.majorLatin ?? undefined;
        case '+mn-lt':
          return fonts.minorLatin ?? undefined;
        case '+mj-ea':
          return fonts.majorEastAsian ?? undefined;
        case '+mn-ea':
          return fonts.minorEastAsian ?? undefined;
        case '+mj-cs':
          return fonts.majorComplexScript ?? undefined;
        case '+mn-cs':
          return fonts.minorComplexScript ?? undefined;
        default:
          return undefined;
      }
    };
    if (typeof result.font === 'string' && result.font.startsWith('+')) {
      const resolved = resolveThemeToken(result.font);
      if (resolved) result.font = resolved;
    }
    if (result.font === undefined) {
      const useMajor = phType === 'title' || phType === 'ctrTitle';
      const fallback = useMajor ? fonts.majorLatin : fonts.minorLatin;
      if (fallback) result.font = fallback;
    }
  }

  return result as TextFormat;
};

// -- Effective pPr cascade --------------------------------------------------
//
// Mirror of the rPr cascade for paragraph-level properties: alignment,
// indents, line spacing, paragraph spacing, rtl. Walks the same layers:
//
//   1. The paragraph's own `<a:pPr>`
//   2. The text body's `<a:lstStyle><a:lvl{N+1}pPr>` (paragraph defaults)
//   3. The matching layout placeholder's lstStyle
//   4. The matching master placeholder's lstStyle, then
//      `<p:txStyles>/{title|body|other}Style/<a:lvl{N+1}pPr>`
//
// Each property merges independently — innermost layer that supplies a
// value wins for that one property.

/** Effective paragraph properties returned by `getParagraphPropertiesEffective`. */
export interface ParagraphProperties {
  /** Horizontal alignment per `ParagraphAlignment`. */
  align: ParagraphAlignment | null;
  /** Outline level (0..8). 0 = top-level paragraph. */
  level: number;
  /** Left indent in EMU. */
  marL: number | null;
  /** Right indent in EMU. */
  marR: number | null;
  /** First-line indent in EMU; negative for hanging indents. */
  indent: number | null;
  /** Line spacing — either a percent multiplier or a fixed point value. */
  lineSpacing:
    | { readonly kind: 'pct'; readonly value: number }
    | { readonly kind: 'pts'; readonly value: number }
    | null;
  /** Space before the paragraph in points. */
  spcBefPts: number | null;
  /** Space after the paragraph in points. */
  spcAftPts: number | null;
  /** Right-to-left paragraph (`<a:pPr rtl="1"/>`). */
  rtl: boolean | null;
}

const ALIGN_TOKEN_MAP: Record<string, ParagraphProperties['align']> = {
  l: 'left',
  ctr: 'center',
  r: 'right',
  just: 'justify',
  justLow: 'justify',
  dist: 'distribute',
  thaiDist: 'distribute',
};

const parsePPrLikeElement = (pPr: XmlElement): Partial<ParagraphProperties> => {
  const out: Partial<ParagraphProperties> = {};
  const algn = getAttrValue(pPr, qname('', 'algn', ''));
  if (algn !== null && ALIGN_TOKEN_MAP[algn] !== undefined) out.align = ALIGN_TOKEN_MAP[algn];
  const marL = getAttrValue(pPr, qname('', 'marL', ''));
  if (marL !== null) {
    const n = Number.parseInt(marL, 10);
    if (Number.isFinite(n)) out.marL = n;
  }
  const marR = getAttrValue(pPr, qname('', 'marR', ''));
  if (marR !== null) {
    const n = Number.parseInt(marR, 10);
    if (Number.isFinite(n)) out.marR = n;
  }
  const indent = getAttrValue(pPr, qname('', 'indent', ''));
  if (indent !== null) {
    const n = Number.parseInt(indent, 10);
    if (Number.isFinite(n)) out.indent = n;
  }
  const rtl = getAttrValue(pPr, qname('', 'rtl', ''));
  if (rtl !== null) out.rtl = rtl === '1' || rtl === 'true';
  const lnSpc = firstChildElement(pPr, qname('a', 'lnSpc', NS.dml));
  if (lnSpc) {
    const pct = firstChildElement(lnSpc, qname('a', 'spcPct', NS.dml));
    if (pct) {
      const v = getAttrValue(pct, qname('', 'val', ''));
      if (v !== null) {
        let n = Number.parseFloat(v);
        if (Number.isFinite(n)) {
          if (Math.abs(n) > 1) n = n / 100000;
          out.lineSpacing = { kind: 'pct', value: n };
        }
      }
    } else {
      const pts = firstChildElement(lnSpc, qname('a', 'spcPts', NS.dml));
      if (pts) {
        const v = getAttrValue(pts, qname('', 'val', ''));
        if (v !== null) {
          const n = Number.parseInt(v, 10);
          if (Number.isFinite(n)) out.lineSpacing = { kind: 'pts', value: n / 100 };
        }
      }
    }
  }
  const readSpcSide = (local: 'spcBef' | 'spcAft'): number | null => {
    const side = firstChildElement(pPr, qname('a', local, NS.dml));
    if (!side) return null;
    const pts = firstChildElement(side, qname('a', 'spcPts', NS.dml));
    if (!pts) return null;
    const v = getAttrValue(pts, qname('', 'val', ''));
    if (v === null) return null;
    const n = Number.parseInt(v, 10);
    return Number.isFinite(n) ? n / 100 : null;
  };
  const before = readSpcSide('spcBef');
  if (before !== null) out.spcBefPts = before;
  const after = readSpcSide('spcAft');
  if (after !== null) out.spcAftPts = after;
  return out;
};

const mergePPrLayer = (
  base: Partial<ParagraphProperties>,
  layer: Partial<ParagraphProperties>,
): void => {
  if (base.align === undefined && layer.align !== undefined) base.align = layer.align;
  if (base.marL === undefined && layer.marL !== undefined) base.marL = layer.marL;
  if (base.marR === undefined && layer.marR !== undefined) base.marR = layer.marR;
  if (base.indent === undefined && layer.indent !== undefined) base.indent = layer.indent;
  if (base.rtl === undefined && layer.rtl !== undefined) base.rtl = layer.rtl;
  if (base.lineSpacing === undefined && layer.lineSpacing !== undefined) {
    base.lineSpacing = layer.lineSpacing;
  }
  if (base.spcBefPts === undefined && layer.spcBefPts !== undefined) {
    base.spcBefPts = layer.spcBefPts;
  }
  if (base.spcAftPts === undefined && layer.spcAftPts !== undefined) {
    base.spcAftPts = layer.spcAftPts;
  }
};

/**
 * Resolves a paragraph's effective properties by walking the same
 * inheritance chain `getShapeRunFormatEffective` uses, but for the
 * paragraph-level surface:
 *
 *   - alignment, indent (left / right / first-line), line spacing,
 *     paragraph spacing (before / after), rtl.
 *
 * Each property is resolved independently; the innermost layer that
 * sets it wins. Fields the cascade can't resolve come through as `null`
 * so renderers know to fall back to their own defaults.
 *
 * Companion to `getParagraphAlignment` / `getParagraphLineSpacing` /
 * `getParagraphIndent` / `getParagraphSpacing`, which only surface the
 * literal `<a:pPr>` and skip the layout / master cascade.
 */
export const getParagraphPropertiesEffective = (
  pres: PresentationData,
  shape: SlideShapeData,
  paragraphIndex: number,
): ParagraphProperties => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);

  let level = 0;
  if (pPr) {
    const lvlAttr = getAttrValue(pPr, ATTR_LVL);
    if (lvlAttr !== null) {
      const parsed = Number.parseInt(lvlAttr, 10);
      if (Number.isFinite(parsed)) level = parsed;
    }
  }

  const result: Partial<ParagraphProperties> = {};

  // 1. Paragraph's own pPr.
  if (pPr) mergePPrLayer(result, parsePPrLikeElement(pPr));

  // 2. Text-body lstStyle at the paragraph's level.
  const shapeLstStyle = findShapeLstStyleElement(shape);
  const shapeLvlPPr = lstStyleLevelPPr(shapeLstStyle, level);
  if (shapeLvlPPr) mergePPrLayer(result, parsePPrLikeElement(shapeLvlPPr));

  const phIdx = getShapePlaceholderIdx(shape);
  const phType = getShapePlaceholderType(shape);
  const isPlaceholder = shapeIsPlaceholder(shape);
  const slide = shape[SHAPE_SLIDE];
  const layout = getSlideLayout(slide);

  // Placeholder inheritance only: a plain text box does not read the master's
  // txStyles for paragraph defaults (align / indent / spacing) either.
  if (layout && isPlaceholder) {
    // 3. Layout placeholder lstStyle.
    const layoutPh = findPlaceholderShapeIn(layout[LAYOUT_PART].shapes, phIdx, phType);
    if (layoutPh) {
      const layoutLst = extractPlaceholderLstStyle(layoutPh.element);
      const layoutLvlPPr = lstStyleLevelPPr(layoutLst, level);
      if (layoutLvlPPr) mergePPrLayer(result, parsePPrLikeElement(layoutLvlPPr));
    }

    // 4. Master placeholder lstStyle + master txStyles.
    const pkg = pres[INTERNAL_PACKAGE];
    const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
    const layoutRels = pkg.getRels(layoutPartName);
    if (layoutRels) {
      const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
      if (masterRel) {
        const masterPart = pkg.getPart(resolveTarget(layoutPartName, masterRel.target));
        if (masterPart) {
          const masterRoot = parseXml(decode(masterPart.data)).root;
          const { shapes: masterShapes } = readShapeTreeFromCsldRoot(masterRoot, 'sldMaster');
          const masterPh = findPlaceholderShapeIn(masterShapes, phIdx, phType);
          if (masterPh) {
            const masterLst = extractPlaceholderLstStyle(masterPh.element);
            const masterLvlPPr = lstStyleLevelPPr(masterLst, level);
            if (masterLvlPPr) mergePPrLayer(result, parsePPrLikeElement(masterLvlPPr));
          }
          const txStyle = masterTxStyleFor(masterRoot, phType);
          const txLvlPPr = lstStyleLevelPPr(txStyle, level);
          if (txLvlPPr) mergePPrLayer(result, parsePPrLikeElement(txLvlPPr));
        }
      }
    }
  }

  return {
    align: result.align ?? null,
    level,
    marL: result.marL ?? null,
    marR: result.marR ?? null,
    indent: result.indent ?? null,
    lineSpacing: result.lineSpacing ?? null,
    spcBefPts: result.spcBefPts ?? null,
    spcAftPts: result.spcAftPts ?? null,
    rtl: result.rtl ?? null,
  };
};

/**
 * Applies `format` to a single run. Run-property attributes not
 * addressed by `format` are preserved — partial updates compose.
 *
 * Example: bold the second word of the first paragraph:
 *
 *   setShapeRunFormat(shape, 0, 1, { bold: true, color: '#FF0000' });
 */
export const setShapeRunFormat = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
  format: TextFormat,
): void => {
  const run = requireRun(shape, paragraphIndex, runIndex);
  const rPr = ensureRPr(run);
  applyRunFormatInternal(rPr, format);
  commitAndRefresh(shape);
};

/**
 * Reads the external URL the first run in the shape's text-body links
 * to (set via `setShapeHyperlink`). Returns `null` when no run carries
 * an `<a:hlinkClick r:id=…/>` or the rId resolves to a non-hyperlink
 * target.
 */
export const getShapeHyperlink = (shape: SlideShapeData): string | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'shape') return null;
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (!txBody) return null;
  for (const p of txBody.children) {
    if (p.kind !== 'element' || p.name.namespaceURI !== NS.dml || p.name.localName !== 'p')
      continue;
    for (const r of p.children) {
      if (r.kind !== 'element' || r.name.namespaceURI !== NS.dml || r.name.localName !== 'r')
        continue;
      const rPr = firstChildElement(r, qname('a', 'rPr', NS.dml));
      if (!rPr) continue;
      const hlink = firstChildElement(rPr, qname('a', 'hlinkClick', NS.dml));
      if (!hlink) continue;
      const rId = getAttrValue(hlink, qname('r', 'id', NS.officeDocRels));
      if (!rId) continue;
      const slide = shape[SHAPE_SLIDE];
      const rels = slide[INTERNAL_PACKAGE].getRels(slide[SLIDE_PART_NAME]);
      if (!rels) continue;
      const rel = rels.items.find((x) => x.id === rId);
      if (rel?.type === REL_TYPES.hyperlink && rel.targetMode === 'External') {
        return rel.target;
      }
    }
  }
  return null;
};

/**
 * Sets an external hyperlink on every run in the shape's text. Allocates
 * (or reuses) a `hyperlink` relationship on the slide's `.rels`. Pass
 * `null` to clear.
 */
export const setShapeHyperlink = (
  shape: SlideShapeData,
  url: string | null,
  tooltip?: string,
): void => {
  const slide = shape[SHAPE_SLIDE];
  const txBody = requireTxBody(shape);
  if (url === null) {
    applyHyperlinkToAllRuns(txBody, null);
  } else {
    const pkg = slide[INTERNAL_PACKAGE];
    const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
    const existing = rels.items.find(
      (r) => r.type === REL_TYPES.hyperlink && r.target === url && r.targetMode === 'External',
    );
    const rId =
      existing?.id ??
      (() => {
        const nextId = nextRelId(rels.items.map((r) => r.id));
        rels.items.push({
          id: nextId,
          type: REL_TYPES.hyperlink,
          target: url,
          targetMode: 'External',
        });
        pkg.setRels(slide[SLIDE_PART_NAME], rels);
        return nextId;
      })();
    applyHyperlinkToAllRuns(txBody, rId, tooltip);
  }
  commitAndRefresh(shape);
};

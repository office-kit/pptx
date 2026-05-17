// Tree-shakeable free-function entry points — the canonical public API.
//
// Every operation is a standalone export that operates on the opaque
// `PresentationData` / `SlideData` interfaces defined in
// `_internal-symbols.ts`. Consumers can import only what they use and
// modern bundlers drop the rest.

import {
  type BulletStyle,
  type GradientFillOptions,
  type ParagraphAlignment,
  type PatternFillOptions,
  type Position,
  type Size,
  type StrokeOptions,
  type TextFormat,
  applyAlignmentToAllParagraphs,
  applyBulletToAllParagraphs,
  applyBulletToParagraph,
  applyFormatToAllRuns,
  applyHyperlinkToAllRuns,
  applyRunFormat as applyRunFormatInternal,
  clearEffects as clearEffectsImpl,
  clearFill as clearFillImpl,
  clearStroke as clearStrokeImpl,
  type ArrowOptions,
  type GlowOptions,
  type LineDash,
  type ShadowOptions,
  getPictureEmbedRId,
  readFlip,
  readPosition,
  readRotation,
  readSize,
  replaceTextInTree,
  replaceTokensInTree,
  setFlip as writeFlip,
  setGlow,
  setGradientFill,
  setPatternFill,
  setShadow,
  setNoFill as setNoFillImpl,
  setNoStroke as setNoStrokeImpl,
  setPosition as writePosition,
  setRotation as writeRotation,
  setSize as writeSize,
  setSolidFill,
  setSolidStroke,
  setStrokeArrow,
  setStrokeDash,
  setTextBody,
} from '../internal/drawingml/index.ts';
import type { Emu } from './units.ts';
import {
  basename,
  contentTypeForFormat,
  detectImageFormat,
  emptyRels,
  extensionForFormat,
  type ImageFormat,
  nextRelId,
  type PartName,
  partName,
  relsPartNameFor,
  resolveTarget,
} from '../internal/opc/index.ts';
import { OpcPackage } from '../internal/parts/index.ts';
import {
  REL_TYPES,
  type AnimationEffect,
  type AnimationOptions,
  type CommentAuthor,
  type CommentPosition,
  type PresetShape,
  type ShapeKind,
  type SlideComment,
  type SlideLayoutType,
  type TransitionOptions,
  buildCommentAuthorListDoc,
  buildCommentListDoc,
  buildConnector,
  buildEmptyNotesSlide,
  buildPicture,
  buildShape,
  buildSingleEffectTiming,
  buildSlideFromLayout,
  buildTable,
  buildTableCell,
  buildTableRow,
  buildTextBox,
  buildTransition,
  readCommentAuthorList,
  readCommentList,
  readPresentationPart,
  readGroupChildren,
  readShapeTreeFromCsldRoot,
  readSlideLayoutPart,
  readSlidePart,
  slideText,
} from '../internal/presentationml/index.ts';
import {
  type IssueSeverity,
  type ValidationIssue,
  validatePresentationPackage,
} from '../internal/validator/index.ts';
import {
  type ChartKind,
  type ChartSeries,
  type ChartSpec,
  buildChartSpaceDoc,
  buildEmbeddedXlsx,
  readChartSpec,
} from '../internal/chartml/index.ts';
import {
  NS,
  type XmlDocument,
  type XmlElement,
  allChildElements,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  parseXml,
  qname,
  serializeXml,
  text as textNode,
} from '../internal/xml/index.ts';
import {
  CELL_COL,
  CELL_ELEMENT,
  CELL_ROW,
  CELL_TABLE,
  COMMENT_SLIDE,
  COMMENT_SNAPSHOT,
  INTERNAL_PACKAGE,
  LAYOUT_PART,
  LAYOUT_PART_NAME,
  type PresentationData,
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  SHAPE_SNAPSHOT,
  SLIDE_DOCUMENT,
  SLIDE_PART,
  SLIDE_PART_NAME,
  SLIDE_SHAPES,
  type SlideCommentData,
  type SlideData,
  type SlideLayoutData,
  type SlideShapeData,
  type TableCellData,
} from './_internal-symbols.ts';

const TEXT_DECODER = new TextDecoder();
const TEXT_ENCODER = new TextEncoder();
const decode = (b: Uint8Array): string => TEXT_DECODER.decode(b);
const encode = (s: string): Uint8Array => TEXT_ENCODER.encode(s);

const SLIDE_LAYOUT_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml';
const SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const PRES_PART_NAME = partName('/ppt/presentation.xml');

const NAME_PRESENTATION = qname('p', 'presentation', NS.pml);
const NAME_SLD_MASTER_ID_LST = qname('p', 'sldMasterIdLst', NS.pml);
const NAME_SLD_ID_LST = qname('p', 'sldIdLst', NS.pml);
const NAME_SLD_ID = qname('p', 'sldId', NS.pml);
const NAME_CSLD = qname('p', 'cSld', NS.pml);
const NAME_SP_TREE = qname('p', 'spTree', NS.pml);
const ATTR_ID = qname('', 'id', '');
const ATTR_R_ID = qname('r', 'id', NS.officeDocRels);

// PowerPoint accepts sldIds in [256, 2³¹−1024]. See plan §Risks.
const SLD_ID_MIN = 256;
const SLD_ID_MAX = 2147482623;

/**
 * Anything that can be turned into a `Uint8Array` of PPTX bytes:
 *
 *   - `Uint8Array` — used as-is.
 *   - `ArrayBuffer` — wrapped without copying.
 *   - `Blob` or `File` — read via `arrayBuffer()`.
 */
export type PresentationInput = Uint8Array | ArrayBuffer | Blob;

const normalize = async (input: PresentationInput): Promise<Uint8Array> => {
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  if (typeof Blob !== 'undefined' && input instanceof Blob) {
    return new Uint8Array(await input.arrayBuffer());
  }
  throw new TypeError('loadPresentation: expected Uint8Array | ArrayBuffer | Blob');
};

/**
 * Loads an existing `.pptx` and returns a `PresentationData` value.
 */
export const loadPresentation = async (input: PresentationInput): Promise<PresentationData> => {
  const bytes = await normalize(input);
  const pkg = OpcPackage.load(bytes);
  return { [INTERNAL_PACKAGE]: pkg, _slidesCache: null };
};

/**
 * Creates a fresh, empty `PresentationData`. The result is NOT yet a
 * valid PPTX — it carries only the OPC defaults.
 */
export const createPresentation = (): PresentationData => {
  const pkg = OpcPackage.empty();
  return { [INTERNAL_PACKAGE]: pkg, _slidesCache: null };
};

/**
 * Serializes a presentation back to PPTX bytes.
 */
export const savePresentation = (pres: PresentationData): Promise<Uint8Array> => {
  return Promise.resolve(pres[INTERNAL_PACKAGE].save());
};

// ---------------------------------------------------------------------------
// Slide sections (p14 extension).
//
// PowerPoint 2010+ supports grouping slides into named sections. The
// data lives in `<p:presentation>/<p:extLst>/<p:ext uri="{…}">/<p14:sectionLst>`.
// We expose it as a flat array of `SlideSection` objects, each
// carrying the section name and the slides in display order.

/** One section in the deck. `slides` is a snapshot at read time. */
export interface SlideSection {
  readonly name: string;
  readonly slides: ReadonlyArray<SlideData>;
}

const NS_P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const SECTION_LST_EXT_URI = '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}';

const NAME_EXT_LST = qname('p', 'extLst', NS.pml);
const NAME_P_EXT = qname('p', 'ext', NS.pml);
const NAME_P14_SECTION_LST = qname('p14', 'sectionLst', NS_P14);
const NAME_P14_SECTION = qname('p14', 'section', NS_P14);
const NAME_P14_SLD_ID_LST = qname('p14', 'sldIdLst', NS_P14);
const NAME_P14_SLD_ID = qname('p14', 'sldId', NS_P14);
const ATTR_URI = qname('', 'uri', '');
const ATTR_NAME_SEC = qname('', 'name', '');
const ATTR_ID_SEC = qname('', 'id', '');
const ATTR_SLD_ID_REF = qname('', 'id', '');

const findSectionLstElement = (presDocRoot: XmlElement): XmlElement | null => {
  const extLst = firstChildElement(presDocRoot, NAME_EXT_LST);
  if (!extLst) return null;
  for (const ext of allChildElements(extLst, NAME_P_EXT)) {
    if (getAttrValue(ext, ATTR_URI) === SECTION_LST_EXT_URI) {
      return firstChildElement(ext, NAME_P14_SECTION_LST);
    }
  }
  return null;
};

const ensureSectionLst = (presDocRoot: XmlElement): XmlElement => {
  let extLst = firstChildElement(presDocRoot, NAME_EXT_LST);
  if (!extLst) {
    extLst = elem(NAME_EXT_LST);
    presDocRoot.children.push(extLst);
  }
  for (const ext of allChildElements(extLst, NAME_P_EXT)) {
    if (getAttrValue(ext, ATTR_URI) === SECTION_LST_EXT_URI) {
      let sectionLst = firstChildElement(ext, NAME_P14_SECTION_LST);
      if (!sectionLst) {
        sectionLst = elem(NAME_P14_SECTION_LST);
        ext.children.push(sectionLst);
      }
      return sectionLst;
    }
  }
  const sectionLst = elem(NAME_P14_SECTION_LST, {
    prefixDecls: new Map([['p14', NS_P14]]),
  });
  const ext = elem(NAME_P_EXT, {
    attrs: [attr(ATTR_URI, SECTION_LST_EXT_URI)],
    children: [sectionLst],
  });
  extLst.children.push(ext);
  return sectionLst;
};

/**
 * Returns every section defined in `<p14:sectionLst>`, with each
 * section's slides resolved to `SlideData` handles via the
 * presentation's `<p:sldIdLst>`. Sections referencing missing slides
 * are dropped silently; empty arrays are preserved.
 *
 * Returns an empty array when no sectionLst is present.
 */
export const getSlideSections = (pres: PresentationData): ReadonlyArray<SlideSection> => {
  const pkg = pres[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (!presPart) return [];
  const doc = parseXml(decode(presPart.data));
  const sectionLst = findSectionLstElement(doc.root);
  if (!sectionLst) return [];

  // Build a map from `sldId` to SlideData by walking the deck.
  const slideById = new Map<string, SlideData>();
  const presRels = pkg.getRels(PRES_PART_NAME);
  if (presRels) {
    const presModel = readPresentationPart(doc.root);
    const slides = getSlides(pres);
    // Map presModel.slides[i].id → slides[i] (they're in the same order).
    for (let i = 0; i < presModel.slides.length && i < slides.length; i++) {
      slideById.set(String(presModel.slides[i]!.id), slides[i]!);
    }
  }

  const out: SlideSection[] = [];
  for (const sec of allChildElements(sectionLst, NAME_P14_SECTION)) {
    const name = getAttrValue(sec, ATTR_NAME_SEC) ?? '';
    const sldIdLst = firstChildElement(sec, NAME_P14_SLD_ID_LST);
    const slides: SlideData[] = [];
    if (sldIdLst) {
      for (const sldId of allChildElements(sldIdLst, NAME_P14_SLD_ID)) {
        const id = getAttrValue(sldId, ATTR_SLD_ID_REF);
        if (id !== null) {
          const slide = slideById.get(id);
          if (slide) slides.push(slide);
        }
      }
    }
    out.push({ name, slides });
  }
  return out;
};

/**
 * Replaces the deck's section list with `sections`. Each section is
 * given a fresh GUID `id` attribute (PowerPoint generates one per
 * section; we synthesize a deterministic-ish one based on index +
 * timestamp for v1).
 *
 * Pass `[]` to clear all sections — the helper drops the
 * `<p14:sectionLst>` extension entirely when no sections remain.
 */
export const setSlideSections = (
  pres: PresentationData,
  sections: ReadonlyArray<{ name: string; slides: ReadonlyArray<SlideData> }>,
): void => {
  const pkg = pres[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (!presPart) throw new Error('presentation.xml is missing');
  const doc = parseXml(decode(presPart.data));

  // Build a reverse map from SlideData identity → sldId.
  const idBySlide = new Map<SlideData, string>();
  const presModel = readPresentationPart(doc.root);
  const slides = getSlides(pres);
  for (let i = 0; i < presModel.slides.length && i < slides.length; i++) {
    idBySlide.set(slides[i]!, String(presModel.slides[i]!.id));
  }

  // Reach for the matching slide id by part name, since SlideData
  // identity may have been rebuilt across calls.
  const idByPartName = new Map<string, string>();
  if (presModel.slides.length > 0) {
    const presRels = pkg.getRels(PRES_PART_NAME);
    if (presRels) {
      for (const s of presModel.slides) {
        const rel = presRels.items.find((r) => r.id === s.rId);
        if (!rel) continue;
        const slideName = rel.target.startsWith('/') ? rel.target : `/ppt/${rel.target}`;
        idByPartName.set(slideName, String(s.id));
      }
    }
  }
  const sldIdFor = (slide: SlideData): string | null => {
    const direct = idBySlide.get(slide);
    if (direct) return direct;
    return idByPartName.get(slide[SLIDE_PART_NAME]) ?? null;
  };

  if (sections.length === 0) {
    // Drop the sectionLst entirely.
    const extLst = firstChildElement(doc.root, NAME_EXT_LST);
    if (extLst) {
      extLst.children = extLst.children.filter(
        (c) =>
          !(
            c.kind === 'element' &&
            c.name.namespaceURI === NS.pml &&
            c.name.localName === 'ext' &&
            getAttrValue(c, ATTR_URI) === SECTION_LST_EXT_URI
          ),
      );
      // Drop the empty extLst itself if no other extensions remain.
      if (extLst.children.length === 0) {
        doc.root.children = doc.root.children.filter((c) => c !== extLst);
      }
    }
    presPart.data = encode(serializeXml(doc));
    return;
  }

  const sectionLst = ensureSectionLst(doc.root);
  sectionLst.children = sections.map((section, i) => {
    // Synthesize a GUID-shaped id from index + timestamp.
    const ts = Date.now().toString(16).padStart(8, '0').slice(-8).toUpperCase();
    const id = `{${ts.slice(0, 8)}-${String(i).padStart(4, '0')}-4000-8000-000000000000}`;
    const sldIds: XmlElement[] = [];
    for (const slide of section.slides) {
      const sldId = sldIdFor(slide);
      if (sldId !== null) {
        sldIds.push(elem(NAME_P14_SLD_ID, { attrs: [attr(ATTR_SLD_ID_REF, sldId)] }));
      }
    }
    return elem(NAME_P14_SECTION, {
      attrs: [attr(ATTR_NAME_SEC, section.name), attr(ATTR_ID_SEC, id)],
      children: [elem(NAME_P14_SLD_ID_LST, { children: sldIds })],
    });
  });

  presPart.data = encode(serializeXml(doc));
};

// ---------------------------------------------------------------------------
// Slide layouts.

/** PowerPoint's user-visible layout name. */
export const getSlideLayoutName = (layout: SlideLayoutData): string => layout[LAYOUT_PART].name;

/**
 * Returns the package part name (e.g. `/ppt/slideLayouts/slideLayout3.xml`)
 * of `layout`. Useful for surfacing layouts in validator output and
 * other path-keyed UIs.
 */
export const getSlideLayoutPartName = (layout: SlideLayoutData): string => layout[LAYOUT_PART_NAME];

/**
 * Returns the slide layout whose package part name equals
 * `partName`, or `null` when no such layout exists. Mirror of
 * `findSlideByPartName` for layouts.
 */
export const findSlideLayoutByPartName = (
  pres: PresentationData,
  partName: string,
): SlideLayoutData | null => {
  for (const layout of getSlideLayouts(pres)) {
    if (layout[LAYOUT_PART_NAME] === partName) return layout;
  }
  return null;
};

/**
 * Read-only view of one placeholder on a slide layout. Surfaces the
 * three fields a slide-author cares about when binding a slide to a
 * layout: which slot is for the title, which is for the body, etc.
 */
export interface SlideLayoutPlaceholder {
  /** `<p:ph type="...">`. `null` when omitted — spec default is `body`. */
  readonly type: string | null;
  /** `<p:ph idx="...">`. `null` when omitted — spec default is `0`. */
  readonly idx: number | null;
  /** `<p:cNvPr name="...">` — what PowerPoint shows in the selection pane. */
  readonly name: string;
  /**
   * Layout-defined position + size in EMU. A slide placeholder with no
   * `<a:xfrm>` of its own inherits these. `null` when the layout
   * placeholder also lacks an explicit transform (rare — usually the
   * master defines it then).
   */
  readonly bounds: ShapeBounds | null;
}

/**
 * Enumerates the placeholder shapes on a slide layout. Non-placeholder
 * shapes (decorative rectangles, watermarks added to the layout) are
 * filtered out; only entries with a `<p:ph>` element are returned.
 *
 * Use this when you need to discover which placeholder indices a
 * layout exposes — e.g. before `findSlidePlaceholder(slide, ...)` to
 * confirm the slot exists.
 */
export const getSlideLayoutPlaceholders = (
  layout: SlideLayoutData,
): ReadonlyArray<SlideLayoutPlaceholder> => {
  const out: SlideLayoutPlaceholder[] = [];
  for (const shape of layout[LAYOUT_PART].shapes) {
    // Only `p:sp` shapes carry placeholders in real templates; pictures
    // and connectors can technically have `<p:ph>` per the schema but
    // PowerPoint never authors that. Filter for safety either way.
    if (shape.placeholderType === null && shape.placeholderIdx === null) continue;
    const pos = readPosition(shape.element, shape.kind);
    const size = readSize(shape.element, shape.kind);
    const bounds: ShapeBounds | null =
      pos === null || size === null
        ? null
        : { x: pos.x as Emu, y: pos.y as Emu, w: size.w as Emu, h: size.h as Emu };
    out.push({
      type: shape.placeholderType,
      idx: shape.placeholderIdx,
      name: shape.name,
      bounds,
    });
  }
  return out;
};

// ---------------------------------------------------------------------------
// Theme.

/**
 * The named color scheme on a presentation's theme. Each slot is a
 * `#RRGGBB` string — `sysClr` slots are flattened to their cached
 * `lastClr` value.
 */
export interface PresentationTheme {
  readonly name: string;
  readonly dark1: string;
  readonly light1: string;
  readonly dark2: string;
  readonly light2: string;
  readonly accent1: string;
  readonly accent2: string;
  readonly accent3: string;
  readonly accent4: string;
  readonly accent5: string;
  readonly accent6: string;
  readonly hyperlink: string;
  readonly followedHyperlink: string;
}

const THEME_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.theme+xml';

const NAME_THEME_ELEMENTS = qname('a', 'themeElements', NS.dml);
const NAME_CLR_SCHEME = qname('a', 'clrScheme', NS.dml);
const NAME_SRGB_CLR = qname('a', 'srgbClr', NS.dml);
const NAME_SYS_CLR = qname('a', 'sysClr', NS.dml);

const readSchemeSlot = (parent: XmlElement, local: string): string => {
  const slot = firstChildElement(parent, qname('a', local, NS.dml));
  if (!slot) return '';
  const srgb = firstChildElement(slot, NAME_SRGB_CLR);
  if (srgb) {
    const v = getAttrValue(srgb, qname('', 'val', ''));
    if (v) return `#${v.toUpperCase()}`;
  }
  const sys = firstChildElement(slot, NAME_SYS_CLR);
  if (sys) {
    const last = getAttrValue(sys, qname('', 'lastClr', ''));
    if (last) return `#${last.toUpperCase()}`;
  }
  return '';
};

/**
 * Returns the first theme's color scheme as `#RRGGBB` strings, or
 * `null` if the package carries no theme. Each accent slot maps
 * directly to the `accent1`–`accent6` chart palette defaults.
 *
 * Multi-master decks may carry several themes — v1 surfaces only the
 * first one found (alphabetical by part name). Per-master theme
 * lookup will land if a concrete user need shows up.
 */
export const getPresentationTheme = (pres: PresentationData): PresentationTheme | null => {
  const pkg = pres[INTERNAL_PACKAGE];
  const themePart = pkg.parts
    .filter((p) => p.contentType === THEME_CONTENT_TYPE)
    .sort((a, b) => a.name.localeCompare(b.name))[0];
  if (!themePart) return null;
  const root = parseXml(decode(themePart.data)).root;
  const themeElements = firstChildElement(root, NAME_THEME_ELEMENTS);
  if (!themeElements) return null;
  const clrScheme = firstChildElement(themeElements, NAME_CLR_SCHEME);
  if (!clrScheme) return null;
  return {
    name: getAttrValue(clrScheme, qname('', 'name', '')) ?? '',
    dark1: readSchemeSlot(clrScheme, 'dk1'),
    light1: readSchemeSlot(clrScheme, 'lt1'),
    dark2: readSchemeSlot(clrScheme, 'dk2'),
    light2: readSchemeSlot(clrScheme, 'lt2'),
    accent1: readSchemeSlot(clrScheme, 'accent1'),
    accent2: readSchemeSlot(clrScheme, 'accent2'),
    accent3: readSchemeSlot(clrScheme, 'accent3'),
    accent4: readSchemeSlot(clrScheme, 'accent4'),
    accent5: readSchemeSlot(clrScheme, 'accent5'),
    accent6: readSchemeSlot(clrScheme, 'accent6'),
    hyperlink: readSchemeSlot(clrScheme, 'hlink'),
    followedHyperlink: readSchemeSlot(clrScheme, 'folHlink'),
  };
};

/**
 * The theme's font scheme, flattened to the typefaces the runs in a
 * deck inherit by default. `major*` is the heading font (slide
 * titles, chart titles); `minor*` is the body font.
 *
 * Each field carries the Latin / East-Asian / Complex-Script
 * typeface name as written on the theme. Empty values are
 * normalized to `null`.
 */
export interface PresentationFonts {
  readonly majorLatin: string | null;
  readonly majorEastAsian: string | null;
  readonly majorComplexScript: string | null;
  readonly minorLatin: string | null;
  readonly minorEastAsian: string | null;
  readonly minorComplexScript: string | null;
}

const readTypeface = (parent: XmlElement | null, local: string): string | null => {
  if (!parent) return null;
  const el = firstChildElement(parent, qname('a', local, NS.dml));
  if (!el) return null;
  const v = getAttrValue(el, qname('', 'typeface', ''));
  if (!v) return null;
  return v;
};

/**
 * Returns the first theme's font scheme, or `null` when the package
 * carries no theme. As with `getPresentationTheme`, multi-master
 * decks surface only the first theme found (alphabetical by part
 * name); per-master font lookup will land if needed.
 */
export const getPresentationFonts = (pres: PresentationData): PresentationFonts | null => {
  const pkg = pres[INTERNAL_PACKAGE];
  const themePart = pkg.parts
    .filter((p) => p.contentType === THEME_CONTENT_TYPE)
    .sort((a, b) => a.name.localeCompare(b.name))[0];
  if (!themePart) return null;
  const root = parseXml(decode(themePart.data)).root;
  const themeElements = firstChildElement(root, NAME_THEME_ELEMENTS);
  if (!themeElements) return null;
  const fontScheme = firstChildElement(themeElements, qname('a', 'fontScheme', NS.dml));
  if (!fontScheme) return null;
  const majorFont = firstChildElement(fontScheme, qname('a', 'majorFont', NS.dml));
  const minorFont = firstChildElement(fontScheme, qname('a', 'minorFont', NS.dml));
  return {
    majorLatin: readTypeface(majorFont, 'latin'),
    majorEastAsian: readTypeface(majorFont, 'ea'),
    majorComplexScript: readTypeface(majorFont, 'cs'),
    minorLatin: readTypeface(minorFont, 'latin'),
    minorEastAsian: readTypeface(minorFont, 'ea'),
    minorComplexScript: readTypeface(minorFont, 'cs'),
  };
};

// ---------------------------------------------------------------------------
// Core properties (`/docProps/core.xml`).

const NS_CORE_PROPS = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
const NS_DC = 'http://purl.org/dc/elements/1.1/';
const NS_DCTERMS = 'http://purl.org/dc/terms/';
const CORE_PROPS_PART_NAME = partName('/docProps/core.xml');

/**
 * Document-level metadata from `/docProps/core.xml` (Open Packaging
 * Conventions). Surfaces the fields PowerPoint, Keynote, and
 * everyone else exchange via OPC core-properties — these are the
 * values shown in PowerPoint's "File › Properties" / "Info" panel.
 */
export interface CoreProperties {
  readonly title: string | null;
  readonly subject: string | null;
  readonly creator: string | null;
  readonly keywords: string | null;
  readonly description: string | null;
  readonly lastModifiedBy: string | null;
  readonly revision: string | null;
  /** ISO-8601 timestamp string when set; `null` otherwise. */
  readonly created: string | null;
  /** ISO-8601 timestamp string when set; `null` otherwise. */
  readonly modified: string | null;
  readonly category: string | null;
}

/**
 * Reads `/docProps/core.xml`. Returns `null` when the package has
 * no core-properties part. Each field is `null` when the
 * corresponding element is absent or empty.
 */
/**
 * Convenience: bumps core-properties' `cp:revision` by one (treating
 * an unset / unparseable value as 0). Returns the new revision
 * number. Useful right before `savePresentation` so consumers can
 * tell decks apart.
 */
export const incrementRevision = (pres: PresentationData): number => {
  const props = getCoreProperties(pres);
  const current =
    props?.revision === null || props?.revision === undefined
      ? 0
      : Number.parseInt(props.revision, 10);
  const next = (Number.isFinite(current) ? current : 0) + 1;
  setCoreProperties(pres, { revision: String(next) });
  return next;
};

/**
 * Convenience: writes `new Date().toISOString()` to
 * `dcterms:modified`. Useful right before `savePresentation` so
 * "last edited" shows the actual save time. Pass an explicit
 * `Date` to set a specific value.
 */
export const touchModified = (pres: PresentationData, at: Date = new Date()): void => {
  setCoreProperties(pres, { modified: at.toISOString() });
};

/**
 * Convenience: the timestamp from core-properties' `dcterms:created`,
 * parsed as a `Date`. Returns `null` when no created field is set
 * or the value isn't a recognizable W3C-DTF / ISO-8601 string.
 */
export const getPresentationCreated = (pres: PresentationData): Date | null => {
  const props = getCoreProperties(pres);
  if (!props || props.created === null) return null;
  const d = new Date(props.created);
  return Number.isFinite(d.getTime()) ? d : null;
};

/**
 * Convenience: the timestamp from core-properties' `dcterms:modified`,
 * parsed as a `Date`. Returns `null` when no modified field is set
 * or the value isn't a recognizable W3C-DTF / ISO-8601 string.
 */
export const getPresentationModified = (pres: PresentationData): Date | null => {
  const props = getCoreProperties(pres);
  if (!props || props.modified === null) return null;
  const d = new Date(props.modified);
  return Number.isFinite(d.getTime()) ? d : null;
};

export const getCoreProperties = (pres: PresentationData): CoreProperties | null => {
  const pkg = pres[INTERNAL_PACKAGE];
  const part = pkg.getPart(CORE_PROPS_PART_NAME);
  if (!part) return null;
  const root = parseXml(decode(part.data)).root;
  const read = (uri: string, local: string): string | null => {
    const el = firstChildElement(root, qname('', local, uri));
    if (!el) return null;
    let s = '';
    for (const c of el.children) if (c.kind === 'text') s += c.data;
    return s.length === 0 ? null : s;
  };
  return {
    title: read(NS_DC, 'title'),
    subject: read(NS_DC, 'subject'),
    creator: read(NS_DC, 'creator'),
    keywords: read(NS_CORE_PROPS, 'keywords'),
    description: read(NS_DC, 'description'),
    lastModifiedBy: read(NS_CORE_PROPS, 'lastModifiedBy'),
    revision: read(NS_CORE_PROPS, 'revision'),
    created: read(NS_DCTERMS, 'created'),
    modified: read(NS_DCTERMS, 'modified'),
    category: read(NS_CORE_PROPS, 'category'),
  };
};

const CORE_PROPS_CONTENT_TYPE = 'application/vnd.openxmlformats-package.core-properties+xml';

const CORE_PROP_FIELDS: ReadonlyArray<{
  key: keyof CoreProperties;
  uri: string;
  prefix: string;
  local: string;
}> = [
  { key: 'title', uri: NS_DC, prefix: 'dc', local: 'title' },
  { key: 'subject', uri: NS_DC, prefix: 'dc', local: 'subject' },
  { key: 'creator', uri: NS_DC, prefix: 'dc', local: 'creator' },
  { key: 'keywords', uri: NS_CORE_PROPS, prefix: 'cp', local: 'keywords' },
  { key: 'description', uri: NS_DC, prefix: 'dc', local: 'description' },
  { key: 'lastModifiedBy', uri: NS_CORE_PROPS, prefix: 'cp', local: 'lastModifiedBy' },
  { key: 'revision', uri: NS_CORE_PROPS, prefix: 'cp', local: 'revision' },
  { key: 'created', uri: NS_DCTERMS, prefix: 'dcterms', local: 'created' },
  { key: 'modified', uri: NS_DCTERMS, prefix: 'dcterms', local: 'modified' },
  { key: 'category', uri: NS_CORE_PROPS, prefix: 'cp', local: 'category' },
];

const buildEmptyCorePropsRoot = (): XmlElement => {
  const prefixDecls = new Map<string, string>([
    ['cp', NS_CORE_PROPS],
    ['dc', NS_DC],
    ['dcterms', NS_DCTERMS],
  ]);
  return {
    kind: 'element',
    name: qname('cp', 'coreProperties', NS_CORE_PROPS),
    attrs: [],
    prefixDecls,
    children: [],
  };
};

/**
 * Writes selected fields on `/docProps/core.xml`. Unspecified fields
 * are left as-is; pass `null` to clear a field that's currently set.
 * Bootstraps the part (and the `/_rels/.rels` entry + content-type
 * override) if the package didn't have one.
 *
 * Note: setting `created` / `modified` requires an ISO-8601 timestamp
 * string (e.g. `'2026-05-15T12:34:56Z'`). PowerPoint expects the
 * `xsi:type="dcterms:W3CDTF"` attribute on these elements but readers
 * we tested all accept missing-attribute output too; this helper
 * therefore omits the attribute for simplicity.
 */
export const setCoreProperties = (
  pres: PresentationData,
  values: Partial<CoreProperties>,
): void => {
  const pkg = pres[INTERNAL_PACKAGE];
  let part = pkg.getPart(CORE_PROPS_PART_NAME);
  let root: XmlElement;
  let doc: ReturnType<typeof parseXml>;
  if (part) {
    doc = parseXml(decode(part.data));
    root = doc.root;
  } else {
    root = buildEmptyCorePropsRoot();
    doc = { kind: 'document', decl: null, prolog: [], root, epilog: [] };
  }

  for (const field of CORE_PROP_FIELDS) {
    if (!(field.key in values)) continue;
    const value = values[field.key] ?? null;
    const name = qname(field.prefix, field.local, field.uri);
    const existing = firstChildElement(root, name);
    if (value === null) {
      if (existing) {
        existing.children = [];
      }
      continue;
    }
    if (existing) {
      existing.children = [textNode(value)];
    } else {
      root.children.push(elem(name, { children: [textNode(value)] }));
    }
  }

  const bytes = encode(serializeXml(doc));
  if (part) {
    part.data = bytes;
    return;
  }

  // Bootstrap: register override, add part, wire root rel.
  pkg.contentTypes.overrides.push({
    partName: CORE_PROPS_PART_NAME,
    contentType: CORE_PROPS_CONTENT_TYPE,
  });
  pkg.addPart(CORE_PROPS_PART_NAME, CORE_PROPS_CONTENT_TYPE, bytes);

  const rootRels = pkg.rootRels() ?? emptyRels();
  const rId = nextRelId(rootRels.items.map((r) => r.id));
  rootRels.items.push({
    id: rId,
    type: REL_TYPES.coreProperties,
    target: 'docProps/core.xml',
    targetMode: 'Internal',
  });
  pkg.setRootRels(rootRels);
};

// ---------------------------------------------------------------------------
// Extended properties (`/docProps/app.xml`).

const NS_EXT_PROPS = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
const EXT_PROPS_PART_NAME = partName('/docProps/app.xml');

/**
 * Selected string fields from `/docProps/app.xml`
 * (extended-properties / "app props"). PowerPoint exposes these
 * under File › Info / Properties as the "Origin" and "Related
 * People" groups.
 *
 * Numeric / derived fields (`Slides`, `Words`, `Paragraphs`, …) are
 * intentionally omitted — they're recomputed by PowerPoint on save
 * and reading them tends to lie about decks edited outside Office.
 */
export interface ExtendedProperties {
  readonly application: string | null;
  readonly appVersion: string | null;
  readonly company: string | null;
  readonly manager: string | null;
  readonly presentationFormat: string | null;
  readonly hyperlinkBase: string | null;
}

/**
 * Reads `/docProps/app.xml`. Returns `null` if the package has no
 * extended-properties part. Each field is `null` when the
 * corresponding element is absent or empty.
 */
export const getExtendedProperties = (pres: PresentationData): ExtendedProperties | null => {
  const pkg = pres[INTERNAL_PACKAGE];
  const part = pkg.getPart(EXT_PROPS_PART_NAME);
  if (!part) return null;
  const root = parseXml(decode(part.data)).root;
  const read = (local: string): string | null => {
    const el = firstChildElement(root, qname('', local, NS_EXT_PROPS));
    if (!el) return null;
    let s = '';
    for (const c of el.children) if (c.kind === 'text') s += c.data;
    return s.length === 0 ? null : s;
  };
  return {
    application: read('Application'),
    appVersion: read('AppVersion'),
    company: read('Company'),
    manager: read('Manager'),
    presentationFormat: read('PresentationFormat'),
    hyperlinkBase: read('HyperlinkBase'),
  };
};

const EXT_PROP_FIELDS: ReadonlyArray<{ key: keyof ExtendedProperties; local: string }> = [
  { key: 'application', local: 'Application' },
  { key: 'appVersion', local: 'AppVersion' },
  { key: 'company', local: 'Company' },
  { key: 'manager', local: 'Manager' },
  { key: 'presentationFormat', local: 'PresentationFormat' },
  { key: 'hyperlinkBase', local: 'HyperlinkBase' },
];

/**
 * Writes selected fields on `/docProps/app.xml`. Throws when the
 * package has no extended-properties part — unlike core-properties,
 * we don't bootstrap app.xml from scratch because its schema
 * requires several derived `<vt:*>` elements (`HeadingPairs`,
 * `TitlesOfParts`, …) that aren't user-facing.
 *
 * Pass `null` to clear an existing field's text. Unspecified keys
 * are left untouched.
 */
export const setExtendedProperties = (
  pres: PresentationData,
  values: Partial<ExtendedProperties>,
): void => {
  const pkg = pres[INTERNAL_PACKAGE];
  const part = pkg.getPart(EXT_PROPS_PART_NAME);
  if (!part) {
    throw new Error('setExtendedProperties: /docProps/app.xml not present; cannot bootstrap');
  }
  const doc = parseXml(decode(part.data));
  for (const field of EXT_PROP_FIELDS) {
    if (!(field.key in values)) continue;
    const value = values[field.key] ?? null;
    const name = qname('', field.local, NS_EXT_PROPS);
    const existing = firstChildElement(doc.root, name);
    if (value === null) {
      if (existing) existing.children = [];
      continue;
    }
    if (existing) {
      existing.children = [textNode(value)];
    } else {
      doc.root.children.push(elem(name, { children: [textNode(value)] }));
    }
  }
  part.data = encode(serializeXml(doc));
};

// ---------------------------------------------------------------------------
// Document thumbnail (`/docProps/thumbnail.jpeg` typically).

/**
 * The package's thumbnail image, when present. PowerPoint, the OS
 * file picker, and SharePoint preview decks via this. Format is
 * what's encoded in the thumbnail part — usually JPEG.
 */
export interface PresentationThumbnail {
  readonly format: ImageFormat;
  readonly bytes: Uint8Array;
}

const findThumbnailRel = (pkg: OpcPackage): { partName: PartName; rId: string } | null => {
  const rootRels = pkg.rootRels();
  if (!rootRels) return null;
  const rel = rootRels.items.find((r) => r.type === REL_TYPES.thumbnail);
  if (!rel) return null;
  const target = rel.target;
  const name = target.startsWith('/') ? partName(target) : partName(`/${target}`);
  return { partName: name, rId: rel.id };
};

/**
 * Returns the package's thumbnail bytes, plus the detected image
 * format. Returns `null` when the package has no thumbnail rel or
 * the rel target part is missing.
 *
 * The returned `bytes` is a live view into the thumbnail part —
 * treat it as read-only; copy if you need an independent buffer.
 */
export const getThumbnail = (pres: PresentationData): PresentationThumbnail | null => {
  const pkg = pres[INTERNAL_PACKAGE];
  const hit = findThumbnailRel(pkg);
  if (!hit) return null;
  const part = pkg.getPart(hit.partName);
  if (!part) return null;
  const format = detectImageFormat(part.data);
  if (format === null) return null;
  return { format, bytes: part.data };
};

/**
 * Replaces the package's thumbnail. Auto-detects the image format
 * from the bytes (pass `options.format` to override). Bootstraps the
 * thumbnail part + root rel + content-type default if the package
 * had no thumbnail; otherwise replaces the existing thumbnail in
 * place, switching its filename / extension if the format changed.
 */
export const setThumbnail = (
  pres: PresentationData,
  bytes: Uint8Array,
  options: { format?: ImageFormat } = {},
): void => {
  const format = options.format ?? detectImageFormat(bytes);
  if (format === null) {
    throw new Error('setThumbnail: could not detect image format. Pass options.format explicitly.');
  }
  const pkg = pres[INTERNAL_PACKAGE];
  const contentType = contentTypeForFormat(format);
  const extension = extensionForFormat(format);
  const desiredName = partName(`/docProps/thumbnail.${extension}`);

  const hit = findThumbnailRel(pkg);
  if (hit) {
    // Replace in place if the existing part is the same path; otherwise
    // remove the old one and add a new one with the right extension.
    if (hit.partName === desiredName) {
      const part = pkg.getPart(hit.partName);
      if (!part) throw new Error(`thumbnail rel points at missing part ${hit.partName}`);
      part.data = bytes;
      part.contentType = contentType;
      setOpcDefault(pkg, extension, contentType);
      return;
    }
    pkg.removePart(hit.partName);
    pkg.addPart(desiredName, contentType, bytes);
    setOpcDefault(pkg, extension, contentType);
    const rootRels = pkg.rootRels() ?? emptyRels();
    const existing = rootRels.items.find((r) => r.id === hit.rId);
    if (existing) {
      existing.target = `docProps/thumbnail.${extension}`;
    }
    pkg.setRootRels(rootRels);
    return;
  }

  // Bootstrap.
  setOpcDefault(pkg, extension, contentType);
  pkg.addPart(desiredName, contentType, bytes);
  const rootRels = pkg.rootRels() ?? emptyRels();
  const rId = nextRelId(rootRels.items.map((r) => r.id));
  rootRels.items.push({
    id: rId,
    type: REL_TYPES.thumbnail,
    target: `docProps/thumbnail.${extension}`,
    targetMode: 'Internal',
  });
  pkg.setRootRels(rootRels);
};

/**
 * Removes the package's thumbnail entirely (drops the rel + the
 * underlying part). No-op when the package has no thumbnail.
 */
export const removeThumbnail = (pres: PresentationData): void => {
  const pkg = pres[INTERNAL_PACKAGE];
  const hit = findThumbnailRel(pkg);
  if (!hit) return;
  pkg.removePart(hit.partName);
  const rootRels = pkg.rootRels();
  if (!rootRels) return;
  rootRels.items = rootRels.items.filter((r) => r.id !== hit.rId);
  pkg.setRootRels(rootRels);
};

/**
 * Finds the first slide layout whose user-visible name matches `name`,
 * or `null` if none does. Convenience over `getSlideLayouts(...).find(...)`.
 */
export const findSlideLayout = (pres: PresentationData, name: string): SlideLayoutData | null => {
  for (const layout of getSlideLayouts(pres)) {
    if (layout[LAYOUT_PART].name === name) return layout;
  }
  return null;
};

/**
 * Returns every slide layout in the package that exposes a
 * placeholder of the given type token (`'title'`, `'body'`,
 * `'ftr'`, etc.). Useful for "find every layout that can host a
 * body" lookups before `addSlide`.
 */
export const findLayoutsWithPlaceholderType = (
  pres: PresentationData,
  type: string,
): ReadonlyArray<SlideLayoutData> => {
  const out: SlideLayoutData[] = [];
  for (const layout of getSlideLayouts(pres)) {
    const phs = getSlideLayoutPlaceholders(layout);
    const hit = phs.some(
      (p) => p.type === type || (type === 'body' && p.type === null && p.idx !== null),
    );
    if (hit) out.push(layout);
  }
  return out;
};

/**
 * Finds the first slide layout with the given `<p:sldLayout type="...">`
 * token. Unlike `findSlideLayout` (which matches the user-visible
 * name, and is therefore locale-sensitive), this matches the spec
 * token — `title`, `obj`, `twoObj`, `blank`, etc. — and is stable
 * across PowerPoint UI languages.
 */
export const findSlideLayoutByType = (
  pres: PresentationData,
  layoutType: SlideLayoutType | string,
): SlideLayoutData | null => {
  for (const layout of getSlideLayouts(pres)) {
    if (layout[LAYOUT_PART].layoutType === layoutType) return layout;
  }
  return null;
};

/**
 * Layout type token, when present (`title`, `obj`, `twoObj`, ...).
 * `null` when omitted — the spec default for that case is `cust`.
 */
export const getSlideLayoutType = (layout: SlideLayoutData): SlideLayoutType | string | null =>
  layout[LAYOUT_PART].layoutType;

/**
 * Enumerates every slide layout in the package.
 */
export const getSlideLayouts = (pres: PresentationData): ReadonlyArray<SlideLayoutData> => {
  const pkg = pres[INTERNAL_PACKAGE];
  const out: SlideLayoutData[] = [];
  for (const part of pkg.parts) {
    if (part.contentType !== SLIDE_LAYOUT_CONTENT_TYPE) continue;
    const root = parseXml(decode(part.data)).root;
    out.push({
      [LAYOUT_PART_NAME]: part.name,
      [LAYOUT_PART]: readSlideLayoutPart(root),
    });
  }
  return out;
};

// ---------------------------------------------------------------------------
// SlideData factory + cached enumeration.

const buildSlideData = (pkg: OpcPackage, partNameValue: PartName, bytes: Uint8Array): SlideData => {
  const doc = parseXml(decode(bytes));
  const part = readSlidePart(doc.root);
  const shapes: SlideShapeData[] = [];
  const slide: SlideData = {
    [INTERNAL_PACKAGE]: pkg,
    [SLIDE_PART_NAME]: partNameValue,
    [SLIDE_DOCUMENT]: doc,
    [SLIDE_PART]: part,
    [SLIDE_SHAPES]: shapes,
  };
  for (const snap of part.shapes) {
    shapes.push({
      [SHAPE_SLIDE]: slide,
      [SHAPE_ELEMENT]: snap.element,
      [SHAPE_SNAPSHOT]: snap,
    });
  }
  return slide;
};

/**
 * Enumerates slides in presentation order. Returns opaque `SlideData`
 * handles that pass to every slide-level fn-API helper.
 *
 * Throws if any referenced slide part is missing — a structurally
 * invalid PPTX cannot honor the L1 contract.
 */
/**
 * Returns the shape at the given 0-based `index` on the slide, or
 * `null` when `index` is out of range. Convenience over
 * `getSlideShapes(slide)[index] ?? null`.
 */
export const getShapeAt = (slide: SlideData, index: number): SlideShapeData | null =>
  slide[SLIDE_SHAPES][index] ?? null;

/**
 * Returns the slide count without forcing every slide part to be
 * parsed. Reads only `presentation.xml` and walks `<p:sldIdLst>`.
 *
 * Equivalent to `getSlides(pres).length`, but cheaper on cold reads
 * — useful when the caller only wants a count for UI badges or
 * validation logic.
 */
export const getSlideCount = (pres: PresentationData): number => {
  const cached = pres._slidesCache;
  if (cached !== null) return cached.length;
  const pkg = pres[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (!presPart) return 0;
  const root = parseXml(decode(presPart.data)).root;
  const model = readPresentationPart(root);
  return model.slides.length;
};

/**
 * Returns the number of slide layouts in the package. Walks the
 * content-type map; cheaper than `getSlideLayouts(pres).length`
 * because it avoids parsing each layout's XML body.
 */
export const getSlideLayoutCount = (pres: PresentationData): number => {
  const pkg = pres[INTERNAL_PACKAGE];
  let n = 0;
  for (const part of pkg.parts) {
    if (part.contentType === SLIDE_LAYOUT_CONTENT_TYPE) n++;
  }
  return n;
};

export const getSlides = (pres: PresentationData): ReadonlyArray<SlideData> => {
  const cached = pres._slidesCache;
  if (cached !== null) return cached as ReadonlyArray<SlideData>;

  const pkg = pres[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (presPart === null) {
    const empty: SlideData[] = [];
    pres._slidesCache = empty;
    return empty;
  }
  const presRels = pkg.getRels(PRES_PART_NAME);
  if (presRels === null) {
    const empty: SlideData[] = [];
    pres._slidesCache = empty;
    return empty;
  }

  const presRoot = parseXml(decode(presPart.data)).root;
  const presModel = readPresentationPart(presRoot);

  const out: SlideData[] = [];
  for (const sld of presModel.slides) {
    const rel = presRels.items.find((r) => r.id === sld.rId);
    if (!rel) throw new Error(`presentation.xml.rels missing entry for ${sld.rId}`);
    const target = rel.target;
    const slideName = partName(target.startsWith('/') ? target : `/ppt/${target}`);
    const slidePart = pkg.getPart(slideName);
    if (slidePart === null) throw new Error(`slide part ${slideName} not found`);
    out.push(buildSlideData(pkg, slideName, slidePart.data));
  }
  pres._slidesCache = out;
  return out;
};

/**
 * Concatenated visible text of a slide. Convenience wrapper that walks
 * the slide's shape tree without instantiating any class.
 */
export const getSlideText = (slide: SlideData): string => slideText(slide[SLIDE_PART]);

/**
 * Length (in code points) of the slide's concatenated visible text.
 * Counts Unicode code points via `Array.from`, so surrogate-pair
 * characters (emoji, supplementary CJK) count as 1, matching the
 * library's invariant on text editing.
 */
export const getSlideTextLength = (slide: SlideData): number =>
  Array.from(getSlideText(slide)).length;

/**
 * A single slide's outline entry: its index, title (or `null`), and
 * the text of its body placeholder (or `null`). Used by
 * `getSlideOutline`.
 */
export interface SlideOutlineEntry {
  readonly index: number;
  readonly title: string | null;
  readonly body: string | null;
}

/**
 * Returns a one-entry-per-slide outline: title + body-placeholder
 * text. Useful for building thumbnail panes, table-of-contents
 * inserts, or quick-look previews without rendering the slide.
 *
 * "Body" here is the first matching `body` placeholder per
 * `findSlidePlaceholder(slide, 'body')`. Non-placeholder text on
 * the slide isn't surfaced; use `getSlideText(slide)` for that.
 */
export const getSlideOutline = (pres: PresentationData): ReadonlyArray<SlideOutlineEntry> => {
  const out: SlideOutlineEntry[] = [];
  const slides = getSlides(pres);
  for (let i = 0; i < slides.length; i++) {
    const slide = slides[i]!;
    const body = findSlidePlaceholder(slide, 'body');
    out.push({
      index: i,
      title: getSlideTitle(slide),
      body: body !== null ? body[SHAPE_SNAPSHOT].text : null,
    });
  }
  return out;
};

/**
 * Concatenated title + body text from every slide, joined with the
 * given `separator` (defaults to `'\n\n'`). Useful for generating
 * a table-of-contents handout from a deck. Slides without a title
 * still contribute their body (and vice versa); slides without
 * both are skipped.
 */
export const getOutlineText = (pres: PresentationData, separator: string = '\n\n'): string => {
  const parts: string[] = [];
  for (const entry of getSlideOutline(pres)) {
    const segments: string[] = [];
    if (entry.title !== null) segments.push(entry.title);
    if (entry.body !== null) segments.push(entry.body);
    if (segments.length > 0) parts.push(segments.join('\n'));
  }
  return parts.join(separator);
};

/**
 * Concatenated visible text from every slide, joined with the
 * given `separator` (defaults to a form-feed, `\f`, between slides).
 * Useful for search-indexing a whole deck without iterating slides
 * yourself.
 */
export const getPresentationText = (pres: PresentationData, separator: string = '\f'): string => {
  const parts: string[] = [];
  for (const slide of getSlides(pres)) parts.push(slideText(slide[SLIDE_PART]));
  return parts.join(separator);
};

/**
 * Total code-point length of visible text across every slide.
 * Sibling of `getSlideTextLength`; counts surrogate-pair characters
 * (emoji, supplementary CJK) as 1 each. Cheaper than building the
 * concatenated string when only the length matters.
 */
export const getPresentationTextLength = (pres: PresentationData): number => {
  let n = 0;
  for (const slide of getSlides(pres)) {
    n += Array.from(slideText(slide[SLIDE_PART])).length;
  }
  return n;
};

/**
 * Returns the 0-based index of `slide` within `pres`, or `-1` if the
 * slide doesn't belong to this presentation (e.g. after a removeSlide
 * call, or if it was constructed from a different package).
 */
export const getSlideIndex = (pres: PresentationData, slide: SlideData): number => {
  const slides = getSlides(pres);
  return slides.indexOf(slide);
};

/**
 * Returns the slide at the given 0-based `index`, or `null` if `index`
 * is out of range. Convenience over `getSlides(pres)[index] ?? null`.
 */
export const getSlideAt = (pres: PresentationData, index: number): SlideData | null => {
  const slides = getSlides(pres);
  return slides[index] ?? null;
};

/**
 * Per-slide summary suitable for slide pickers or audit reports.
 *
 * `index` is the 0-based position in the deck. `title` falls back to
 * `null` when the slide has no title placeholder. `layoutName` falls
 * back to `null` when the slide has no `slideLayout` rel or the
 * layout doesn't carry a user-visible name.
 */
export interface SlideInfo {
  readonly index: number;
  readonly title: string | null;
  readonly hidden: boolean;
  readonly shapeCount: number;
  readonly layoutName: string | null;
}

export const getSlideInfo = (pres: PresentationData, slide: SlideData): SlideInfo => {
  const layout = getSlideLayout(slide);
  return {
    index: getSlideIndex(pres, slide),
    title: getSlideTitle(slide),
    hidden: isSlideHidden(slide),
    shapeCount: getSlideShapes(slide).length,
    layoutName: layout ? layout[LAYOUT_PART].name : null,
  };
};

/**
 * Returns every slide bound to the given `layout`. Useful for "find
 * all slides using the `Title and Content` layout" workflows; pair
 * with `findSlideLayout(pres, name)` to look up the layout by name.
 */
export const getSlidesByLayout = (
  pres: PresentationData,
  layout: SlideLayoutData,
): ReadonlyArray<SlideData> => {
  const target = layout[LAYOUT_PART_NAME];
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    const found = getSlideLayout(slide);
    if (found !== null && found[LAYOUT_PART_NAME] === target) out.push(slide);
  }
  return out;
};

/**
 * Returns the first slide whose title-placeholder text equals
 * `title` exactly. Returns `null` when no slide matches. Different
 * from `findSlideByText`, which matches any visible text via
 * substring / regex — this one is a strict equality match against
 * the title placeholder only.
 */
export const findSlideByTitle = (pres: PresentationData, title: string): SlideData | null => {
  for (const slide of getSlides(pres)) {
    if (getSlideTitle(slide) === title) return slide;
  }
  return null;
};

/**
 * Returns the slide whose package part name equals `partName`
 * (typically `/ppt/slides/slideN.xml`), or `null` when no such
 * slide exists. Useful for callers that walk validator output,
 * `listPackageParts`, or any other low-level path-keyed API and
 * need the matching `SlideData` handle.
 */
export const findSlideByPartName = (pres: PresentationData, partName: string): SlideData | null => {
  for (const slide of getSlides(pres)) {
    if (slide[SLIDE_PART_NAME] === partName) return slide;
  }
  return null;
};

/**
 * Returns the package part name (e.g. `/ppt/slides/slide3.xml`) of
 * `slide`. Useful for surfacing slides in validator output, error
 * messages, or any path-keyed sidebar UI.
 */
export const getSlidePartName = (slide: SlideData): string => slide[SLIDE_PART_NAME];

/**
 * Returns the slide's current XML body as a string. Re-serializes
 * from the typed AST, so this reflects any pending edits that
 * haven't been written back to the package yet.
 *
 * Intended for diagnostics: dumping into bug reports, asserting
 * structure in tests, snapshotting before / after a transformation.
 * Do NOT parse this back yourself — round-trip safely through
 * `loadPresentation` / `savePresentation`.
 */
export const getSlideXmlString = (slide: SlideData): string => serializeXml(slide[SLIDE_DOCUMENT]);

/**
 * Returns the first slide whose concatenated visible text contains
 * `needle` (substring; case-sensitive). Pass a `RegExp` to test
 * against the slide's text body instead. Returns `null` when no
 * slide matches.
 */
export const findSlideByText = (
  pres: PresentationData,
  needle: string | RegExp,
): SlideData | null => {
  for (const slide of getSlides(pres)) {
    const text = slideText(slide[SLIDE_PART]);
    if (typeof needle === 'string' ? text.includes(needle) : needle.test(text)) {
      return slide;
    }
  }
  return null;
};

/**
 * Every slide whose concatenated visible text contains `needle`
 * (substring) or matches the given `RegExp`.
 */
export const findSlidesByText = (
  pres: PresentationData,
  needle: string | RegExp,
): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    const text = slideText(slide[SLIDE_PART]);
    if (typeof needle === 'string' ? text.includes(needle) : needle.test(text)) {
      out.push(slide);
    }
  }
  return out;
};

/**
 * Every slide whose speaker notes contain `needle` (substring) or
 * match the given `RegExp`. Sibling of `findSlidesByText` — useful
 * for surfacing reviewer-facing annotations stored in notes (e.g.
 * "every slide whose notes say TODO").
 *
 * Slides without notes are skipped.
 */
export const findSlidesByNotes = (
  pres: PresentationData,
  needle: string | RegExp,
): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    const notes = getSlideNotes(slide);
    if (notes === null || notes.length === 0) continue;
    if (typeof needle === 'string' ? notes.includes(needle) : needle.test(notes)) {
      out.push(slide);
    }
  }
  return out;
};

/**
 * Returns every slide where `needle` appears in **either** the
 * visible text or the speaker notes. Each slide is reported once,
 * in document order. Sibling of `findSlidesByText` /
 * `findSlidesByNotes` — useful for a generic "search the deck"
 * UX where the caller doesn't care which surface matched.
 */
export const searchSlides = (
  pres: PresentationData,
  needle: string | RegExp,
): ReadonlyArray<SlideData> => {
  const matchesString = (haystack: string): boolean =>
    typeof needle === 'string' ? haystack.includes(needle) : needle.test(haystack);
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    if (matchesString(slideText(slide[SLIDE_PART]))) {
      out.push(slide);
      continue;
    }
    const notes = getSlideNotes(slide);
    if (notes !== null && notes.length > 0 && matchesString(notes)) out.push(slide);
  }
  return out;
};

/**
 * Replaces every occurrence of `from` in every slide's speaker
 * notes with `to`. Returns the number of slides updated. Sibling
 * of `replaceTextInPresentation` for the reviewer-annotation pass.
 *
 * Slides without notes are skipped. For string needles, replacement
 * is global; for RegExp needles, the global flag is forced.
 */
export const replaceTextInNotes = (
  pres: PresentationData,
  from: string | RegExp,
  to: string,
): number => {
  let n = 0;
  for (const slide of getSlides(pres)) {
    if (replaceTextInSlideNotes(slide, from, to)) n++;
  }
  return n;
};

/**
 * Slide-scoped sibling of `replaceTextInNotes`. Replaces every
 * occurrence of `from` in the slide's speaker notes with `to`.
 * Returns `true` if the notes changed, `false` otherwise (no notes
 * present, or no match).
 */
export const replaceTextInSlideNotes = (
  slide: SlideData,
  from: string | RegExp,
  to: string,
): boolean => {
  const notes = getSlideNotes(slide);
  if (notes === null || notes.length === 0) return false;
  let next: string;
  if (typeof from === 'string') {
    if (!notes.includes(from)) return false;
    next = notes.split(from).join(to);
  } else {
    const re = from.global ? from : new RegExp(from.source, `${from.flags}g`);
    if (!re.test(notes)) return false;
    re.lastIndex = 0;
    next = notes.replace(re, to);
  }
  if (next === notes) return false;
  setSlideNotes(slide, next);
  return true;
};

/**
 * Every shape on every slide, paired with its slide and the slide's
 * 0-based index. Useful for cross-deck audits ("find every picture",
 * "list shape names anywhere") without writing the nested-loop
 * boilerplate yourself.
 */
export interface AllShapesEntry {
  readonly slide: SlideData;
  readonly slideIndex: number;
  readonly shape: SlideShapeData;
}

export const getAllShapes = (pres: PresentationData): ReadonlyArray<AllShapesEntry> => {
  const out: AllShapesEntry[] = [];
  const slides = getSlides(pres);
  for (let i = 0; i < slides.length; i++) {
    const slide = slides[i]!;
    for (const shape of getSlideShapes(slide)) {
      out.push({ slide, slideIndex: i, shape });
    }
  }
  return out;
};

// ---------------------------------------------------------------------------
// Slide visibility — `<p:sld show="0">` hides a slide from the slideshow
// without removing it from the deck. `show="1"` (or omission) is visible.

const ATTR_SHOW = qname('', 'show', '');

/**
 * Returns `true` when the slide carries `show="0"` on its root
 * `<p:sld>` element. PowerPoint hides such slides from the slideshow
 * but keeps them in the editing surface.
 */
export const isSlideHidden = (slide: SlideData): boolean => {
  const show = getAttrValue(slide[SLIDE_DOCUMENT].root, ATTR_SHOW);
  return show === '0';
};

/**
 * Toggles the slide's visibility in the slideshow. Hiding adds
 * `show="0"`; showing removes the attribute (PowerPoint treats absence
 * as the default `show="1"`).
 */
export const setSlideHidden = (slide: SlideData, hidden: boolean): void => {
  const root = slide[SLIDE_DOCUMENT].root;
  root.attrs = root.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'show'),
  );
  if (hidden) root.attrs.push(attr(ATTR_SHOW, '0'));
  commitSlideData(slide);
};

/**
 * Replaces `{{key}}` tokens on every slide. Returns the total number of
 * substitutions performed.
 */
export const replaceTokensInPresentation = (
  pres: PresentationData,
  tokens: Record<string, string>,
): number => {
  const pkg = pres[INTERNAL_PACKAGE];
  let count = 0;
  for (const part of pkg.parts) {
    if (part.contentType !== SLIDE_CONTENT_TYPE) continue;
    const doc = parseXml(decode(part.data));
    const n = replaceTokensInTree(doc.root, tokens);
    if (n > 0) {
      part.data = encode(serializeXml(doc));
      count += n;
    }
  }
  pres._slidesCache = null;
  return count;
};

/**
 * Replaces every occurrence of `from` in every slide's text with `to`.
 * `from` may be a string (treated as a literal) or a `RegExp`. Returns
 * the number of `<a:t>` elements mutated across the whole deck.
 *
 * Use this for the broad "rename product X to Y" pattern; for
 * `{{token}}` style substitutions, prefer
 * `replaceTokensInPresentation` which is more explicit about intent.
 */
export const replaceTextInPresentation = (
  pres: PresentationData,
  from: string | RegExp,
  to: string,
): number => {
  const pkg = pres[INTERNAL_PACKAGE];
  let count = 0;
  for (const part of pkg.parts) {
    if (part.contentType !== SLIDE_CONTENT_TYPE) continue;
    const doc = parseXml(decode(part.data));
    const n = replaceTextInTree(doc.root, from, to);
    if (n > 0) {
      part.data = encode(serializeXml(doc));
      count += n;
    }
  }
  pres._slidesCache = null;
  return count;
};

/**
 * Replaces every occurrence of `from` in the slide's text with `to`.
 * Returns the number of `<a:t>` elements mutated on this slide.
 */
export const replaceTextInSlide = (slide: SlideData, from: string | RegExp, to: string): number => {
  const n = replaceTextInTree(slide[SLIDE_DOCUMENT].root, from, to);
  if (n > 0) {
    commitSlideData(slide);
    refreshSlideData(slide);
  }
  return n;
};

// ---------------------------------------------------------------------------
// Deck manipulation.

const ensureSldIdLst = (presentationRoot: XmlElement): XmlElement => {
  const existing = firstChildElement(presentationRoot, NAME_SLD_ID_LST);
  if (existing !== null) return existing;
  const fresh = elem(NAME_SLD_ID_LST);
  const masterLst = firstChildElement(presentationRoot, NAME_SLD_MASTER_ID_LST);
  if (masterLst === null) {
    presentationRoot.children.unshift(fresh);
    return fresh;
  }
  const idx = presentationRoot.children.indexOf(masterLst);
  presentationRoot.children.splice(idx + 1, 0, fresh);
  return fresh;
};

const allocateSldId = (sldIdLst: XmlElement): number => {
  let max = SLD_ID_MIN - 1;
  for (const sldId of allChildElements(sldIdLst, NAME_SLD_ID)) {
    const raw = getAttrValue(sldId, ATTR_ID);
    if (raw === null) continue;
    const n = Number.parseInt(raw, 10);
    if (Number.isFinite(n) && n > max) max = n;
  }
  const next = Math.max(SLD_ID_MIN, max + 1);
  if (next > SLD_ID_MAX) {
    throw new Error(`sldId allocator exhausted (next would be ${next}, max ${SLD_ID_MAX})`);
  }
  return next;
};

const allocateSlideN = (pkg: OpcPackage): number => {
  let next = 1;
  for (const p of pkg.parts) {
    const m = p.name.match(/^\/ppt\/slides\/slide(\d+)\.xml$/);
    if (m?.[1] !== undefined) {
      const n = Number.parseInt(m[1], 10);
      if (Number.isFinite(n) && n >= next) next = n + 1;
    }
  }
  return next;
};

const requirePresentationDoc = (pkg: OpcPackage): XmlDocument => {
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (!presPart) throw new Error('presentation.xml is missing');
  const doc = parseXml(decode(presPart.data));
  if (
    doc.root.name.namespaceURI !== NAME_PRESENTATION.namespaceURI ||
    doc.root.name.localName !== 'presentation'
  ) {
    throw new Error('presentation.xml root is not <p:presentation>');
  }
  return doc;
};

/**
 * Adds a new slide bound to `layout`. Returns the new `SlideData`.
 *
 * Allocates a fresh part name, sldId, and rId; clones layout
 * placeholders into the slide; writes `[Content_Types].xml`, the
 * slide's `.rels`, presentation's `.rels`, and `<p:sldIdLst>`. The
 * deck-cache on `pres` is invalidated so the next `getSlides` call
 * sees the new entry.
 */
/**
 * Convenience over `addSlide` that picks a layout automatically:
 *
 *   1. The layout with `<p:sldLayout type="blank">`, if present.
 *   2. Otherwise, the first available layout (alphabetical by
 *      part name).
 *
 * Throws when the package carries no layouts at all (which would
 * be a structurally-broken deck).
 */
export const addBlankSlide = (pres: PresentationData): SlideData => {
  const blank = findSlideLayoutByType(pres, 'blank');
  if (blank) return addSlide(pres, { layout: blank });
  const layouts = getSlideLayouts(pres);
  if (layouts.length === 0) {
    throw new Error('addBlankSlide: package has no slide layouts to inherit from');
  }
  return addSlide(pres, { layout: layouts[0]! });
};

/**
 * Sugar over `addSlide` + `setSlideTitle` + `setSlideBody` for the
 * "title + body" pattern. Picks the `obj` (Title and Content)
 * layout when present, falling back to the first layout with a
 * body placeholder.
 *
 * Throws if no layout in the package offers a body slot.
 */
export const addContentSlide = (
  pres: PresentationData,
  opts: { title?: string; body?: string },
): SlideData => {
  const objLayout = findSlideLayoutByType(pres, 'obj');
  const layout =
    objLayout ??
    getSlideLayouts(pres).find((l) =>
      getSlideLayoutPlaceholders(l).some((p) => p.type === null || p.type === 'body'),
    );
  if (!layout) {
    throw new Error('addContentSlide: no layout with a body placeholder found');
  }
  const slide = addSlide(pres, { layout });
  if (opts.title !== undefined) setSlideTitle(slide, opts.title);
  if (opts.body !== undefined) setSlideBody(slide, opts.body);
  return slide;
};

/**
 * Sugar over `addSlide` + `setSlideTitle` for the section-divider
 * pattern. Picks `<p:sldLayout type="secHead">` when present (the
 * PowerPoint "Section Header" layout); otherwise falls back to a
 * `title`-typed layout or the first available layout.
 */
export const addSectionHeaderSlide = (pres: PresentationData, title: string): SlideData => {
  const layout =
    findSlideLayoutByType(pres, 'secHead') ??
    findSlideLayoutByType(pres, 'title') ??
    getSlideLayouts(pres)[0];
  if (!layout) {
    throw new Error('addSectionHeaderSlide: package has no slide layouts to inherit from');
  }
  const slide = addSlide(pres, { layout });
  setSlideTitle(slide, title);
  return slide;
};

/**
 * Sugar over `addSlide` + `setSlideTitle` for the common
 * "title slide + set heading" pattern. Picks the `title` layout
 * first, then falls back to the first non-blank layout.
 *
 * Throws when the package carries no layouts at all.
 */
export const addTitleSlide = (pres: PresentationData, title: string): SlideData => {
  const titleLayout =
    findSlideLayoutByType(pres, 'title') ?? findSlideLayoutByType(pres, 'obj') ?? null;
  const layout =
    titleLayout ??
    getSlideLayouts(pres).find((l) => getSlideLayoutType(l) !== 'blank') ??
    getSlideLayouts(pres)[0];
  if (!layout) {
    throw new Error('addTitleSlide: package has no slide layouts to inherit from');
  }
  const slide = addSlide(pres, { layout });
  setSlideTitle(slide, title);
  return slide;
};

export const addSlide = (
  pres: PresentationData,
  options: { layout: SlideLayoutData },
): SlideData => {
  const pkg = pres[INTERNAL_PACKAGE];
  const layout = options.layout;
  const layoutPart = layout[LAYOUT_PART];
  const layoutPartName = layout[LAYOUT_PART_NAME];

  const presDoc = requirePresentationDoc(pkg);
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (!presPart) throw new Error('presentation.xml is missing');

  const sldIdLst = ensureSldIdLst(presDoc.root);
  const newSldId = allocateSldId(sldIdLst);
  const slideN = allocateSlideN(pkg);
  const newSlidePartName = partName(`/ppt/slides/slide${slideN}.xml`);

  const layoutCsld = firstChildElement(layoutPart.root, NAME_CSLD);
  if (!layoutCsld) throw new Error(`layout ${layoutPartName} missing <p:cSld>`);
  const layoutSpTree = firstChildElement(layoutCsld, NAME_SP_TREE);
  if (!layoutSpTree) throw new Error(`layout ${layoutPartName} missing <p:spTree>`);

  const slideDoc = buildSlideFromLayout(layoutSpTree);
  const slideBytes = encode(serializeXml(slideDoc));
  pkg.addPart(newSlidePartName, SLIDE_CONTENT_TYPE, slideBytes);

  const slideRels = emptyRels();
  slideRels.items.push({
    id: 'rId1',
    type: REL_TYPES.slideLayout,
    target: `../slideLayouts/${basename(layoutPartName)}`,
    targetMode: 'Internal',
  });
  pkg.setRels(newSlidePartName, slideRels);

  const presRels = pkg.getRels(PRES_PART_NAME) ?? emptyRels();
  const newRId = nextRelId(presRels.items.map((r) => r.id));
  presRels.items.push({
    id: newRId,
    type: REL_TYPES.slide,
    target: `slides/slide${slideN}.xml`,
    targetMode: 'Internal',
  });
  pkg.setRels(PRES_PART_NAME, presRels);

  sldIdLst.children.push(
    elem(NAME_SLD_ID, {
      attrs: [attr(ATTR_ID, String(newSldId)), attr(ATTR_R_ID, newRId)],
    }),
  );
  presPart.data = encode(serializeXml(presDoc));

  pres._slidesCache = null;
  const slides = getSlides(pres);
  const last = slides[slides.length - 1];
  if (!last) throw new Error('addSlide: post-condition failed; slide not in cache');
  return last;
};

/**
 * Removes the given slide from the deck. Removes the `<p:sldId>`, the
 * `presentation.xml.rels` entry, and the slide part + its `.rels` part.
 *
 * Media parts are intentionally NOT cleaned up — they may be shared
 * with other slides. The freed `sldId` is NOT reused on subsequent
 * `addSlide` calls (PowerPoint quirk, see plan §Risks).
 */
export const removeSlide = (pres: PresentationData, slide: SlideData): void => {
  const pkg = pres[INTERNAL_PACKAGE];
  const slidePartName = slide[SLIDE_PART_NAME];
  if (pkg.getPart(slidePartName) === null) {
    throw new Error(`removeSlide: ${slidePartName} not present in package`);
  }

  const presRels = pkg.getRels(PRES_PART_NAME);
  if (!presRels) throw new Error('presentation.xml has no rels');
  const slideTargetRel = `slides/${basename(slidePartName)}`;
  const removedRel = presRels.items.find(
    (r) => r.type === REL_TYPES.slide && r.target === slideTargetRel,
  );
  if (!removedRel) {
    throw new Error(`presentation.xml.rels missing entry for slide ${slidePartName}`);
  }
  presRels.items = presRels.items.filter((r) => r.id !== removedRel.id);
  pkg.setRels(PRES_PART_NAME, presRels);

  const presPart = pkg.getPart(PRES_PART_NAME);
  if (!presPart) throw new Error('presentation.xml missing');
  const presDoc = parseXml(decode(presPart.data));
  const sldIdLst = firstChildElement(presDoc.root, NAME_SLD_ID_LST);
  if (sldIdLst !== null) {
    sldIdLst.children = sldIdLst.children.filter((c) => {
      if (c.kind !== 'element') return true;
      if (c.name.namespaceURI !== NS.pml || c.name.localName !== 'sldId') return true;
      return getAttrValue(c, ATTR_R_ID) !== removedRel.id;
    });
  }
  presPart.data = encode(serializeXml(presDoc));

  pkg.removePart(relsPartNameFor(slidePartName));
  pkg.removePart(slidePartName);
  pres._slidesCache = null;
};

/**
 * Reorders every slide in the deck via a custom comparator. The
 * comparator is invoked with two `SlideData` handles and returns the
 * usual `Array.prototype.sort` ordering (-1 / 0 / 1).
 *
 *   sortSlides(pres, (a, b) => getSlideTitle(a)?.localeCompare(getSlideTitle(b) ?? '') ?? 0);
 *
 * Internally walks `<p:sldIdLst>` and re-emits its `<p:sldId>` children
 * in the new order. Slide parts and rels are untouched — only the
 * order in which PowerPoint plays them changes.
 */
export const sortSlides = (
  pres: PresentationData,
  compareFn: (a: SlideData, b: SlideData) => number,
): void => {
  const pkg = pres[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (!presPart) throw new Error('presentation.xml missing');
  const doc = parseXml(decode(presPart.data));
  const sldIdLst = firstChildElement(doc.root, NAME_SLD_ID_LST);
  if (!sldIdLst) return; // nothing to reorder

  const slides = getSlides(pres);
  const presRels = pkg.getRels(PRES_PART_NAME);
  if (!presRels) return;

  // Build a map from rId → SlideData and from rId → its <p:sldId> element.
  const slideByRId = new Map<string, SlideData>();
  for (const slide of slides) {
    const rel = presRels.items.find(
      (r) =>
        r.type === REL_TYPES.slide && r.target === `slides/${basename(slide[SLIDE_PART_NAME])}`,
    );
    if (rel) slideByRId.set(rel.id, slide);
  }
  const sldIdElements = sldIdLst.children.filter(
    (c): c is XmlElement =>
      c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'sldId',
  );
  const sortedSlides = [...slides].sort(compareFn);
  const newOrder: XmlElement[] = [];
  for (const slide of sortedSlides) {
    let matchedRId: string | undefined;
    for (const [rId, s] of slideByRId.entries()) {
      if (s === slide) {
        matchedRId = rId;
        break;
      }
    }
    if (matchedRId === undefined) continue;
    const el = sldIdElements.find((e) => getAttrValue(e, ATTR_R_ID) === matchedRId);
    if (el) newOrder.push(el);
  }

  // Replace the children, preserving any non-sldId children (whitespace
  // or comments — unlikely but defensive).
  const nonSldIdChildren = sldIdLst.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'sldId'),
  );
  sldIdLst.children = [...nonSldIdChildren, ...newOrder];
  presPart.data = encode(serializeXml(doc));
  pres._slidesCache = null;
};

/**
 * Reverses the slide order across the whole deck. Built on
 * `sortSlides` for predictable rels behavior.
 */
export const reverseSlides = (pres: PresentationData): void => {
  const indexBy = new Map<SlideData, number>();
  for (const [i, slide] of getSlides(pres).entries()) indexBy.set(slide, i);
  sortSlides(pres, (a, b) => (indexBy.get(b) ?? 0) - (indexBy.get(a) ?? 0));
};

/**
 * Swaps the positions of the slides at `indexA` and `indexB`.
 * No-op when the indices are equal. Throws on out-of-range indices.
 * Implemented on top of `moveSlide` for predictable rels behavior.
 */
export const swapSlides = (pres: PresentationData, indexA: number, indexB: number): void => {
  if (indexA === indexB) return;
  const slides = getSlides(pres);
  const a = slides[indexA];
  const b = slides[indexB];
  if (!a) throw new RangeError(`swapSlides: indexA ${indexA} out of range (have ${slides.length})`);
  if (!b) throw new RangeError(`swapSlides: indexB ${indexB} out of range (have ${slides.length})`);
  // Move the lower-index slide to the higher index first so the
  // remaining slide stays at its original index.
  const [lo, hi] = indexA < indexB ? [indexA, indexB] : [indexB, indexA];
  moveSlide(pres, slides[lo]!, hi);
  // After the first move, the slide originally at hi is now at hi-1.
  const refreshed = getSlides(pres);
  moveSlide(pres, refreshed[hi - 1]!, lo);
};

/**
 * Reorders a slide. The slide's identity (part, rels, sldId) is
 * unchanged — only `<p:sldIdLst>`'s child order changes.
 */
export const moveSlide = (pres: PresentationData, slide: SlideData, toIndex: number): void => {
  const pkg = pres[INTERNAL_PACKAGE];
  const slideRelTarget = `slides/${basename(slide[SLIDE_PART_NAME])}`;
  const presRels = pkg.getRels(PRES_PART_NAME);
  if (!presRels) throw new Error('presentation.xml has no rels');
  const slideRel = presRels.items.find(
    (r) => r.type === REL_TYPES.slide && r.target === slideRelTarget,
  );
  if (!slideRel) throw new Error(`moveSlide: slide ${slide[SLIDE_PART_NAME]} has no rel`);

  const presPart = pkg.getPart(PRES_PART_NAME);
  if (!presPart) throw new Error('presentation.xml missing');
  const presDoc = parseXml(decode(presPart.data));
  const sldIdLst = firstChildElement(presDoc.root, NAME_SLD_ID_LST);
  if (!sldIdLst) throw new Error('presentation.xml has no <p:sldIdLst>');

  const sldIdElements = sldIdLst.children.filter(
    (c): c is XmlElement =>
      c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'sldId',
  );
  const target = sldIdElements.find((e) => getAttrValue(e, ATTR_R_ID) === slideRel.id);
  if (!target) throw new Error(`moveSlide: <p:sldId> for ${slideRel.id} not found`);

  const remaining = sldIdLst.children.filter((c) => c !== target);
  const remainingSldIds = remaining.filter(
    (c): c is XmlElement =>
      c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'sldId',
  );
  const clamped = Math.max(0, Math.min(toIndex, remainingSldIds.length));
  if (clamped === remainingSldIds.length) {
    remaining.push(target);
  } else {
    const before = remainingSldIds[clamped];
    const insertAt = before === undefined ? remaining.length : remaining.indexOf(before);
    remaining.splice(insertAt, 0, target);
  }
  sldIdLst.children = remaining;
  presPart.data = encode(serializeXml(presDoc));
  pres._slidesCache = null;
};

/**
 * Duplicates a slide. Returns the new `SlideData` appended to deck order.
 *
 * Part bytes and rels are cloned verbatim; media parts are NOT copied —
 * the duplicate shares the original's media references (PowerPoint
 * does the same).
 */
export const duplicateSlide = (pres: PresentationData, slide: SlideData): SlideData => {
  const pkg = pres[INTERNAL_PACKAGE];
  const sourcePartName = slide[SLIDE_PART_NAME];
  const sourcePart = pkg.getPart(sourcePartName);
  if (!sourcePart) throw new Error(`duplicateSlide: source ${sourcePartName} not found`);

  const presPart = pkg.getPart(PRES_PART_NAME);
  if (!presPart) throw new Error('presentation.xml missing');
  const presDoc = parseXml(decode(presPart.data));
  const sldIdLst = ensureSldIdLst(presDoc.root);
  const newSldId = allocateSldId(sldIdLst);

  const slideN = allocateSlideN(pkg);
  const newSlidePartName = partName(`/ppt/slides/slide${slideN}.xml`);
  pkg.addPart(newSlidePartName, sourcePart.contentType, new Uint8Array(sourcePart.data));

  const sourceRels = pkg.getRels(sourcePartName);
  if (sourceRels !== null) {
    pkg.setRels(newSlidePartName, { items: sourceRels.items.map((r) => ({ ...r })) });
  }

  const presRels = pkg.getRels(PRES_PART_NAME) ?? emptyRels();
  const newRId = nextRelId(presRels.items.map((r) => r.id));
  presRels.items.push({
    id: newRId,
    type: REL_TYPES.slide,
    target: `slides/slide${slideN}.xml`,
    targetMode: 'Internal',
  });
  pkg.setRels(PRES_PART_NAME, presRels);

  sldIdLst.children.push(
    elem(NAME_SLD_ID, {
      attrs: [attr(ATTR_ID, String(newSldId)), attr(ATTR_R_ID, newRId)],
    }),
  );
  presPart.data = encode(serializeXml(presDoc));

  pres._slidesCache = null;
  const slides = getSlides(pres);
  const dup = slides[slides.length - 1];
  if (!dup) throw new Error('duplicateSlide: post-condition failed');
  return dup;
};

/**
 * Convenience over `addSlide` + `moveSlide`. Inserts the new slide
 * at the given 0-based index (clamped to `[0, getSlides(pres).length]`).
 */
export const addSlideAt = (
  pres: PresentationData,
  atIndex: number,
  options: { layout: SlideLayoutData },
): SlideData => {
  const slide = addSlide(pres, options);
  moveSlide(pres, slide, atIndex);
  const slides = getSlides(pres);
  const clamped = Math.max(0, Math.min(atIndex, slides.length - 1));
  return slides[clamped]!;
};

/**
 * Convenience over `duplicateSlide` + `moveSlide`. Duplicates `slide`
 * and inserts the duplicate at `atIndex` instead of at the end.
 */
export const duplicateSlideAt = (
  pres: PresentationData,
  atIndex: number,
  slide: SlideData,
): SlideData => {
  const dup = duplicateSlide(pres, slide);
  moveSlide(pres, dup, atIndex);
  const slides = getSlides(pres);
  const clamped = Math.max(0, Math.min(atIndex, slides.length - 1));
  return slides[clamped]!;
};

/**
 * Imports a slide from another presentation into `targetPres`. The
 * slide's part bytes are copied verbatim; image rels are followed and
 * the linked media is copied into the target package with fresh part
 * names. The new slide is bound to the supplied `targetLayout` so it
 * still renders without the original deck's layouts.
 *
 * Limitations (v1):
 *
 *   - Only `image` rels are copied across. Other rels (charts, embedded
 *     workbooks, oleObjects, comments) are dropped from the imported
 *     slide. A diagnostic message is appended for each dropped rel.
 *   - Hyperlinks (external URLs) are preserved.
 *   - Slide → notesSlide is dropped (notes don't follow imports).
 *
 * Returns the new `SlideData` appended to `targetPres`.
 */
export const importSlide = (
  targetPres: PresentationData,
  sourceSlide: SlideData,
  targetLayout: SlideLayoutData,
): SlideData => {
  const sourcePkg = sourceSlide[INTERNAL_PACKAGE];
  const sourcePartName = sourceSlide[SLIDE_PART_NAME];
  const sourcePart = sourcePkg.getPart(sourcePartName);
  if (!sourcePart) throw new Error(`importSlide: source ${sourcePartName} not found`);
  const sourceRels = sourcePkg.getRels(sourcePartName);

  const targetPkg = targetPres[INTERNAL_PACKAGE];
  const presPart = targetPkg.getPart(PRES_PART_NAME);
  if (!presPart) throw new Error('presentation.xml missing in target');
  const presDoc = parseXml(decode(presPart.data));
  const sldIdLst = ensureSldIdLst(presDoc.root);
  const newSldId = allocateSldId(sldIdLst);

  const slideN = allocateSlideN(targetPkg);
  const newSlidePartName = partName(`/ppt/slides/slide${slideN}.xml`);

  // Copy the source slide bytes verbatim.
  targetPkg.addPart(newSlidePartName, sourcePart.contentType, new Uint8Array(sourcePart.data));

  // Build the new slide's rels:
  //   - one slideLayout pointing at the supplied target layout
  //   - one image rel per source image (with media imported)
  //   - external hyperlink rels copied verbatim
  const newRels = emptyRels();
  const layoutPartName = targetLayout[LAYOUT_PART_NAME];
  if (targetPkg.getPart(layoutPartName) === null) {
    throw new Error(`importSlide: layout ${layoutPartName} not in target package`);
  }
  newRels.items.push({
    id: 'rId1',
    type: REL_TYPES.slideLayout,
    target: `../slideLayouts/${basename(layoutPartName)}`,
    targetMode: 'Internal',
  });

  // Map from source rId → new rId so we can rewrite blip references later
  // (skipped in v1; we just preserve original rId values when possible).
  if (sourceRels !== null) {
    for (const rel of sourceRels.items) {
      if (rel.type === REL_TYPES.slideLayout) continue; // handled above
      if (rel.type === REL_TYPES.notesSlide) continue;
      if (rel.type === REL_TYPES.image && rel.targetMode === 'Internal') {
        // Copy the media part across with a fresh name.
        const mediaName = rel.target.startsWith('/')
          ? partName(rel.target)
          : resolveTarget(sourcePartName, rel.target);
        const mediaPart = sourcePkg.getPart(mediaName);
        if (!mediaPart) continue;
        const dotIdx = mediaName.lastIndexOf('.');
        const extension = dotIdx >= 0 ? mediaName.slice(dotIdx + 1) : 'bin';
        let nextN = 1;
        const re = /^\/ppt\/media\/image(\d+)\./;
        for (const p of targetPkg.parts) {
          const m = p.name.match(re);
          if (m?.[1] !== undefined) {
            const n = Number.parseInt(m[1], 10);
            if (Number.isFinite(n) && n >= nextN) nextN = n + 1;
          }
        }
        const newMediaName = partName(`/ppt/media/image${nextN}.${extension}`);
        setOpcDefault(targetPkg, extension.toLowerCase(), mediaPart.contentType);
        targetPkg.addPart(newMediaName, mediaPart.contentType, new Uint8Array(mediaPart.data));
        newRels.items.push({
          id: rel.id,
          type: REL_TYPES.image,
          target: `../media/image${nextN}.${extension}`,
          targetMode: 'Internal',
        });
        continue;
      }
      if (rel.type === REL_TYPES.hyperlink) {
        newRels.items.push({ ...rel });
        continue;
      }
      // Other internal rels (chart/oleObject/etc) are dropped in v1.
    }
  }
  targetPkg.setRels(newSlidePartName, newRels);

  // presentation → slide rel + sldIdLst entry.
  const presRels = targetPkg.getRels(PRES_PART_NAME) ?? emptyRels();
  const newRId = nextRelId(presRels.items.map((r) => r.id));
  presRels.items.push({
    id: newRId,
    type: REL_TYPES.slide,
    target: `slides/slide${slideN}.xml`,
    targetMode: 'Internal',
  });
  targetPkg.setRels(PRES_PART_NAME, presRels);

  sldIdLst.children.push(
    elem(NAME_SLD_ID, {
      attrs: [attr(ATTR_ID, String(newSldId)), attr(ATTR_R_ID, newRId)],
    }),
  );
  presPart.data = encode(serializeXml(presDoc));

  targetPres._slidesCache = null;
  const slides = getSlides(targetPres);
  const last = slides[slides.length - 1];
  if (!last) throw new Error('importSlide: post-condition failed');
  return last;
};

/**
 * Appends every slide from `sourcePres` into `targetPres`, in source
 * order. Built on top of `importSlide`: media is propagated, charts
 * are dropped (not yet supported across decks), and the slide's
 * layout is rebound to `targetLayout` on the target side.
 *
 * `targetLayout` can be a single layout used for every imported
 * slide (common), or a function called once per source slide for
 * per-slide layout selection.
 *
 * Returns the imported slides in target order.
 */
export const mergePresentations = (
  targetPres: PresentationData,
  sourcePres: PresentationData,
  targetLayout: SlideLayoutData | ((sourceSlide: SlideData, index: number) => SlideLayoutData),
): ReadonlyArray<SlideData> => {
  const sourceSlides = getSlides(sourcePres);
  const out: SlideData[] = [];
  const resolveLayout =
    typeof targetLayout === 'function' ? targetLayout : (): SlideLayoutData => targetLayout;
  for (let i = 0; i < sourceSlides.length; i++) {
    const src = sourceSlides[i]!;
    const layout = resolveLayout(src, i);
    out.push(importSlide(targetPres, src, layout));
  }
  return out;
};

// ---------------------------------------------------------------------------
// Slide-level reads.

/**
 * Shapes on a slide, in document order with group children flattened.
 */
export const getSlideShapes = (slide: SlideData): ReadonlyArray<SlideShapeData> =>
  slide[SLIDE_SHAPES];

/**
 * Rebinds the slide to a different layout. The slide's own content
 * (shapes, text, geometry) is preserved verbatim; only the
 * `slideLayout` rel is updated so PowerPoint re-renders with the new
 * layout's placeholder positions and theme.
 *
 * The new layout must already be a part of the package — pass one
 * returned by `getSlideLayouts(pres)` or `findSlideLayout(pres, name)`.
 */
export const setSlideLayout = (slide: SlideData, layout: SlideLayoutData): void => {
  const pkg = slide[INTERNAL_PACKAGE];
  const layoutPartName = layout[LAYOUT_PART_NAME];
  if (pkg.getPart(layoutPartName) === null) {
    throw new Error(`setSlideLayout: layout ${layoutPartName} not in package`);
  }
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
  const layoutBase = basename(layoutPartName);
  const newTarget = `../slideLayouts/${layoutBase}`;

  // Replace any existing slideLayout rel. Keep the same rId where
  // possible so other parts that already reference it stay valid.
  const existing = rels.items.find((r) => r.type === REL_TYPES.slideLayout);
  if (existing) {
    existing.target = newTarget;
  } else {
    rels.items.push({
      id: nextRelId(rels.items.map((r) => r.id)),
      type: REL_TYPES.slideLayout,
      target: newTarget,
      targetMode: 'Internal',
    });
  }
  pkg.setRels(slide[SLIDE_PART_NAME], rels);
};

/**
 * The slide layout this slide is bound to, or `null` if the slide has
 * no layout relationship.
 */
export const getSlideLayout = (slide: SlideData): SlideLayoutData | null => {
  const pkg = slide[INTERNAL_PACKAGE];
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
  if (rels === null) return null;
  const layoutRel = rels.items.find((r) => r.type === REL_TYPES.slideLayout);
  if (!layoutRel) return null;
  const layoutName = layoutRel.target.startsWith('/')
    ? partName(layoutRel.target)
    : resolveTarget(slide[SLIDE_PART_NAME], layoutRel.target);
  const layoutPart = pkg.getPart(layoutName);
  if (layoutPart === null) return null;
  const root = parseXml(decode(layoutPart.data)).root;
  return {
    [LAYOUT_PART_NAME]: layoutName,
    [LAYOUT_PART]: readSlideLayoutPart(root),
  };
};

/**
 * Returns the first placeholder shape with the given `type` (or `null`
 * if no match). Shapes whose `<p:ph>` omits an explicit type default to
 * `'body'` per ECMA-376 §19.7.10.
 */
export const findSlidePlaceholder = (slide: SlideData, type: string): SlideShapeData | null => {
  for (const shape of slide[SLIDE_SHAPES]) {
    const snap = shape[SHAPE_SNAPSHOT];
    if (snap.placeholderType === type) return shape;
    if (type === 'body' && snap.placeholderType === null && snap.placeholderIdx !== null) {
      return shape;
    }
  }
  return null;
};

/**
 * Returns every placeholder shape on the slide whose text body is
 * empty. Useful for "spot the slots that still need filling" UIs
 * before a slide is published, and for validation hooks that warn
 * about empty slots.
 */
export const findEmptyPlaceholders = (slide: SlideData): ReadonlyArray<SlideShapeData> => {
  const out: SlideShapeData[] = [];
  for (const shape of slide[SLIDE_SHAPES]) {
    if (!isShapePlaceholder(shape)) continue;
    if (hasShapeText(shape)) continue;
    out.push(shape);
  }
  return out;
};

/**
 * Returns the union bounding box of a group of shapes, or `null`
 * when none of them have bounds. Useful for "select all and move
 * together" patterns where the caller needs a single rectangle
 * across the group.
 */
export const getShapesBounds = (shapes: ReadonlyArray<SlideShapeData>): ShapeBounds | null => {
  let minX = Number.POSITIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  let maxY = Number.NEGATIVE_INFINITY;
  let found = false;
  for (const shape of shapes) {
    const b = getShapeBounds(shape);
    if (!b) continue;
    found = true;
    if (b.x < minX) minX = b.x;
    if (b.y < minY) minY = b.y;
    if (b.x + b.w > maxX) maxX = b.x + b.w;
    if (b.y + b.h > maxY) maxY = b.y + b.h;
  }
  if (!found) return null;
  return {
    x: minX as Emu,
    y: minY as Emu,
    w: (maxX - minX) as Emu,
    h: (maxY - minY) as Emu,
  };
};

/**
 * Translates every shape in `shapes` by `(dxEmu, dyEmu)`. Useful
 * for "move this group of shapes 1cm right" patterns without
 * looping yourself. Shapes without bounds are skipped silently.
 */
export const translateShapes = (
  shapes: ReadonlyArray<SlideShapeData>,
  dxEmu: number,
  dyEmu: number,
): void => {
  for (const shape of shapes) {
    const bounds = getShapeBounds(shape);
    if (bounds === null) continue;
    setShapePosition(shape, (bounds.x + dxEmu) as Emu, (bounds.y + dyEmu) as Emu);
  }
};

/**
 * Returns every slide in the deck that has at least one empty
 * placeholder shape. Built on `findEmptyPlaceholders`. Useful for
 * "which slides still need editorial attention?" pre-publish
 * audits.
 */
export const getSlidesWithEmptyPlaceholders = (
  pres: PresentationData,
): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    if (findEmptyPlaceholders(slide).length > 0) out.push(slide);
  }
  return out;
};

/**
 * Returns the first placeholder shape whose `<p:ph idx="...">`
 * matches `idx`, or `null` when none does. Real templates often
 * disambiguate same-type placeholders (e.g. two body slots) by
 * `idx`, so this is what you reach for when type-only lookup is
 * ambiguous.
 */
export const findSlidePlaceholderByIdx = (slide: SlideData, idx: number): SlideShapeData | null => {
  for (const shape of slide[SLIDE_SHAPES]) {
    if (shape[SHAPE_SNAPSHOT].placeholderIdx === idx) return shape;
  }
  return null;
};

/**
 * Returns every placeholder shape with the given `type`. Useful for
 * "two-content" / "comparison" layouts where multiple body
 * placeholders share a type and the caller needs to fill them all.
 * Like `findSlidePlaceholder`, omitted `<p:ph type>` is treated as
 * `body` per ECMA-376 §19.7.10.
 */
export const findSlidePlaceholders = (
  slide: SlideData,
  type: string,
): ReadonlyArray<SlideShapeData> => {
  const out: SlideShapeData[] = [];
  for (const shape of slide[SLIDE_SHAPES]) {
    const snap = shape[SHAPE_SNAPSHOT];
    if (snap.placeholderType === type) {
      out.push(shape);
      continue;
    }
    if (type === 'body' && snap.placeholderType === null && snap.placeholderIdx !== null) {
      out.push(shape);
    }
  }
  return out;
};

/**
 * First shape on the slide whose `cNvPr@name` equals `name`, or `null`
 * if none. Use the multi-match variant when more than one shape can
 * share the same name (common with template-cloned shapes).
 */
export const findShapeByName = (slide: SlideData, name: string): SlideShapeData | null => {
  for (const shape of slide[SLIDE_SHAPES]) {
    if (shape[SHAPE_SNAPSHOT].name === name) return shape;
  }
  return null;
};

/**
 * Returns the shape with the given OOXML internal id (`cNvPr@id`), or
 * `null` when no such shape exists. Shape ids are unique within a
 * slide; pair with `getShapeId` to round-trip references that arrive
 * from external XML (e.g. animations, hyperlinks).
 */
export const findShapeById = (slide: SlideData, id: number): SlideShapeData | null => {
  for (const shape of slide[SLIDE_SHAPES]) {
    if (shape[SHAPE_SNAPSHOT].id === id) return shape;
  }
  return null;
};

/** Every shape on the slide whose `cNvPr@name` equals `name`. */
export const findShapesByName = (slide: SlideData, name: string): ReadonlyArray<SlideShapeData> =>
  slide[SLIDE_SHAPES].filter((s) => s[SHAPE_SNAPSHOT].name === name);

/**
 * First shape on the slide whose visible text matches `needle`
 * (substring or `RegExp`), or `null` when none does. Convenience
 * over `getSlideShapes(slide).find(...)` when the caller is hunting
 * for a label in a template ("find the box that says 'Q1'").
 */
export const findShapeByText = (
  slide: SlideData,
  needle: string | RegExp,
): SlideShapeData | null => {
  for (const shape of slide[SLIDE_SHAPES]) {
    const text = shape[SHAPE_SNAPSHOT].text;
    if (typeof needle === 'string' ? text.includes(needle) : needle.test(text)) {
      return shape;
    }
  }
  return null;
};

/**
 * Every shape on the slide whose visible text matches `needle`. Use
 * when more than one shape can share the same text (common with
 * cloned bullet templates) — multi-match variant of
 * `findShapeByText`.
 */
export const findShapesByText = (
  slide: SlideData,
  needle: string | RegExp,
): ReadonlyArray<SlideShapeData> => {
  const out: SlideShapeData[] = [];
  for (const shape of slide[SLIDE_SHAPES]) {
    const text = shape[SHAPE_SNAPSHOT].text;
    if (typeof needle === 'string' ? text.includes(needle) : needle.test(text)) {
      out.push(shape);
    }
  }
  return out;
};

/** Every shape on the slide of the given kind. */
export const findShapesByKind = (
  slide: SlideData,
  kind: ShapeKind,
): ReadonlyArray<SlideShapeData> =>
  slide[SLIDE_SHAPES].filter((s) => s[SHAPE_SNAPSHOT].kind === kind);

/**
 * Returns the slide that owns `shape`. Useful when callers receive a
 * shape from an unfiltered walk (`getAllShapes`, `findShapeInPresentation`,
 * search results) and need to know which slide it's on.
 */
export const getShapeSlide = (shape: SlideShapeData): SlideData => shape[SHAPE_SLIDE];

/**
 * Returns the shape's current XML element as a string. Diagnostic
 * sibling of `getSlideXmlString`; useful for snapshot tests, bug
 * reports, and before/after dumps during transformations.
 */
export const getShapeXmlString = (shape: SlideShapeData): string =>
  serializeXml({
    kind: 'document',
    decl: null,
    root: shape[SHAPE_ELEMENT],
    prolog: [],
    epilog: [],
  });

/**
 * Returns the 0-based document-order index of `shape` on its slide,
 * or `-1` when the shape is stale (e.g. after a `removeShape` that
 * rebuilt the slide's shape list).
 */
export const getShapeIndex = (shape: SlideShapeData): number => {
  const shapes = shape[SHAPE_SLIDE][SLIDE_SHAPES];
  return shapes.indexOf(shape);
};

/**
 * Walks every slide and returns the first shape whose name matches.
 * Useful for "find the logo placeholder anywhere in the deck."
 */
export const findShapeInPresentation = (
  pres: PresentationData,
  name: string,
): SlideShapeData | null => {
  for (const slide of getSlides(pres)) {
    const hit = findShapeByName(slide, name);
    if (hit !== null) return hit;
  }
  return null;
};

/**
 * Replaces `{{key}}` tokens in every text-bearing shape on this slide.
 * Returns the number of substitutions performed.
 *
 * Tokens must fit within a single text run (see `replaceTokensInTree`
 * in `drawingml/`). Cross-run replacements aren't supported — use
 * `findSlidePlaceholder` + a setText path when PowerPoint has
 * fragmented the run sequence.
 */
export const replaceTokensInSlide = (slide: SlideData, tokens: Record<string, string>): number => {
  const n = replaceTokensInTree(slide[SLIDE_DOCUMENT].root, tokens);
  if (n > 0) {
    commitSlideData(slide);
    refreshSlideData(slide);
  }
  return n;
};

// ---------------------------------------------------------------------------
// SlideShape-level reads.

export const getShapeKind = (shape: SlideShapeData): ShapeKind => shape[SHAPE_SNAPSHOT].kind;

export const getShapeId = (shape: SlideShapeData): number => shape[SHAPE_SNAPSHOT].id;

/**
 * Returns the preset-geometry token (`'rect'`, `'ellipse'`, `'star5'`,
 * `'rightArrow'`, ...) for shapes whose body carries a
 * `<a:prstGeom prst="…"/>`. Returns `null` for:
 *
 *   - non-`'shape'` kinds (pictures, connectors, group shapes, tables,
 *     charts — they have their own geometry tags or no geometry),
 *   - shapes using custom geometry (`<a:custGeom>`),
 *   - shapes whose preset is missing (malformed but possible).
 *
 * Useful for renderers / inspectors that want to draw a faithful
 * approximation of each shape without dropping to the raw XML.
 */
export const getShapePreset = (shape: SlideShapeData): string | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'shape' && shape[SHAPE_SNAPSHOT].kind !== 'connector')
    return null;
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const prstGeom = firstChildElement(spPr, qname('a', 'prstGeom', NS.dml));
  if (!prstGeom) return null;
  for (const a of prstGeom.attrs) {
    if (a.name.localName === 'prst') return a.value;
  }
  return null;
};

/**
 * Reads the preset's adjust-handle values (`<a:prstGeom><a:avLst>
 * <a:gd name="adj" fmla="val 30000"/></a:avLst>`) as a map from guide
 * name → numeric value. Per ECMA-376 §20.1.9.4, guides are stored
 * with a formula prefix — `val 12345` is a literal number, and the
 * other prefixes (`pin`, `+-`, etc.) compute from other guides. We
 * only surface the `val` form because other formulas reference the
 * preset's built-in guides and don't make sense without them.
 *
 * Returns an empty record when no adjust values are authored (the
 * shape paints at its preset defaults).
 */
export const getShapeAdjustValues = (shape: SlideShapeData): Record<string, number> => {
  const out: Record<string, number> = {};
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return out;
  const prstGeom = firstChildElement(spPr, qname('a', 'prstGeom', NS.dml));
  if (!prstGeom) return out;
  const avLst = firstChildElement(prstGeom, qname('a', 'avLst', NS.dml));
  if (!avLst) return out;
  for (const gd of avLst.children) {
    if (gd.kind !== 'element' || gd.name.namespaceURI !== NS.dml || gd.name.localName !== 'gd')
      continue;
    const name = getAttrValue(gd, qname('', 'name', ''));
    const fmla = getAttrValue(gd, qname('', 'fmla', ''));
    if (!name || !fmla) continue;
    const match = /^val\s+(-?\d+(?:\.\d+)?)$/.exec(fmla);
    if (!match) continue;
    const n = Number.parseFloat(match[1]!);
    if (Number.isFinite(n)) out[name] = n;
  }
  return out;
};

/**
 * Returns the highest `cNvPr@id` used by any shape on the slide,
 * or `0` when the slide has no shapes with non-negative ids.
 *
 * Useful when hand-rolling a custom shape and you need an id known
 * not to collide. The official allocator inside `addSlideShape` /
 * `addSlideTextBox` etc. already does this — call those instead
 * when you don't need a custom id.
 */
export const getMaxShapeId = (slide: SlideData): number => {
  let max = 0;
  for (const shape of slide[SLIDE_SHAPES]) {
    const id = shape[SHAPE_SNAPSHOT].id;
    if (id > max) max = id;
  }
  return max;
};

/**
 * Deck-wide sibling of `getMaxShapeId`. Returns the highest
 * `cNvPr@id` across every shape on every slide, or `0` when the
 * deck has no shapes.
 *
 * Note: shape ids are scoped to a slide in OOXML — collisions
 * across slides are fine. This helper is for the rare cases where
 * a caller wants a single id known to be higher than anything in
 * the deck (e.g. to keep ids monotonically increasing).
 */
export const getMaxShapeIdInPresentation = (pres: PresentationData): number => {
  let max = 0;
  for (const slide of getSlides(pres)) {
    for (const shape of slide[SLIDE_SHAPES]) {
      const id = shape[SHAPE_SNAPSHOT].id;
      if (id > max) max = id;
    }
  }
  return max;
};

/**
 * Returns the number of slide masters declared in the
 * presentation's `<p:sldMasterIdLst>`. Most decks use exactly one
 * master; multi-master decks come from templates that combine
 * brand variants (e.g. a corporate master + a sponsor master).
 *
 * Returns `0` if `presentation.xml` is missing.
 */
export const getSlideMasterCount = (pres: PresentationData): number => {
  const pkg = pres[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (presPart === null) return 0;
  const root = parseXml(decode(presPart.data)).root;
  const model = readPresentationPart(root);
  return model.slideMasters.length;
};

/**
 * Returns the package part name of every slide master declared in
 * `presentation.xml`, resolved through the presentation's `.rels`.
 * Sibling of `getSlideMasterCount` for downstream tooling that
 * needs the master URIs (e.g. byte-level diff, custom validators).
 *
 * Returns an empty array when `presentation.xml` or its `.rels`
 * are missing.
 */
/**
 * Returns the part name of the slide master a slide inherits from
 * (`/ppt/slideMasters/slideMaster1.xml`), or `null` when the slide
 * has no layout or its layout has no master rel.
 *
 * Useful for multi-master decks where different slides live under
 * different brand templates and the caller needs to scope theme /
 * fontScheme / clrMap lookups to the correct master.
 */
export const getSlideMasterPartName = (slide: SlideData): string | null => {
  const layout = getSlideLayout(slide);
  if (!layout) return null;
  const pkg = slide[INTERNAL_PACKAGE];
  const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
  const layoutRels = pkg.getRels(layoutPartName);
  if (!layoutRels) return null;
  const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
  if (!masterRel) return null;
  return resolveTarget(layoutPartName, masterRel.target);
};

export const getSlideMasterPartNames = (pres: PresentationData): ReadonlyArray<string> => {
  const pkg = pres[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (presPart === null) return [];
  const root = parseXml(decode(presPart.data)).root;
  const model = readPresentationPart(root);
  const rels = pkg.getRels(PRES_PART_NAME);
  if (rels === null) return [];
  const out: string[] = [];
  for (const m of model.slideMasters) {
    const rel = rels.items.find((r) => r.id === m.rId);
    if (rel === undefined) continue;
    const resolved = rel.target.startsWith('/')
      ? partName(rel.target)
      : resolveTarget(PRES_PART_NAME, rel.target);
    out.push(resolved);
  }
  return out;
};

export const getShapeName = (shape: SlideShapeData): string => shape[SHAPE_SNAPSHOT].name;

/**
 * Renames the shape's `cNvPr@name`. The display name is what
 * PowerPoint shows in the Selection Pane and what `findShapeByName`
 * matches on. Empty strings are allowed (matches PowerPoint behavior).
 */
/**
 * Reads the shape's alt-text description (`<p:cNvPr descr="...">`).
 * Accessibility tools (screen readers, contrast checkers) and
 * PowerPoint's "Alt Text" pane look at this field. Returns `null`
 * when no description is set.
 */
export const getShapeDescription = (shape: SlideShapeData): string | null => {
  const cNvPr = findCNvPr(shape);
  if (!cNvPr) return null;
  return getAttrValue(cNvPr, qname('', 'descr', ''));
};

/**
 * Sets the shape's alt-text description (`<p:cNvPr descr="...">`).
 * Pass `null` to clear. Important for accessibility — image shapes
 * and decorative graphics should carry a descr that conveys the
 * visual meaning to screen readers.
 */
export const setShapeDescription = (shape: SlideShapeData, description: string | null): void => {
  const cNvPr = findCNvPr(shape);
  if (!cNvPr) {
    throw new Error(`setShapeDescription: ${shape[SHAPE_SNAPSHOT].kind} shape has no cNvPr`);
  }
  cNvPr.attrs = cNvPr.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'descr'),
  );
  if (description !== null && description !== '') {
    cNvPr.attrs.push(attr(qname('', 'descr', ''), description));
  }
  commitAndRefresh(shape);
};

/**
 * Reads the shape's alt-text title (`<p:cNvPr title="...">`).
 * PowerPoint surfaces this alongside `descr` in its Alt Text pane
 * as a short heading. Returns `null` when no title is set.
 */
export const getShapeAltTitle = (shape: SlideShapeData): string | null => {
  const cNvPr = findCNvPr(shape);
  if (!cNvPr) return null;
  return getAttrValue(cNvPr, qname('', 'title', ''));
};

/**
 * Sets the shape's alt-text title (`<p:cNvPr title="...">`). Pass
 * `null` to clear. Distinct from `renameShape`, which writes the
 * `name` attribute used in the selection pane.
 */
export const setShapeAltTitle = (shape: SlideShapeData, title: string | null): void => {
  const cNvPr = findCNvPr(shape);
  if (!cNvPr) {
    throw new Error(`setShapeAltTitle: ${shape[SHAPE_SNAPSHOT].kind} shape has no cNvPr`);
  }
  cNvPr.attrs = cNvPr.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'title'),
  );
  if (title !== null && title !== '') {
    cNvPr.attrs.push(attr(qname('', 'title', ''), title));
  }
  commitAndRefresh(shape);
};

/**
 * `true` when the shape's `<p:cNvPr hidden="1">` is set. Hidden
 * shapes are skipped by PowerPoint's renderer but stay in the
 * shape tree — useful for variant slides that toggle which boxes
 * are visible.
 */
export const isShapeHidden = (shape: SlideShapeData): boolean => {
  const cNvPr = findCNvPr(shape);
  if (!cNvPr) return false;
  const v = getAttrValue(cNvPr, qname('', 'hidden', ''));
  return v === '1' || v === 'true';
};

/**
 * Sets or clears `<p:cNvPr hidden="...">` on the shape. Hidden
 * shapes remain in the document but PowerPoint doesn't render them.
 */
export const setShapeHidden = (shape: SlideShapeData, hidden: boolean): void => {
  const cNvPr = findCNvPr(shape);
  if (!cNvPr) {
    throw new Error(`setShapeHidden: ${shape[SHAPE_SNAPSHOT].kind} shape has no cNvPr`);
  }
  cNvPr.attrs = cNvPr.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'hidden'),
  );
  if (hidden) cNvPr.attrs.push(attr(qname('', 'hidden', ''), '1'));
  commitAndRefresh(shape);
};

export const renameShape = (shape: SlideShapeData, newName: string): void => {
  const cNvPr = findCNvPr(shape);
  if (!cNvPr) {
    throw new Error(`renameShape: ${shape[SHAPE_SNAPSHOT].kind} shape has no cNvPr to rename`);
  }
  cNvPr.attrs = cNvPr.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'name'),
  );
  cNvPr.attrs.push(attr(qname('', 'name', ''), newName));
  commitAndRefresh(shape);
};

export const getShapePlaceholderType = (shape: SlideShapeData): string | null =>
  shape[SHAPE_SNAPSHOT].placeholderType;

export const getShapePlaceholderIdx = (shape: SlideShapeData): number | null =>
  shape[SHAPE_SNAPSHOT].placeholderIdx;

/**
 * `true` when the shape carries `<p:nvSpPr><p:nvPr><p:ph>` — i.e. it
 * inherits from a layout/master placeholder. False for decorative
 * geometry the slide author dropped onto the canvas. Decoupled from
 * the more specific `getShapePlaceholderType` / `getShapePlaceholderIdx`
 * (either can be null on a real placeholder; together they identify it).
 */
export const isShapePlaceholder = (shape: SlideShapeData): boolean => {
  const snap = shape[SHAPE_SNAPSHOT];
  return snap.placeholderType !== null || snap.placeholderIdx !== null;
};

export const getShapeText = (shape: SlideShapeData): string => shape[SHAPE_SNAPSHOT].text;

export const getShapePosition = (shape: SlideShapeData): Position | null =>
  readPosition(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind);

export const getShapeSize = (shape: SlideShapeData): Size | null =>
  readSize(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind);

export const getShapeRotation = (shape: SlideShapeData): number =>
  readRotation(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind);

export const getShapeFlip = (
  shape: SlideShapeData,
): { horizontal: boolean; vertical: boolean } | null =>
  readFlip(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind);

/**
 * Enumerates the shapes nested inside a `<p:grpSp>` group, one level
 * deep (nested groups come through as `kind: 'group'` themselves —
 * call this again on each one to recurse).
 *
 * Returns an empty array for non-group shapes. Each child carries its
 * own bounds in the group's *internal* coordinate system; pair with
 * `getGroupTransform` to project them onto the slide.
 */
export const getGroupChildren = (shape: SlideShapeData): ReadonlyArray<SlideShapeData> => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'group') return [];
  const children = readGroupChildren(shape[SHAPE_ELEMENT]);
  return children.map((child) => ({
    [SHAPE_SLIDE]: shape[SHAPE_SLIDE],
    [SHAPE_ELEMENT]: child.element,
    [SHAPE_SNAPSHOT]: child,
  }));
};

/**
 * Returns the group's slide-relative bounds (`outer`) and the internal
 * coordinate system the children's `<a:xfrm>` values live in
 * (`inner`). Renderers project a child point `(cx, cy)` onto the slide
 * with:
 *
 *   slideX = outer.x + (cx - inner.x) * (outer.w / inner.w)
 *   slideY = outer.y + (cy - inner.y) * (outer.h / inner.h)
 *
 * Returns `null` for non-group shapes or for groups whose
 * `<p:grpSpPr>` omits an `<a:xfrm>`.
 */
export const getGroupTransform = (
  shape: SlideShapeData,
): {
  readonly outer: ShapeBounds;
  readonly inner: ShapeBounds;
} | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'group') return null;
  const grpSpPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'grpSpPr', NS.pml));
  if (!grpSpPr) return null;
  const xfrm = firstChildElement(grpSpPr, qname('a', 'xfrm', NS.dml));
  if (!xfrm) return null;
  const off = firstChildElement(xfrm, qname('a', 'off', NS.dml));
  const ext = firstChildElement(xfrm, qname('a', 'ext', NS.dml));
  const chOff = firstChildElement(xfrm, qname('a', 'chOff', NS.dml));
  const chExt = firstChildElement(xfrm, qname('a', 'chExt', NS.dml));
  if (!off || !ext) return null;
  const parseAttr = (el: XmlElement, name: string): number | null => {
    const raw = getAttrValue(el, qname('', name, ''));
    if (raw === null) return null;
    const n = Number.parseInt(raw, 10);
    return Number.isFinite(n) ? n : null;
  };
  const ox = parseAttr(off, 'x');
  const oy = parseAttr(off, 'y');
  const ow = parseAttr(ext, 'cx');
  const oh = parseAttr(ext, 'cy');
  if (ox === null || oy === null || ow === null || oh === null) return null;
  // Per ECMA-376, `<a:chOff>/<a:chExt>` default to the same values as
  // `<a:off>/<a:ext>` when omitted (i.e. no internal-to-outer
  // transform).
  const ix = chOff ? (parseAttr(chOff, 'x') ?? ox) : ox;
  const iy = chOff ? (parseAttr(chOff, 'y') ?? oy) : oy;
  const iw = chExt ? (parseAttr(chExt, 'cx') ?? ow) : ow;
  const ih = chExt ? (parseAttr(chExt, 'cy') ?? oh) : oh;
  return {
    outer: { x: ox as Emu, y: oy as Emu, w: ow as Emu, h: oh as Emu },
    inner: { x: ix as Emu, y: iy as Emu, w: iw as Emu, h: ih as Emu },
  };
};

/**
 * Combined bounds — position + size in one object. Returns `null` when
 * the shape inherits both position and size from its layout (so the
 * `<a:xfrm>` element is absent or incomplete).
 */
export interface ShapeBounds {
  readonly x: Emu;
  readonly y: Emu;
  readonly w: Emu;
  readonly h: Emu;
}

export const getShapeBounds = (shape: SlideShapeData): ShapeBounds | null => {
  const pos = readPosition(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind);
  const size = readSize(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind);
  if (pos === null || size === null) return null;
  return {
    x: pos.x as Emu,
    y: pos.y as Emu,
    w: size.w as Emu,
    h: size.h as Emu,
  };
};

/**
 * Same as `getShapeBounds` but walks the placeholder inheritance chain
 * when the shape has no `<a:xfrm>` of its own:
 *
 *   1. The shape's own bounds.
 *   2. The matching placeholder on the slide's layout (by `<p:ph idx>`,
 *      falling back to `<p:ph type>`).
 *   3. The matching placeholder on the layout's slide master.
 *
 * This is what renderers want when they need to draw a placeholder
 * that the deck author left unsized — real templates only override
 * geometry per placeholder when it differs from the master.
 *
 * Returns `null` when none of the three levels carries explicit bounds.
 */
export const getShapeBoundsResolved = (
  pres: PresentationData,
  shape: SlideShapeData,
): ShapeBounds | null => {
  const direct = getShapeBounds(shape);
  if (direct) return direct;

  const slide = shape[SHAPE_SLIDE];
  const layout = getSlideLayout(slide);
  if (!layout) return null;

  const phIdx = getShapePlaceholderIdx(shape);
  const phType = getShapePlaceholderType(shape);

  const findInShapes = (
    shapes: ReadonlyArray<{
      placeholderIdx: number | null;
      placeholderType: string | null;
      element: XmlElement;
      kind: ShapeKind;
    }>,
  ): ShapeBounds | null => {
    let match = phIdx !== null ? shapes.find((s) => s.placeholderIdx === phIdx) : undefined;
    if (!match && phType !== null) {
      match = shapes.find((s) => s.placeholderType === phType);
    }
    if (!match) return null;
    const pos = readPosition(match.element, match.kind);
    const size = readSize(match.element, match.kind);
    if (pos === null || size === null) return null;
    return {
      x: pos.x as Emu,
      y: pos.y as Emu,
      w: size.w as Emu,
      h: size.h as Emu,
    };
  };

  const layoutHit = findInShapes(layout[LAYOUT_PART].shapes);
  if (layoutHit) return layoutHit;

  // Walk one level up: layout → slideMaster rel.
  const pkg = pres[INTERNAL_PACKAGE];
  const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
  const layoutRels = pkg.getRels(layoutPartName);
  if (!layoutRels) return null;
  const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
  if (!masterRel) return null;
  const masterPart = pkg.getPart(resolveTarget(layoutPartName, masterRel.target));
  if (!masterPart) return null;

  const masterRoot = parseXml(decode(masterPart.data)).root;
  const { shapes: masterShapes } = readShapeTreeFromCsldRoot(masterRoot, 'sldMaster');
  return findInShapes(masterShapes);
};

/**
 * Returns the center point of the shape's bounds in EMU, or `null`
 * when the shape has no `<a:xfrm>`. Convenience for layout
 * pipelines that compute alignment / overlap from center points.
 */
export const getShapeCenter = (
  shape: SlideShapeData,
): { readonly x: Emu; readonly y: Emu } | null => {
  const bounds = getShapeBounds(shape);
  if (bounds === null) return null;
  return {
    x: (bounds.x + Math.round(bounds.w / 2)) as Emu,
    y: (bounds.y + Math.round(bounds.h / 2)) as Emu,
  };
};

/**
 * `true` when point `(x, y)` (in EMU) lies inside the shape's
 * axis-aligned bounds. Closed on the top-left edge, open on the
 * bottom-right (standard half-open rectangle). Returns `false`
 * when the shape has no bounds.
 *
 * Useful for hit-testing in custom interaction handlers.
 */
export const pointInShape = (shape: SlideShapeData, x: number, y: number): boolean => {
  const bounds = getShapeBounds(shape);
  if (bounds === null) return false;
  return x >= bounds.x && x < bounds.x + bounds.w && y >= bounds.y && y < bounds.y + bounds.h;
};

/**
 * Returns every shape on the slide whose bounds contain `(x, y)`
 * (in EMU). Built on `pointInShape`. The list is in document
 * order, so callers can index by z-stack from front (last) to
 * back (first) if they want one-hit semantics.
 */
export const findShapesAtPoint = (
  slide: SlideData,
  x: number,
  y: number,
): ReadonlyArray<SlideShapeData> => slide[SLIDE_SHAPES].filter((s) => pointInShape(s, x, y));

/**
 * Moves the shape so its center sits at the slide canvas center.
 * Reads the presentation's slide size, then sets the shape's
 * position to `(slideWidth/2 - shapeWidth/2, slideHeight/2 - shapeHeight/2)`.
 *
 * No-op when the shape has no bounds or the presentation has no
 * configured slide size.
 */
export const centerShapeOnSlide = (shape: SlideShapeData): void => {
  const bounds = getShapeBounds(shape);
  if (bounds === null) return;
  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (!presPart) return;
  const presRoot = parseXml(decode(presPart.data)).root;
  const model = readPresentationPart(presRoot);
  if (model.slideSize === null) return;
  const newX = Math.round(model.slideSize.cx / 2 - bounds.w / 2) as Emu;
  const newY = Math.round(model.slideSize.cy / 2 - bounds.h / 2) as Emu;
  setShapePosition(shape, newX, newY);
};

/**
 * `true` when two shapes' axis-aligned bounding boxes overlap.
 * Returns `false` when either shape has no bounds. Doesn't account
 * for rotation — uses the raw `<a:xfrm>` rectangle, not the
 * visual bounding box after rotation.
 *
 * Useful for collision detection in custom layout pipelines.
 */
export const shapesOverlap = (a: SlideShapeData, b: SlideShapeData): boolean => {
  const ba = getShapeBounds(a);
  const bb = getShapeBounds(b);
  if (ba === null || bb === null) return false;
  return ba.x < bb.x + bb.w && ba.x + ba.w > bb.x && ba.y < bb.y + bb.h && ba.y + ba.h > bb.y;
};

/**
 * Sets both position and size in one call. Equivalent to calling
 * `setShapePosition` followed by `setShapeSize`, but commits the slide
 * just once.
 */
export const setShapeBounds = (shape: SlideShapeData, bounds: ShapeBounds): void => {
  writePosition(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind, bounds.x, bounds.y);
  writeSize(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind, bounds.w, bounds.h);
  commitAndRefresh(shape);
};

/**
 * Reads back the fill choice on the shape's `<p:spPr>`. Returns:
 *
 *   - `{ kind: 'solid', color: '#RRGGBB' }` for a solid sRGB fill.
 *   - `{ kind: 'solid', color: 'scheme:accent1' }` for a scheme color.
 *   - `{ kind: 'gradient' }` / `'pattern'` / `'image'` for those choices
 *     (without breaking out their parameters — call the dedicated
 *     setter to overwrite).
 *   - `{ kind: 'none' }` for `<a:noFill>`.
 *   - `{ kind: 'inherit' }` when no fill choice is present on this
 *     shape (it inherits from the layout / master placeholder).
 */
export type ShapeFill =
  | { readonly kind: 'solid'; readonly color: string }
  | { readonly kind: 'gradient' }
  | { readonly kind: 'pattern' }
  | { readonly kind: 'image' }
  | { readonly kind: 'none' }
  | { readonly kind: 'inherit' };

/**
 * Reads back the shape's stroke (`<a:ln>`). Returns:
 *
 *   - `{ kind: 'solid', color, widthEmu? }` for a solid-color outline.
 *   - `{ kind: 'none' }` when an `<a:noFill>` sits inside `<a:ln>`.
 *   - `{ kind: 'inherit' }` when no `<a:ln>` is present.
 */
export type ShapeStroke =
  | { readonly kind: 'solid'; readonly color: string; readonly widthEmu?: number }
  | { readonly kind: 'none' }
  | { readonly kind: 'inherit' };

/**
 * Convenience over `getShapeStroke(shape)`: returns the solid-
 * stroke color (`#RRGGBB` / `scheme:<token>`) or `null` when the
 * stroke is inherited / removed.
 */
export const getShapeStrokeColor = (shape: SlideShapeData): string | null => {
  const stroke = getShapeStroke(shape);
  return stroke.kind === 'solid' ? stroke.color : null;
};

/**
 * Convenience over `getShapeStroke(shape)`: returns the stroke
 * width in EMU when the stroke is solid and an explicit width is
 * set, or `null` otherwise.
 */
export const getShapeStrokeWidth = (shape: SlideShapeData): number | null => {
  const stroke = getShapeStroke(shape);
  return stroke.kind === 'solid' && stroke.widthEmu !== undefined ? stroke.widthEmu : null;
};

/**
 * Returns the shape's stroke color resolved to a concrete `#RRGGBB`:
 * scheme tokens are mapped through the deck's color scheme and
 * `<a:lumMod>` / `<a:tint>` / `<a:shade>` / etc. transform children
 * are applied. Returns `null` when the stroke isn't a solid color
 * (inherits / `noFill`) or when the color can't be resolved.
 *
 * Companion to `getShapeStrokeColor`, which surfaces only the raw
 * `#RRGGBB` / `scheme:<token>` string — fine for round-tripping but
 * wrong for rendering, because PowerPoint paints the *transformed*
 * color, not the base one.
 */
export const getShapeStrokeColorResolved = (
  pres: PresentationData,
  shape: SlideShapeData,
): string | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
  if (!ln) return null;
  const solid = firstChildElement(ln, qname('a', 'solidFill', NS.dml));
  if (!solid) return null;
  for (const inner of solid.children) {
    if (inner.kind !== 'element' || inner.name.namespaceURI !== NS.dml) continue;
    return resolveDrawingColor(inner, getPresentationTheme(pres));
  }
  return null;
};

/**
 * Reads the stroke's line cap style — `'rnd'` (round), `'sq'` (square),
 * `'flat'`, or `null` when the attribute isn't set. Per ECMA-376
 * §20.1.2.3.10 (`ST_LineCap`).
 */
export const getShapeStrokeCap = (shape: SlideShapeData): 'rnd' | 'sq' | 'flat' | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
  if (!ln) return null;
  const v = getAttrValue(ln, qname('', 'cap', ''));
  if (v === 'rnd' || v === 'sq' || v === 'flat') return v;
  return null;
};

/**
 * Reads the stroke's line join style — `'round'` / `'bevel'` / `'miter'`,
 * or `null` when no explicit join element is present. Maps from the
 * three child-element variants `<a:round/>`, `<a:bevel/>`, `<a:miter/>`.
 */
export const getShapeStrokeJoin = (shape: SlideShapeData): 'round' | 'bevel' | 'miter' | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
  if (!ln) return null;
  for (const c of ln.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (c.name.localName === 'round') return 'round';
    if (c.name.localName === 'bevel') return 'bevel';
    if (c.name.localName === 'miter') return 'miter';
  }
  return null;
};

/**
 * Reads the stroke's compound-line style (`<a:ln cmpd="…">`) — single,
 * double, triple, or thick/thin / thin/thick parallel lines. ECMA-376
 * §20.1.2.3.11 (`ST_CompoundLine`).
 */
export const getShapeStrokeCompound = (
  shape: SlideShapeData,
): 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri' | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
  if (!ln) return null;
  const v = getAttrValue(ln, qname('', 'cmpd', ''));
  if (v === 'sng' || v === 'dbl' || v === 'thickThin' || v === 'thinThick' || v === 'tri') return v;
  return null;
};

/**
 * Same as `getShapeStroke` but walks the layout → master placeholder
 * cascade when the shape itself reports `'inherit'`. First non-inherit
 * stroke layer wins.
 */
export const getShapeStrokeEffective = (
  pres: PresentationData,
  shape: SlideShapeData,
): ShapeStroke => {
  const own = getShapeStroke(shape);
  if (own.kind !== 'inherit') return own;

  const phIdx = getShapePlaceholderIdx(shape);
  const phType = getShapePlaceholderType(shape);
  if (phIdx === null && phType === null) return own;

  const layout = getSlideLayout(shape[SHAPE_SLIDE]);
  if (!layout) return own;

  const readStrokeFromSpPr = (el: XmlElement): ShapeStroke | null => {
    const spPr = firstChildElement(el, qname('p', 'spPr', NS.pml));
    if (!spPr) return null;
    const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
    if (!ln) return null;
    const wRaw = getAttrValue(ln, qname('', 'w', ''));
    const widthEmu = wRaw !== null ? Number.parseInt(wRaw, 10) : undefined;
    for (const c of ln.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      if (c.name.localName === 'noFill') return { kind: 'none' };
      if (c.name.localName === 'solidFill') {
        for (const inner of c.children) {
          if (inner.kind !== 'element' || inner.name.namespaceURI !== NS.dml) continue;
          if (inner.name.localName === 'srgbClr') {
            const val = getAttrValue(inner, qname('', 'val', ''));
            if (val !== null) {
              return {
                kind: 'solid',
                color: `#${val.toUpperCase()}`,
                ...(widthEmu !== undefined ? { widthEmu } : {}),
              };
            }
          }
          if (inner.name.localName === 'schemeClr') {
            const val = getAttrValue(inner, qname('', 'val', ''));
            if (val !== null) {
              return {
                kind: 'solid',
                color: `scheme:${val}`,
                ...(widthEmu !== undefined ? { widthEmu } : {}),
              };
            }
          }
        }
      }
    }
    return null;
  };

  const findPh = (
    shapes: ReadonlyArray<{
      placeholderIdx: number | null;
      placeholderType: string | null;
      element: XmlElement;
    }>,
  ): XmlElement | null => {
    let match = phIdx !== null ? shapes.find((s) => s.placeholderIdx === phIdx) : undefined;
    if (!match && phType !== null) match = shapes.find((s) => s.placeholderType === phType);
    return match?.element ?? null;
  };

  const layoutPh = findPh(layout[LAYOUT_PART].shapes);
  if (layoutPh) {
    const s = readStrokeFromSpPr(layoutPh);
    if (s) return s;
  }
  const pkg = pres[INTERNAL_PACKAGE];
  const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
  const layoutRels = pkg.getRels(layoutPartName);
  if (!layoutRels) return own;
  const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
  if (!masterRel) return own;
  const masterPart = pkg.getPart(resolveTarget(layoutPartName, masterRel.target));
  if (!masterPart) return own;
  const masterRoot = parseXml(decode(masterPart.data)).root;
  const { shapes: masterShapes } = readShapeTreeFromCsldRoot(masterRoot, 'sldMaster');
  const masterPh = findPh(masterShapes);
  if (masterPh) {
    const s = readStrokeFromSpPr(masterPh);
    if (s) return s;
  }
  return own;
};

export const getShapeStroke = (shape: SlideShapeData): ShapeStroke => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return { kind: 'inherit' };
  const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
  if (!ln) return { kind: 'inherit' };

  const wRaw = getAttrValue(ln, qname('', 'w', ''));
  const widthEmu = wRaw !== null ? Number.parseInt(wRaw, 10) : undefined;

  for (const c of ln.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (c.name.localName === 'noFill') return { kind: 'none' };
    if (c.name.localName === 'solidFill') {
      for (const inner of c.children) {
        if (inner.kind !== 'element' || inner.name.namespaceURI !== NS.dml) continue;
        if (inner.name.localName === 'srgbClr') {
          const val = getAttrValue(inner, qname('', 'val', ''));
          if (val !== null) {
            return {
              kind: 'solid',
              color: `#${val.toUpperCase()}`,
              ...(widthEmu !== undefined ? { widthEmu } : {}),
            };
          }
        }
        if (inner.name.localName === 'schemeClr') {
          const val = getAttrValue(inner, qname('', 'val', ''));
          if (val !== null) {
            return {
              kind: 'solid',
              color: `scheme:${val}`,
              ...(widthEmu !== undefined ? { widthEmu } : {}),
            };
          }
        }
      }
      return {
        kind: 'solid',
        color: '',
        ...(widthEmu !== undefined ? { widthEmu } : {}),
      };
    }
  }
  return { kind: 'inherit' };
};

/**
 * Convenience over `getShapeFill(shape)`: returns the solid-fill
 * color string (`#RRGGBB` or `scheme:<token>`) when the shape has
 * one, or `null` otherwise. Use when the caller only cares about
 * the color and doesn't need to distinguish "inherit" / "no fill" /
 * "gradient" / "pattern" / "image" from each other.
 */
export const getShapeFillColor = (shape: SlideShapeData): string | null => {
  const fill = getShapeFill(shape);
  return fill.kind === 'solid' ? fill.color : null;
};

/**
 * Returns the shape's solid fill resolved to a concrete `#RRGGBB`:
 * scheme tokens are mapped through the deck's color scheme and
 * `<a:lumMod>` / `<a:tint>` / `<a:shade>` / etc. transform children
 * are applied. Returns `null` when the fill isn't solid (gradient,
 * pattern, image, none, inherit) or when the color can't be resolved.
 *
 * Companion to `getShapeFillColor`, which surfaces only the raw
 * `#RRGGBB` / `scheme:<token>` string. Renderers and exporters that
 * need the color PowerPoint actually paints should call this.
 */
export const getShapeFillColorResolved = (
  pres: PresentationData,
  shape: SlideShapeData,
): string | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const solid = firstChildElement(spPr, qname('a', 'solidFill', NS.dml));
  if (!solid) return null;
  for (const inner of solid.children) {
    if (inner.kind !== 'element' || inner.name.namespaceURI !== NS.dml) continue;
    return resolveDrawingColor(inner, getPresentationTheme(pres));
  }
  return null;
};

/**
 * Same as `getShapeFill` but walks the layout → master placeholder
 * cascade when the shape itself reports `'inherit'`. Returns the first
 * non-inherit fill found, or `{ kind: 'inherit' }` when neither layer
 * supplies one. Useful for renderers that want the actual fill the
 * placeholder will paint with.
 */
export const getShapeFillEffective = (pres: PresentationData, shape: SlideShapeData): ShapeFill => {
  const own = getShapeFill(shape);
  if (own.kind !== 'inherit') return own;

  const phIdx = getShapePlaceholderIdx(shape);
  const phType = getShapePlaceholderType(shape);
  if (phIdx === null && phType === null) return own;

  const layout = getSlideLayout(shape[SHAPE_SLIDE]);
  if (!layout) return own;

  const readFillFromSpPr = (el: XmlElement): ShapeFill | null => {
    const spPr = firstChildElement(el, qname('p', 'spPr', NS.pml));
    if (!spPr) return null;
    for (const c of spPr.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      switch (c.name.localName) {
        case 'noFill':
          return { kind: 'none' };
        case 'solidFill': {
          for (const inner of c.children) {
            if (inner.kind !== 'element' || inner.name.namespaceURI !== NS.dml) continue;
            if (inner.name.localName === 'srgbClr') {
              const val = getAttrValue(inner, qname('', 'val', ''));
              if (val !== null) return { kind: 'solid', color: `#${val.toUpperCase()}` };
            }
            if (inner.name.localName === 'schemeClr') {
              const val = getAttrValue(inner, qname('', 'val', ''));
              if (val !== null) return { kind: 'solid', color: `scheme:${val}` };
            }
          }
          return { kind: 'solid', color: '' };
        }
        case 'gradFill':
          return { kind: 'gradient' };
        case 'pattFill':
          return { kind: 'pattern' };
        case 'blipFill':
          return { kind: 'image' };
      }
    }
    return null;
  };

  const findPh = (
    shapes: ReadonlyArray<{
      placeholderIdx: number | null;
      placeholderType: string | null;
      element: XmlElement;
    }>,
  ): XmlElement | null => {
    let match = phIdx !== null ? shapes.find((s) => s.placeholderIdx === phIdx) : undefined;
    if (!match && phType !== null) match = shapes.find((s) => s.placeholderType === phType);
    return match?.element ?? null;
  };

  const layoutPh = findPh(layout[LAYOUT_PART].shapes);
  if (layoutPh) {
    const f = readFillFromSpPr(layoutPh);
    if (f) return f;
  }

  const pkg = pres[INTERNAL_PACKAGE];
  const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
  const layoutRels = pkg.getRels(layoutPartName);
  if (!layoutRels) return own;
  const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
  if (!masterRel) return own;
  const masterPart = pkg.getPart(resolveTarget(layoutPartName, masterRel.target));
  if (!masterPart) return own;
  const masterRoot = parseXml(decode(masterPart.data)).root;
  const { shapes: masterShapes } = readShapeTreeFromCsldRoot(masterRoot, 'sldMaster');
  const masterPh = findPh(masterShapes);
  if (masterPh) {
    const f = readFillFromSpPr(masterPh);
    if (f) return f;
  }
  return own;
};

export const getShapeFill = (shape: SlideShapeData): ShapeFill => {
  const spPrName = qname('p', 'spPr', NS.pml);
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], spPrName);
  if (!spPr) return { kind: 'inherit' };
  for (const c of spPr.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    switch (c.name.localName) {
      case 'noFill':
        return { kind: 'none' };
      case 'solidFill': {
        // Look for the immediate color choice; report sRGB verbatim,
        // scheme colors as "scheme:<token>".
        for (const inner of c.children) {
          if (inner.kind !== 'element' || inner.name.namespaceURI !== NS.dml) continue;
          if (inner.name.localName === 'srgbClr') {
            const val = getAttrValue(inner, qname('', 'val', ''));
            if (val !== null) return { kind: 'solid', color: `#${val.toUpperCase()}` };
          }
          if (inner.name.localName === 'schemeClr') {
            const val = getAttrValue(inner, qname('', 'val', ''));
            if (val !== null) return { kind: 'solid', color: `scheme:${val}` };
          }
        }
        return { kind: 'solid', color: '' };
      }
      case 'gradFill':
        return { kind: 'gradient' };
      case 'pattFill':
        return { kind: 'pattern' };
      case 'blipFill':
        return { kind: 'image' };
    }
  }
  return { kind: 'inherit' };
};

// ---------------------------------------------------------------------------
// Detailed gradient-fill reader. Companion to `getShapeFill`, which
// only reports the discriminated `kind`. Returns the full stop list +
// angle when the shape carries a `<a:gradFill>` of its own, or
// `null` for solid / pattern / image / none / inherited fills.
//
// Useful for renderers (preview generators, PDF exporters) that need
// to reproduce the gradient instead of substituting a placeholder.

const NAME_A_GRAD_FILL = qname('a', 'gradFill', NS.dml);
const NAME_A_GS_LST = qname('a', 'gsLst', NS.dml);
const NAME_A_LIN = qname('a', 'lin', NS.dml);

const readColorFromContainer = (parent: XmlElement): string | null => {
  for (const c of parent.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (c.name.localName === 'srgbClr') {
      const val = getAttrValue(c, qname('', 'val', ''));
      if (val !== null) return `#${val.toUpperCase()}`;
    }
    if (c.name.localName === 'schemeClr') {
      const val = getAttrValue(c, qname('', 'val', ''));
      if (val !== null) return `scheme:${val}`;
    }
  }
  return null;
};

/**
 * Returns the full gradient definition (`stops` + `angleDeg`) when the
 * shape's `<p:spPr>` carries an `<a:gradFill>`. Returns `null` for any
 * other fill kind, including `inherit` — the function does not walk the
 * layout / master cascade.
 */
export const getShapeGradientFill = (shape: SlideShapeData): GradientFillOptions | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const gradFill = firstChildElement(spPr, NAME_A_GRAD_FILL);
  if (!gradFill) return null;
  const gsLst = firstChildElement(gradFill, NAME_A_GS_LST);
  if (!gsLst) return null;
  const stops: Array<{ offset: number; color: string }> = [];
  for (const c of gsLst.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml || c.name.localName !== 'gs') {
      continue;
    }
    const posRaw = getAttrValue(c, qname('', 'pos', ''));
    if (posRaw === null) continue;
    const pos = Number.parseInt(posRaw, 10);
    if (!Number.isFinite(pos)) continue;
    const color = readColorFromContainer(c);
    if (color === null) continue;
    stops.push({ offset: pos / 100_000, color });
  }
  if (stops.length === 0) return null;
  // ECMA-376 §20.1.8.33: gradFill has either <a:lin> (linear) or <a:path>
  // (radial / rectangular / shape-following) as a child to describe the
  // direction. We surface both so renderers can faithfully reproduce
  // non-linear gradients.
  let angleDeg = 0;
  const lin = firstChildElement(gradFill, NAME_A_LIN);
  if (lin) {
    const angRaw = getAttrValue(lin, qname('', 'ang', ''));
    if (angRaw !== null) {
      const ang = Number.parseInt(angRaw, 10);
      if (Number.isFinite(ang)) angleDeg = ang / 60_000;
    }
  }
  const pathEl = firstChildElement(gradFill, qname('a', 'path', NS.dml));
  if (pathEl) {
    const p = getAttrValue(pathEl, qname('', 'path', ''));
    const pathVal: 'circle' | 'rect' | 'shape' | null =
      p === 'circle' || p === 'rect' || p === 'shape' ? p : null;
    if (pathVal) {
      let focus: GradientFillOptions['focus'];
      const fillToRect = firstChildElement(pathEl, qname('a', 'fillToRect', NS.dml));
      if (fillToRect) {
        const pct = (name: string): number | undefined => {
          const v = getAttrValue(fillToRect, qname('', name, ''));
          if (v === null) return undefined;
          let n = Number.parseFloat(v);
          if (!Number.isFinite(n)) return undefined;
          if (Math.abs(n) > 1) n = n / 100000;
          return n;
        };
        const l = pct('l') ?? 0.5;
        const t = pct('t') ?? 0.5;
        const r = pct('r') ?? 0.5;
        const b = pct('b') ?? 0.5;
        focus = { left: l, top: t, right: r, bottom: b };
      }
      return { stops, angleDeg, path: pathVal, ...(focus ? { focus } : {}) };
    }
  }
  return { stops, angleDeg };
};

// ---------------------------------------------------------------------------
// @internal — used by mutation functions to write SlideData state back
// into the package and rebuild the typed view. Free functions, no class
// dependency.

const commitSlideData = (slide: SlideData): void => {
  const xml = serializeXml(slide[SLIDE_DOCUMENT]);
  const part = slide[INTERNAL_PACKAGE].getPart(slide[SLIDE_PART_NAME]);
  if (!part) throw new Error(`slide part missing: ${slide[SLIDE_PART_NAME]}`);
  part.data = encode(xml);
};

const refreshSlideData = (slide: SlideData): void => {
  const fresh = readSlidePart(slide[SLIDE_DOCUMENT].root);
  slide[SLIDE_PART] = fresh;
  const shapes = slide[SLIDE_SHAPES];
  for (let i = 0; i < shapes.length; i++) {
    const next = fresh.shapes[i];
    const existing = shapes[i];
    if (!next || !existing) continue;
    existing[SHAPE_ELEMENT] = next.element;
    existing[SHAPE_SNAPSHOT] = next;
  }
};

// Rebuild shape handles entirely — used when the shape count changes
// (e.g. removeShape). Existing SlideShapeData identities are dropped;
// SHAPE_SLIDE back-pointers stay consistent because the SlideData
// reference is preserved.
const rebuildShapesFromDocument = (slide: SlideData): void => {
  const fresh = readSlidePart(slide[SLIDE_DOCUMENT].root);
  slide[SLIDE_PART] = fresh;
  const shapes: SlideShapeData[] = [];
  for (const snap of fresh.shapes) {
    shapes.push({
      [SHAPE_SLIDE]: slide,
      [SHAPE_ELEMENT]: snap.element,
      [SHAPE_SNAPSHOT]: snap,
    });
  }
  slide[SLIDE_SHAPES] = shapes;
};

// ---------------------------------------------------------------------------
// Shape mutation — geometry.

const NAME_TX_BODY_FN = qname('p', 'txBody', NS.pml);

const requireSpPr = (shape: SlideShapeData): XmlElement => {
  const kind = shape[SHAPE_SNAPSHOT].kind;
  if (kind !== 'shape' && kind !== 'picture' && kind !== 'connector') {
    throw new Error(`fill/stroke is not supported on ${kind} shapes`);
  }
  const spPrName = qname('p', 'spPr', NS.pml);
  const el = shape[SHAPE_ELEMENT];
  let spPr = firstChildElement(el, spPrName);
  if (spPr === null) {
    spPr = { kind: 'element', name: spPrName, attrs: [], prefixDecls: new Map(), children: [] };
    el.children.push(spPr);
  }
  return spPr;
};

const requireTxBody = (shape: SlideShapeData): XmlElement => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'shape') {
    throw new Error(
      `text operations require a shape kind; ${shape[SHAPE_SNAPSHOT].kind} is not text-bearing`,
    );
  }
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY_FN);
  if (txBody === null) {
    throw new Error(`shape "${shape[SHAPE_SNAPSHOT].name}" has no <p:txBody>`);
  }
  return txBody;
};

const commitAndRefresh = (shape: SlideShapeData): void => {
  commitSlideData(shape[SHAPE_SLIDE]);
  refreshSlideData(shape[SHAPE_SLIDE]);
};

/** Sets the shape's position in EMU. Companion to `setShapeSize`. */
export const setShapePosition = (shape: SlideShapeData, x: Emu, y: Emu): void => {
  writePosition(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind, x, y);
  commitAndRefresh(shape);
};

/** Sets the shape's size in EMU. */
export const setShapeSize = (shape: SlideShapeData, w: Emu, h: Emu): void => {
  writeSize(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind, w, h);
  commitAndRefresh(shape);
};

/**
 * Sets the shape's rotation in degrees (positive clockwise). Values are
 * normalized into `[0, 360)`; pass `0` to clear an existing rotation.
 */
export const setShapeRotation = (shape: SlideShapeData, degrees: number): void => {
  writeRotation(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind, degrees);
  commitAndRefresh(shape);
};

/** Sets the shape's flip flags. Properties default to current state when omitted. */
export const setShapeFlip = (
  shape: SlideShapeData,
  options: { horizontal?: boolean; vertical?: boolean },
): void => {
  writeFlip(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind, options);
  commitAndRefresh(shape);
};

// ---------------------------------------------------------------------------
// Shape mutation — fill / stroke.

/** Sets a solid fill on the shape (color in `#RRGGBB` or scheme token). */
export const setShapeFill = (shape: SlideShapeData, color: string): void => {
  setSolidFill(requireSpPr(shape), color);
  commitAndRefresh(shape);
};

/**
 * Sets a linear gradient fill on the shape. Stops must lie in `[0, 1]`;
 * `angleDeg` defaults to `90` (top → bottom).
 *
 * Example: red → blue top-to-bottom:
 *
 *   setShapeGradientFill(shape, {
 *     stops: [{ offset: 0, color: '#FF0000' }, { offset: 1, color: '#0000FF' }],
 *     angleDeg: 90,
 *   });
 */
export const setShapeGradientFill = (shape: SlideShapeData, options: GradientFillOptions): void => {
  setGradientFill(requireSpPr(shape), options);
  commitAndRefresh(shape);
};

/**
 * Sets a preset pattern fill on the shape (e.g. `pct50`, `dkUpDiag`).
 *
 * `foreground` is the pattern stroke color; `background` fills behind
 * the pattern. Both accept `#RRGGBB`, bare `RRGGBB`, or scheme tokens
 * (`accent1`, `bg1`, ...).
 */
export const setShapePatternFill = (shape: SlideShapeData, options: PatternFillOptions): void => {
  setPatternFill(requireSpPr(shape), options);
  commitAndRefresh(shape);
};

/**
 * Reads back the pattern fill on a shape: returns the preset token
 * plus the foreground / background colors resolved against the theme.
 * Returns `null` when the shape has no `<a:pattFill>`.
 *
 * The preset string is the literal `ST_PresetPatternVal` token from
 * §20.1.10.49 — e.g. `'pct50'`, `'dkUpDiag'`, `'cross'`, `'wave'`.
 * Renderers can map it onto an SVG `<pattern>` definition.
 */
export const getShapePatternFill = (
  pres: PresentationData,
  shape: SlideShapeData,
): { preset: string; foreground: string; background: string } | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const pattFill = firstChildElement(spPr, qname('a', 'pattFill', NS.dml));
  if (!pattFill) return null;
  const preset = getAttrValue(pattFill, qname('', 'prst', '')) ?? 'pct50';
  const theme = getPresentationTheme(pres);
  const colorFrom = (parentName: string, fallback: string): string => {
    const parent = firstChildElement(pattFill, qname('a', parentName, NS.dml));
    if (!parent) return fallback;
    for (const c of parent.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      const hex = resolveDrawingColor(c, theme);
      if (hex) return hex;
    }
    return fallback;
  };
  return {
    preset,
    foreground: colorFrom('fgClr', '#000000'),
    background: colorFrom('bgClr', '#FFFFFF'),
  };
};

/**
 * Sets a picture fill on the shape, embedding `bytes` as a new media
 * part and replacing any prior fill choice on the shape's `<p:spPr>`.
 *
 * The image stretches to fill the shape (`<a:stretch><a:fillRect/>`).
 * Format is detected from magic bytes; pass `options.format` to
 * override (useful for SVG or unusual extensions).
 *
 * Throws if the format can't be detected and isn't provided explicitly,
 * or if the shape kind doesn't carry a `<p:spPr>` (e.g. groups).
 */
export const setShapeImageFill = (
  shape: SlideShapeData,
  bytes: Uint8Array,
  options: { format?: ImageFormat } = {},
): void => {
  const format = options.format ?? detectImageFormat(bytes);
  if (format === null) {
    throw new Error(
      'setShapeImageFill: could not detect image format. Pass options.format explicitly.',
    );
  }
  const contentType = contentTypeForFormat(format);
  const extension = extensionForFormat(format);
  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];

  // Allocate /ppt/media/imageN.<ext> (shared with addSlideImage's
  // numbering — both feed off the same /ppt/media space).
  let nextN = 1;
  const mediaPattern = /^\/ppt\/media\/image(\d+)\./;
  for (const p of pkg.parts) {
    const m = p.name.match(mediaPattern);
    if (m?.[1] !== undefined) {
      const n = Number.parseInt(m[1], 10);
      if (Number.isFinite(n) && n >= nextN) nextN = n + 1;
    }
  }
  const newMediaName = partName(`/ppt/media/image${nextN}.${extension}`);
  setOpcDefault(pkg, extension, contentType);
  pkg.addPart(newMediaName, contentType, bytes);

  // Slide → image rel.
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
  const newRId = nextRelId(rels.items.map((r) => r.id));
  rels.items.push({
    id: newRId,
    type: REL_TYPES.image,
    target: `../media/image${nextN}.${extension}`,
    targetMode: 'Internal',
  });
  pkg.setRels(slide[SLIDE_PART_NAME], rels);

  // Replace the shape's fill choice with <a:blipFill>.
  const spPr = requireSpPr(shape);
  const FILL_CHOICES = new Set([
    'noFill',
    'solidFill',
    'gradFill',
    'blipFill',
    'pattFill',
    'grpFill',
  ]);
  spPr.children = spPr.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        FILL_CHOICES.has(c.name.localName)
      ),
  );
  const blipName = qname('a', 'blip', NS.dml);
  const stretchName = qname('a', 'stretch', NS.dml);
  const fillRectName = qname('a', 'fillRect', NS.dml);
  const blipFillName = qname('a', 'blipFill', NS.dml);
  const blip = elem(blipName, { attrs: [attr(qname('r', 'embed', NS.officeDocRels), newRId)] });
  const stretch = elem(stretchName, { children: [elem(fillRectName)] });
  const blipFill = elem(blipFillName, { children: [blip, stretch] });
  // <a:blipFill> takes the same slot as <a:solidFill>; insert at the
  // current insertion index. We use the same heuristic as setSolidFill —
  // before <a:ln> / effectLst / scene3d / extLst.
  let insertAt = spPr.children.length;
  for (let i = 0; i < spPr.children.length; i++) {
    const c = spPr.children[i];
    if (c?.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (
      c.name.localName === 'ln' ||
      c.name.localName === 'effectLst' ||
      c.name.localName === 'effectDag' ||
      c.name.localName === 'scene3d' ||
      c.name.localName === 'sp3d' ||
      c.name.localName === 'extLst'
    ) {
      insertAt = i;
      break;
    }
  }
  spPr.children.splice(insertAt, 0, blipFill);
  commitAndRefresh(shape);
};

/** Sets `<a:noFill>` on the shape, leaving it transparent. */
export const setShapeNoFill = (shape: SlideShapeData): void => {
  setNoFillImpl(requireSpPr(shape));
  commitAndRefresh(shape);
};

/**
 * Removes any fill choice from the shape; it then inherits its fill
 * from the layout / master placeholder it descends from.
 */
export const clearShapeFill = (shape: SlideShapeData): void => {
  clearFillImpl(requireSpPr(shape));
  commitAndRefresh(shape);
};

/** Sets a solid-color outline on the shape. */
export const setShapeStroke = (
  shape: SlideShapeData,
  options: { color?: string; widthEmu?: number },
): void => {
  setSolidStroke(requireSpPr(shape), options as StrokeOptions);
  commitAndRefresh(shape);
};

/** Sets an explicit "no outline" on the shape. */
export const setShapeNoStroke = (shape: SlideShapeData): void => {
  setNoStrokeImpl(requireSpPr(shape));
  commitAndRefresh(shape);
};

/** Reads back the shape's stroke dash style, or `null` if none. */
export const getShapeStrokeDash = (shape: SlideShapeData): LineDash | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
  if (!ln) return null;
  const prstDash = firstChildElement(ln, qname('a', 'prstDash', NS.dml));
  if (!prstDash) return null;
  const v = getAttrValue(prstDash, qname('', 'val', ''));
  return (v as LineDash | null) ?? null;
};

/**
 * Reads back the shape's arrowhead on one end of `<a:ln>`, or `null`
 * when no `<a:headEnd>` / `<a:tailEnd>` is present.
 */
export const getShapeStrokeArrow = (
  shape: SlideShapeData,
  end: 'head' | 'tail',
): ArrowOptions | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
  if (!ln) return null;
  const arr = firstChildElement(ln, qname('a', end === 'head' ? 'headEnd' : 'tailEnd', NS.dml));
  if (!arr) return null;
  const type = getAttrValue(arr, qname('', 'type', ''));
  if (!type) return null;
  const width = getAttrValue(arr, qname('', 'w', ''));
  const length = getAttrValue(arr, qname('', 'len', ''));
  const result: {
    type: ArrowOptions['type'];
    width?: 'sm' | 'med' | 'lg';
    length?: 'sm' | 'med' | 'lg';
  } = {
    type: type as ArrowOptions['type'],
  };
  if (width === 'sm' || width === 'med' || width === 'lg') result.width = width;
  if (length === 'sm' || length === 'med' || length === 'lg') result.length = length;
  return result;
};

/**
 * Sets the dash pattern for the shape's outline (`<a:prstDash>`). One
 * of ECMA-376's `ST_PresetLineDashVal` tokens:
 *
 *   `'solid'` | `'dot'` | `'dash'` | `'lgDash'` | `'dashDot'` |
 *   `'lgDashDot'` | `'lgDashDotDot'` | `'sysDash'` | `'sysDot'` |
 *   `'sysDashDot'` | `'sysDashDotDot'`
 *
 * Creates `<a:ln>` if absent. Pairs naturally with `setShapeStroke`:
 * users typically set a color + width first, then the dash.
 */
export const setShapeStrokeDash = (shape: SlideShapeData, dash: LineDash): void => {
  setStrokeDash(requireSpPr(shape), dash);
  commitAndRefresh(shape);
};

/**
 * Sets an arrowhead on one end of the shape's outline.
 *
 *   - `end: 'head'` writes `<a:headEnd>` (the start of the line).
 *   - `end: 'tail'` writes `<a:tailEnd>` (the end).
 *
 * Useful primarily on connector shapes added via `addSlideLine`.
 * `type: 'none'` clears the arrowhead.
 */
export const setShapeStrokeArrow = (
  shape: SlideShapeData,
  end: 'head' | 'tail',
  options: ArrowOptions,
): void => {
  setStrokeArrow(requireSpPr(shape), end, options);
  commitAndRefresh(shape);
};

/** Removes any outline override; the shape then inherits stroke from layout. */
export const clearShapeStroke = (shape: SlideShapeData): void => {
  clearStrokeImpl(requireSpPr(shape));
  commitAndRefresh(shape);
};

// ---------------------------------------------------------------------------
// Effects: shadow + glow.

/**
 * Read-back for `setShapeShadow` / `setShapeGlow`. Returns the kind
 * of effect currently on the shape's `<a:effectLst>`, or `null` when
 * none. Decodes the configured color + numeric parameters when
 * present.
 */
export type ShapeEffect =
  | {
      readonly kind: 'shadow';
      readonly color: string;
      readonly blurEmu: number;
      readonly offsetEmu: number;
      readonly angleDeg: number;
      readonly opacity?: number;
    }
  | {
      readonly kind: 'glow';
      readonly color: string;
      readonly radiusEmu: number;
    };

/**
 * Discriminated union covering every effect in
 * `CT_EffectStyleItem` (ECMA-376 §20.1.8.x) — outer shadow, inner
 * shadow, glow, reflection, soft-edge, blur. Returned in document
 * order so renderers can chain filters with the same composition
 * PowerPoint applies.
 *
 * Lengths are EMU; angles are degrees clockwise from 3 o'clock;
 * opacity is a unit fraction (0..1) when the spec exposes one.
 */
export type ShapeEffectAny =
  | {
      readonly kind: 'outerShdw';
      readonly color: string;
      readonly opacity?: number;
      readonly blurEmu: number;
      readonly distEmu: number;
      readonly angleDeg: number;
    }
  | {
      readonly kind: 'innerShdw';
      readonly color: string;
      readonly opacity?: number;
      readonly blurEmu: number;
      readonly distEmu: number;
      readonly angleDeg: number;
    }
  | {
      readonly kind: 'glow';
      readonly color: string;
      readonly opacity?: number;
      readonly radiusEmu: number;
    }
  | {
      readonly kind: 'reflection';
      readonly opacity?: number;
      readonly blurEmu: number;
      readonly distEmu: number;
      readonly angleDeg: number;
    }
  | { readonly kind: 'softEdge'; readonly radiusEmu: number }
  | { readonly kind: 'blur'; readonly radiusEmu: number };

export const getShapeEffect = (shape: SlideShapeData): ShapeEffect | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const effectLst = firstChildElement(spPr, qname('a', 'effectLst', NS.dml));
  if (!effectLst) return null;

  const readColor = (host: XmlElement): { color: string; opacity?: number } => {
    const srgb = firstChildElement(host, qname('a', 'srgbClr', NS.dml));
    if (!srgb) return { color: '' };
    const val = getAttrValue(srgb, qname('', 'val', ''));
    const color = val !== null ? `#${val.toUpperCase()}` : '';
    const alpha = firstChildElement(srgb, qname('a', 'alpha', NS.dml));
    if (alpha) {
      const a = getAttrValue(alpha, qname('', 'val', ''));
      if (a !== null) {
        const n = Number.parseInt(a, 10);
        if (Number.isFinite(n)) return { color, opacity: n / 100000 };
      }
    }
    return { color };
  };

  const outerShdw = firstChildElement(effectLst, qname('a', 'outerShdw', NS.dml));
  if (outerShdw) {
    const blur = Number.parseInt(getAttrValue(outerShdw, qname('', 'blurRad', '')) ?? '0', 10);
    const dist = Number.parseInt(getAttrValue(outerShdw, qname('', 'dist', '')) ?? '0', 10);
    const dirRaw = Number.parseInt(getAttrValue(outerShdw, qname('', 'dir', '')) ?? '0', 10);
    const c = readColor(outerShdw);
    return {
      kind: 'shadow',
      color: c.color,
      blurEmu: blur,
      offsetEmu: dist,
      angleDeg: dirRaw / 60000,
      ...(c.opacity !== undefined ? { opacity: c.opacity } : {}),
    };
  }
  const glow = firstChildElement(effectLst, qname('a', 'glow', NS.dml));
  if (glow) {
    const rad = Number.parseInt(getAttrValue(glow, qname('', 'rad', '')) ?? '0', 10);
    const c = readColor(glow);
    return { kind: 'glow', color: c.color, radiusEmu: rad };
  }
  return null;
};

/**
 * Returns every effect attached to the shape's `<a:effectLst>` in
 * document order — outer shadow, inner shadow, glow, reflection,
 * soft edge, blur. Empty array when no effects apply.
 *
 * Companion to `getShapeEffect`, which is the v1 "first effect only"
 * helper. `getShapeEffects` is what renderers want because PowerPoint
 * composes multiple effects in a single filter (shadow + glow, etc.).
 */
// Parses an `<a:effectLst>` element into the typed effect union.
// Pulled out of `getShapeEffects` so the cascade-aware variant can
// reuse it.
const parseEffectLst = (
  effectLst: XmlElement,
  theme: PresentationTheme | null,
): ShapeEffectAny[] => {
  const readEffectColor = (host: XmlElement): { color: string; opacity?: number } => {
    let inner: XmlElement | null = null;
    for (const c of host.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      if (
        c.name.localName === 'srgbClr' ||
        c.name.localName === 'schemeClr' ||
        c.name.localName === 'sysClr' ||
        c.name.localName === 'prstClr'
      ) {
        inner = c;
        break;
      }
    }
    if (!inner) return { color: '' };
    let opacity: number | undefined;
    const alphaEl = firstChildElement(inner, qname('a', 'alpha', NS.dml));
    if (alphaEl) {
      const a = getAttrValue(alphaEl, qname('', 'val', ''));
      if (a !== null) {
        let n = Number.parseFloat(a);
        if (Number.isFinite(n)) {
          if (Math.abs(n) > 1) n = n / 100000;
          opacity = n;
        }
      }
    }
    const hex = resolveDrawingColor(inner, theme);
    return { color: hex ?? '', ...(opacity !== undefined ? { opacity } : {}) };
  };

  const out: ShapeEffectAny[] = [];
  for (const child of effectLst.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.dml) continue;
    const local = child.name.localName;
    if (local === 'outerShdw' || local === 'innerShdw') {
      const blur = Number.parseInt(getAttrValue(child, qname('', 'blurRad', '')) ?? '0', 10) || 0;
      const dist = Number.parseInt(getAttrValue(child, qname('', 'dist', '')) ?? '0', 10) || 0;
      const dir = Number.parseInt(getAttrValue(child, qname('', 'dir', '')) ?? '0', 10) || 0;
      const c = readEffectColor(child);
      out.push({
        kind: local,
        color: c.color,
        blurEmu: blur,
        distEmu: dist,
        angleDeg: dir / 60000,
        ...(c.opacity !== undefined ? { opacity: c.opacity } : {}),
      });
    } else if (local === 'glow') {
      const rad = Number.parseInt(getAttrValue(child, qname('', 'rad', '')) ?? '0', 10) || 0;
      const c = readEffectColor(child);
      out.push({
        kind: 'glow',
        color: c.color,
        radiusEmu: rad,
        ...(c.opacity !== undefined ? { opacity: c.opacity } : {}),
      });
    } else if (local === 'reflection') {
      const blur = Number.parseInt(getAttrValue(child, qname('', 'blurRad', '')) ?? '0', 10) || 0;
      const dist = Number.parseInt(getAttrValue(child, qname('', 'dist', '')) ?? '0', 10) || 0;
      const dir = Number.parseInt(getAttrValue(child, qname('', 'dir', '')) ?? '0', 10) || 0;
      const endA = getAttrValue(child, qname('', 'endA', ''));
      let opacity: number | undefined;
      if (endA !== null) {
        let n = Number.parseFloat(endA);
        if (Number.isFinite(n)) {
          if (Math.abs(n) > 1) n = n / 100000;
          opacity = n;
        }
      }
      out.push({
        kind: 'reflection',
        blurEmu: blur,
        distEmu: dist,
        angleDeg: dir / 60000,
        ...(opacity !== undefined ? { opacity } : {}),
      });
    } else if (local === 'softEdge') {
      const rad = Number.parseInt(getAttrValue(child, qname('', 'rad', '')) ?? '0', 10) || 0;
      out.push({ kind: 'softEdge', radiusEmu: rad });
    } else if (local === 'blur') {
      const rad = Number.parseInt(getAttrValue(child, qname('', 'rad', '')) ?? '0', 10) || 0;
      out.push({ kind: 'blur', radiusEmu: rad });
    }
  }
  return out;
};

export const getShapeEffects = (
  pres: PresentationData,
  shape: SlideShapeData,
): readonly ShapeEffectAny[] => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return [];
  const effectLst = firstChildElement(spPr, qname('a', 'effectLst', NS.dml));
  if (!effectLst) return [];
  return parseEffectLst(effectLst, getPresentationTheme(pres));
};

/**
 * Same as `getShapeEffects` but walks the layout → master placeholder
 * cascade when the shape itself has no `<a:effectLst>`. Inherits
 * "all or nothing" — once any layer supplies an effect list, that
 * list is used; layers further down aren't merged in. This matches
 * PowerPoint's behaviour (effect lists override rather than compose).
 */
export const getShapeEffectsEffective = (
  pres: PresentationData,
  shape: SlideShapeData,
): readonly ShapeEffectAny[] => {
  const own = getShapeEffects(pres, shape);
  if (own.length > 0) return own;

  const phIdx = getShapePlaceholderIdx(shape);
  const phType = getShapePlaceholderType(shape);
  if (phIdx === null && phType === null) return own;

  const theme = getPresentationTheme(pres);
  const layout = getSlideLayout(shape[SHAPE_SLIDE]);
  if (!layout) return own;

  const findPh = (
    shapes: ReadonlyArray<{
      placeholderIdx: number | null;
      placeholderType: string | null;
      element: XmlElement;
    }>,
  ): XmlElement | null => {
    let match = phIdx !== null ? shapes.find((s) => s.placeholderIdx === phIdx) : undefined;
    if (!match && phType !== null) match = shapes.find((s) => s.placeholderType === phType);
    return match?.element ?? null;
  };

  const readEffectsOn = (el: XmlElement): readonly ShapeEffectAny[] => {
    const spPr = firstChildElement(el, qname('p', 'spPr', NS.pml));
    if (!spPr) return [];
    const eff = firstChildElement(spPr, qname('a', 'effectLst', NS.dml));
    if (!eff) return [];
    return parseEffectLst(eff, theme);
  };

  const layoutPh = findPh(layout[LAYOUT_PART].shapes);
  if (layoutPh) {
    const layoutEffects = readEffectsOn(layoutPh);
    if (layoutEffects.length > 0) return layoutEffects;
  }

  const pkg = pres[INTERNAL_PACKAGE];
  const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
  const layoutRels = pkg.getRels(layoutPartName);
  if (!layoutRels) return own;
  const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
  if (!masterRel) return own;
  const masterPart = pkg.getPart(resolveTarget(layoutPartName, masterRel.target));
  if (!masterPart) return own;
  const masterRoot = parseXml(decode(masterPart.data)).root;
  const { shapes: masterShapes } = readShapeTreeFromCsldRoot(masterRoot, 'sldMaster');
  const masterPh = findPh(masterShapes);
  if (masterPh) {
    const masterEffects = readEffectsOn(masterPh);
    if (masterEffects.length > 0) return masterEffects;
  }
  return own;
};

/**
 * Sets an outer drop shadow on the shape. Defaults: black, 4pt blur,
 * 3pt offset, 45° (down-right). Pass `opacity` (0–1) to soften the
 * shadow.
 */
export const setShapeShadow = (shape: SlideShapeData, options: ShadowOptions = {}): void => {
  setShadow(requireSpPr(shape), options);
  commitAndRefresh(shape);
};

/**
 * Sets a glow around the shape. The radius is in EMU (default 5pt =
 * 63500). Mutually exclusive with `setShapeShadow` in v1 — calling
 * either replaces the prior `<a:effectLst>` entirely.
 */
export const setShapeGlow = (shape: SlideShapeData, options: GlowOptions): void => {
  setGlow(requireSpPr(shape), options);
  commitAndRefresh(shape);
};

/** Removes any effects (shadow / glow / future presets) from the shape. */
export const clearShapeEffects = (shape: SlideShapeData): void => {
  clearEffectsImpl(requireSpPr(shape));
  commitAndRefresh(shape);
};

// ---------------------------------------------------------------------------
// Shape mutation — text.

/**
 * Replaces the shape's visible text with `value`. Newlines start a new
 * paragraph. Existing run/paragraph properties are preserved so font,
 * color, size, alignment, and bullet style stay intact.
 */
export const setShapeText = (
  shape: SlideShapeData,
  value: string,
  options: { bullets?: BulletStyle } = {},
): void => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'shape') {
    throw new Error(
      `setShapeText only works on text-bearing shapes; ${shape[SHAPE_SNAPSHOT].kind} is not one`,
    );
  }
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY_FN);
  if (txBody === null) {
    throw new Error(`shape "${shape[SHAPE_SNAPSHOT].name}" has no <p:txBody>`);
  }
  setTextBody(txBody, value);
  if (options.bullets !== undefined) {
    applyBulletToAllParagraphs(txBody, options.bullets);
  }
  commitAndRefresh(shape);
};

/**
 * Appends `value` to the shape's existing text on a new line. The
 * shape's existing run / paragraph formatting is preserved by
 * `setTextBody`; the new paragraph inherits the same template.
 *
 * Equivalent to `setShapeText(shape, getShapeText(shape) + '\n' + value)`,
 * minus the leading newline when there was no existing text.
 */
export const appendShapeText = (shape: SlideShapeData, value: string): void => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'shape') {
    throw new Error(
      `appendShapeText only works on text-bearing shapes; ${shape[SHAPE_SNAPSHOT].kind} is not one`,
    );
  }
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY_FN);
  if (txBody === null) {
    throw new Error(`shape "${shape[SHAPE_SNAPSHOT].name}" has no <p:txBody>`);
  }
  const existing = shape[SHAPE_SNAPSHOT].text;
  const combined = existing.length === 0 ? value : `${existing}\n${value}`;
  setTextBody(txBody, combined);
  commitAndRefresh(shape);
};

/**
 * Sets the vertical text anchor on the shape's text body
 * (`<a:bodyPr anchor="..."/>`). Choices map to ECMA-376 tokens:
 *
 *   - `'top'`    → `anchor="t"`
 *   - `'center'` → `anchor="ctr"`
 *   - `'bottom'` → `anchor="b"`
 *
 * The bodyPr is created if absent. Throws for non-text-bearing shape
 * kinds.
 */
export type TextAnchor = 'top' | 'center' | 'bottom';

const NAME_A_BODY_PR = qname('a', 'bodyPr', NS.dml);

/**
 * Word wrap mode on a text body. `'square'` (PowerPoint default for
 * textboxes) wraps lines at the shape's width; `'none'` lets text
 * overflow horizontally.
 */
export type TextWrap = 'none' | 'square';

/** Auto-fit mode on a text body. */
export type TextAutoFit =
  | 'none' // <a:noAutofit/>
  | 'normal' // <a:normAutofit/> — shrink text to fit
  | 'shape'; // <a:spAutoFit/> — resize shape to fit text

const AUTO_FIT_LOCALS = new Set(['noAutofit', 'normAutofit', 'spAutoFit']);

const requireBodyPr = (shape: SlideShapeData): XmlElement => {
  const txBody = requireTxBody(shape);
  let bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (bodyPr === null) {
    bodyPr = elem(NAME_A_BODY_PR);
    txBody.children.unshift(bodyPr);
  }
  return bodyPr;
};

/**
 * Sets the text-body word-wrap mode.
 *
 *   - `'square'` writes `wrap="square"` — PowerPoint default for textboxes.
 *   - `'none'`   writes `wrap="none"`  — text can overflow horizontally.
 *
 * Throws for non-text-bearing shape kinds.
 */
export const setShapeTextWrap = (shape: SlideShapeData, wrap: TextWrap): void => {
  const bodyPr = requireBodyPr(shape);
  const ATTR_WRAP = qname('', 'wrap', '');
  bodyPr.attrs = bodyPr.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'wrap'),
  );
  bodyPr.attrs.push(attr(ATTR_WRAP, wrap));
  commitAndRefresh(shape);
};

/** Reads back the bodyPr `wrap` attribute, or `null` when absent. */
export const getShapeTextWrap = (shape: SlideShapeData): TextWrap | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY_FN);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  const v = getAttrValue(bodyPr, qname('', 'wrap', ''));
  if (v === 'none' || v === 'square') return v;
  return null;
};

/**
 * Sets the text-body auto-fit mode:
 *
 *   - `'none'`   → `<a:noAutofit/>`
 *   - `'normal'` → `<a:normAutofit/>`   shrink text to fit the shape
 *   - `'shape'`  → `<a:spAutoFit/>`     grow the shape to fit text
 *
 * Replaces any prior auto-fit child on `<a:bodyPr>`. Throws for
 * non-text-bearing shape kinds.
 */
export const setShapeTextAutoFit = (shape: SlideShapeData, mode: TextAutoFit): void => {
  const bodyPr = requireBodyPr(shape);
  bodyPr.children = bodyPr.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        AUTO_FIT_LOCALS.has(c.name.localName)
      ),
  );
  const local = mode === 'none' ? 'noAutofit' : mode === 'normal' ? 'normAutofit' : 'spAutoFit';
  bodyPr.children.push(elem(qname('a', local, NS.dml)));
  commitAndRefresh(shape);
};

/**
 * Reads back the bodyPr auto-fit child, or `null` when none is
 * present (PowerPoint applies a layout-inherited default in that case).
 */
export const getShapeTextAutoFit = (shape: SlideShapeData): TextAutoFit | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY_FN);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  for (const c of bodyPr.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (c.name.localName === 'noAutofit') return 'none';
    if (c.name.localName === 'normAutofit') return 'normal';
    if (c.name.localName === 'spAutoFit') return 'shape';
  }
  return null;
};

/**
 * Reads the scale parameters PowerPoint stores on `<a:normAutofit>`
 * once it has shrunk a text body to fit. Returns `null` if the body
 * doesn't carry `<a:normAutofit>` or the attributes are absent. Both
 * fields are unitless ratios in `[0, 1]`:
 *
 *   - `fontScale`     — multiply every run's font size by this. Default `1`.
 *   - `lnSpcReduction` — subtract from the line-height ratio. Default `0`.
 *
 * Companion to `getShapeTextAutoFit`. Renderers that want to match
 * PowerPoint's actual on-screen text size apply these factors to the
 * authored font sizes; without them, every long title overflows.
 */
export const getShapeTextAutoFitParams = (
  shape: SlideShapeData,
): { fontScale: number; lnSpcReduction: number } | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY_FN);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  for (const c of bodyPr.children) {
    if (
      c.kind === 'element' &&
      c.name.namespaceURI === NS.dml &&
      c.name.localName === 'normAutofit'
    ) {
      const fsRaw = getAttrValue(c, qname('', 'fontScale', ''));
      const lsRaw = getAttrValue(c, qname('', 'lnSpcReduction', ''));
      const fs = fsRaw === null ? 100_000 : Number.parseInt(fsRaw, 10);
      const ls = lsRaw === null ? 0 : Number.parseInt(lsRaw, 10);
      return {
        fontScale: Number.isFinite(fs) ? fs / 100_000 : 1,
        lnSpcReduction: Number.isFinite(ls) ? ls / 100_000 : 0,
      };
    }
  }
  return null;
};

/**
 * Reads back the vertical text anchor on the shape's `<a:bodyPr>`.
 * Maps the ECMA-376 tokens back to the public union:
 *
 *   `'t'` → `'top'`, `'ctr'` → `'center'`, `'b'` → `'bottom'`
 *
 * Returns `null` when the bodyPr is absent or has no anchor attribute.
 */
export const getShapeTextAnchor = (shape: SlideShapeData): TextAnchor | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY_FN);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  const v = getAttrValue(bodyPr, qname('', 'anchor', ''));
  if (v === 't') return 'top';
  if (v === 'ctr') return 'center';
  if (v === 'b') return 'bottom';
  return null;
};

/**
 * Reads back the internal margins of the shape's text frame. Sides
 * that are absent in the XML default to `null` (PowerPoint applies
 * its built-in default for the missing side).
 */
/**
 * Reads the multi-column layout on a text body — `<a:bodyPr
 * numCol="N" spcCol="EMU"/>`. Returns `null` when columns aren't
 * configured (the default single column). `gapEmu` is the
 * inter-column gap in EMU; omitted when `<a:bodyPr>` has no
 * `spcCol` attribute.
 */
export const getShapeTextColumns = (
  shape: SlideShapeData,
): { count: number; gapEmu?: number } | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY_FN);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  const numColRaw = getAttrValue(bodyPr, qname('', 'numCol', ''));
  if (numColRaw === null) return null;
  const count = Number.parseInt(numColRaw, 10);
  if (!Number.isFinite(count) || count < 2) return null;
  const gapRaw = getAttrValue(bodyPr, qname('', 'spcCol', ''));
  if (gapRaw !== null) {
    const g = Number.parseInt(gapRaw, 10);
    if (Number.isFinite(g)) return { count, gapEmu: g };
  }
  return { count };
};

/**
 * Reads the shape's text-direction token from `<a:bodyPr vert="…"/>`.
 * Per ECMA-376 §17.18.93 `ST_TextVerticalType`:
 *
 *   - `horz` — default left-to-right, top-to-bottom (returns `null`).
 *   - `vert` — 90° rotation, lines run top-to-bottom, columns right-to-left.
 *   - `vert270` — 270° rotation, lines top-to-bottom, columns left-to-right.
 *   - `wordArtVert` — characters not rotated, stacked vertically.
 *   - `eaVert` — East-Asian vertical: characters upright, columns right-to-left.
 *   - `mongolianVert` — Mongolian: rotated 90°, columns left-to-right.
 *   - `wordArtVertRtl` — RTL word-art stacked vertically.
 *
 * Returns `null` when the attribute is absent or set to the default
 * `horz`.
 */
export const getShapeTextDirection = (
  shape: SlideShapeData,
): 'vert' | 'vert270' | 'wordArtVert' | 'eaVert' | 'mongolianVert' | 'wordArtVertRtl' | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY_FN);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  const v = getAttrValue(bodyPr, qname('', 'vert', ''));
  if (
    v === 'vert' ||
    v === 'vert270' ||
    v === 'wordArtVert' ||
    v === 'eaVert' ||
    v === 'mongolianVert' ||
    v === 'wordArtVertRtl'
  )
    return v;
  return null;
};

export const getShapeTextMargins = (
  shape: SlideShapeData,
): {
  readonly left: number | null;
  readonly top: number | null;
  readonly right: number | null;
  readonly bottom: number | null;
} | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY_FN);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  const readSide = (local: string): number | null => {
    const v = getAttrValue(bodyPr, qname('', local, ''));
    if (v === null) return null;
    const n = Number.parseInt(v, 10);
    return Number.isFinite(n) ? n : null;
  };
  return {
    left: readSide('lIns'),
    top: readSide('tIns'),
    right: readSide('rIns'),
    bottom: readSide('bIns'),
  };
};

/**
 * Resolves the effective `<a:bodyPr>` properties — anchor, wrap, vertical
 * direction, and inset margins — by walking the layout / master cascade
 * the same way `getShapeRunFormatEffective` walks rPr. Returns the
 * innermost value that the cascade supplies, or `null` for properties
 * neither the shape nor any inherited placeholder authors.
 *
 * Companion to `getShapeTextAnchor` / `getShapeTextWrap` /
 * `getShapeTextDirection` / `getShapeTextMargins`, which only report the
 * literal value on the shape itself.
 */
export const getShapeBodyPrEffective = (
  pres: PresentationData,
  shape: SlideShapeData,
): {
  anchor: TextAnchor | null;
  wrap: TextWrap | null;
  vert: ReturnType<typeof getShapeTextDirection>;
  margins: { left: number | null; top: number | null; right: number | null; bottom: number | null };
} => {
  const result = {
    anchor: null as TextAnchor | null,
    wrap: null as TextWrap | null,
    vert: null as ReturnType<typeof getShapeTextDirection>,
    margins: {
      left: null as number | null,
      top: null as number | null,
      right: null as number | null,
      bottom: null as number | null,
    },
  };
  const parseBodyPr = (bodyPr: XmlElement): void => {
    if (result.anchor === null) {
      const a = getAttrValue(bodyPr, qname('', 'anchor', ''));
      if (a === 't') result.anchor = 'top';
      else if (a === 'ctr') result.anchor = 'center';
      else if (a === 'b') result.anchor = 'bottom';
    }
    if (result.wrap === null) {
      const w = getAttrValue(bodyPr, qname('', 'wrap', ''));
      if (w === 'square') result.wrap = 'square';
      else if (w === 'none') result.wrap = 'none';
    }
    if (result.vert === null) {
      const v = getAttrValue(bodyPr, qname('', 'vert', ''));
      if (
        v === 'vert' ||
        v === 'vert270' ||
        v === 'wordArtVert' ||
        v === 'eaVert' ||
        v === 'mongolianVert' ||
        v === 'wordArtVertRtl'
      )
        result.vert = v;
    }
    for (const side of ['l', 't', 'r', 'b'] as const) {
      const target =
        side === 'l' ? 'left' : side === 't' ? 'top' : side === 'r' ? 'right' : 'bottom';
      if (result.margins[target] !== null) continue;
      const v = getAttrValue(bodyPr, qname('', `${side}Ins`, ''));
      if (v === null) continue;
      const n = Number.parseInt(v, 10);
      if (Number.isFinite(n)) result.margins[target] = n;
    }
  };

  // 1. The shape's own bodyPr.
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY_FN);
  if (txBody) {
    const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
    if (bodyPr) parseBodyPr(bodyPr);
  }

  // 2-3. Walk layout placeholder and master placeholder bodyPr.
  const phIdx = getShapePlaceholderIdx(shape);
  const phType = getShapePlaceholderType(shape);
  const slide = shape[SHAPE_SLIDE];
  const layout = getSlideLayout(slide);
  if (!layout) return result;

  const findPh = (
    shapes: ReadonlyArray<{
      placeholderIdx: number | null;
      placeholderType: string | null;
      element: XmlElement;
    }>,
  ): XmlElement | null => {
    let match = phIdx !== null ? shapes.find((s) => s.placeholderIdx === phIdx) : undefined;
    if (!match && phType !== null) {
      match = shapes.find((s) => s.placeholderType === phType);
    }
    return match?.element ?? null;
  };

  const layoutPhEl = findPh(layout[LAYOUT_PART].shapes);
  if (layoutPhEl) {
    const layoutTxBody = firstChildElement(layoutPhEl, NAME_TX_BODY_FN);
    if (layoutTxBody) {
      const bodyPr = firstChildElement(layoutTxBody, NAME_A_BODY_PR);
      if (bodyPr) parseBodyPr(bodyPr);
    }
  }

  const pkg = pres[INTERNAL_PACKAGE];
  const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
  const layoutRels = pkg.getRels(layoutPartName);
  if (!layoutRels) return result;
  const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
  if (!masterRel) return result;
  const masterPart = pkg.getPart(resolveTarget(layoutPartName, masterRel.target));
  if (!masterPart) return result;
  const masterRoot = parseXml(decode(masterPart.data)).root;
  const { shapes: masterShapes } = readShapeTreeFromCsldRoot(masterRoot, 'sldMaster');
  const masterPhEl = findPh(masterShapes);
  if (masterPhEl) {
    const masterTxBody = firstChildElement(masterPhEl, NAME_TX_BODY_FN);
    if (masterTxBody) {
      const bodyPr = firstChildElement(masterTxBody, NAME_A_BODY_PR);
      if (bodyPr) parseBodyPr(bodyPr);
    }
  }
  return result;
};

export const setShapeTextAnchor = (shape: SlideShapeData, anchor: TextAnchor): void => {
  const txBody = requireTxBody(shape);
  let bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (bodyPr === null) {
    bodyPr = elem(NAME_A_BODY_PR);
    txBody.children.unshift(bodyPr);
  }
  const token = anchor === 'top' ? 't' : anchor === 'center' ? 'ctr' : 'b';
  const ATTR_ANCHOR = qname('', 'anchor', '');
  // Replace any existing anchor attribute.
  bodyPr.attrs = bodyPr.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'anchor'),
  );
  bodyPr.attrs.push(attr(ATTR_ANCHOR, token));
  commitAndRefresh(shape);
};

/**
 * Sets the internal margins of the shape's text frame in EMU. Each
 * side is independent; omitted sides keep their current value (or the
 * layout-inherited default when the attribute is absent).
 *
 * PowerPoint's defaults for a textbox: left/right 91440 (0.1in),
 * top/bottom 45720 (0.05in).
 *
 *   setShapeTextMargins(shape, { left: 0, right: 0 }); // flush-left text
 */
export const setShapeTextMargins = (
  shape: SlideShapeData,
  margins: { left?: number; top?: number; right?: number; bottom?: number },
): void => {
  const txBody = requireTxBody(shape);
  let bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (bodyPr === null) {
    bodyPr = elem(NAME_A_BODY_PR);
    txBody.children.unshift(bodyPr);
  }
  const writes: Array<{ name: string; value: number }> = [];
  if (margins.left !== undefined) writes.push({ name: 'lIns', value: margins.left });
  if (margins.top !== undefined) writes.push({ name: 'tIns', value: margins.top });
  if (margins.right !== undefined) writes.push({ name: 'rIns', value: margins.right });
  if (margins.bottom !== undefined) writes.push({ name: 'bIns', value: margins.bottom });

  const localsToClear = new Set(writes.map((w) => w.name));
  bodyPr.attrs = bodyPr.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && localsToClear.has(a.name.localName)),
  );
  for (const w of writes) {
    bodyPr.attrs.push(attr(qname('', w.name, ''), String(Math.round(w.value))));
  }
  commitAndRefresh(shape);
};

/** Sets the bullet style on every paragraph in the shape's text body. */
export const setShapeBullets = (shape: SlideShapeData, style: BulletStyle): void => {
  applyBulletToAllParagraphs(requireTxBody(shape), style);
  commitAndRefresh(shape);
};

/** Sets the horizontal alignment of every paragraph in the shape's text. */
export const setShapeAlignment = (shape: SlideShapeData, align: ParagraphAlignment): void => {
  applyAlignmentToAllParagraphs(requireTxBody(shape), align);
  commitAndRefresh(shape);
};

/**
 * Applies `format` to every run in the shape's text. Run-property
 * attributes not addressed by `format` are preserved, so partial
 * updates compose.
 */
export const setShapeTextFormat = (shape: SlideShapeData, format: TextFormat): void => {
  applyFormatToAllRuns(requireTxBody(shape), format);
  commitAndRefresh(shape);
};

// ---------------------------------------------------------------------------
// Per-run text accessors.
//
// Lets callers reach into a shape's text body to read or format a
// specific paragraph or run. `applyFormatToAllRuns` covers the bulk-edit
// case; these helpers cover "make this one word red."

const NAME_A_P = qname('a', 'p', NS.dml);
const NAME_A_R = qname('a', 'r', NS.dml);
const NAME_A_RPR = qname('a', 'rPr', NS.dml);
const NAME_A_T = qname('a', 't', NS.dml);

const paragraphsOf = (txBody: XmlElement): XmlElement[] =>
  txBody.children.filter(
    (c): c is XmlElement =>
      c.kind === 'element' &&
      c.name.namespaceURI === NAME_A_P.namespaceURI &&
      c.name.localName === 'p',
  );

const runsOf = (paragraph: XmlElement): XmlElement[] =>
  paragraph.children.filter(
    (c): c is XmlElement =>
      c.kind === 'element' &&
      c.name.namespaceURI === NAME_A_R.namespaceURI &&
      c.name.localName === 'r',
  );

const requireParagraph = (shape: SlideShapeData, paragraphIndex: number): XmlElement => {
  const txBody = requireTxBody(shape);
  const paragraphs = paragraphsOf(txBody);
  const paragraph = paragraphs[paragraphIndex];
  if (!paragraph) {
    throw new RangeError(
      `paragraph index ${paragraphIndex} out of range (have ${paragraphs.length})`,
    );
  }
  return paragraph;
};

const requireRun = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
): XmlElement => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const runs = runsOf(paragraph);
  const run = runs[runIndex];
  if (!run) {
    throw new RangeError(
      `run index ${runIndex} out of range in paragraph ${paragraphIndex} (have ${runs.length})`,
    );
  }
  return run;
};

const ensureRPr = (run: XmlElement): XmlElement => {
  const existing = firstChildElement(run, NAME_A_RPR);
  if (existing !== null) return existing;
  // `<a:rPr>` is the first child of `<a:r>` per the schema.
  const fresh = elem(NAME_A_RPR);
  run.children.unshift(fresh);
  return fresh;
};

const readRunText = (run: XmlElement): string => {
  const tEl = firstChildElement(run, NAME_A_T);
  if (tEl === null) return '';
  let out = '';
  for (const child of tEl.children) {
    if (child.kind === 'text' || child.kind === 'cdata') out += child.data;
  }
  return out;
};

const writeRunText = (run: XmlElement, value: string): void => {
  let tEl = firstChildElement(run, NAME_A_T);
  if (tEl === null) {
    tEl = elem(NAME_A_T);
    run.children.push(tEl);
  }
  tEl.children = [{ kind: 'text', data: value }];
};

/** Number of paragraphs in the shape's text body. Throws for non-text shapes. */
export const getShapeParagraphCount = (shape: SlideShapeData): number =>
  paragraphsOf(requireTxBody(shape)).length;

/**
 * One inline element in a paragraph as ordered: a literal text run
 * (`<a:r>`), a field substitution (`<a:fld>` — slide number, date, etc.),
 * or a line break (`<a:br>`). Renderers walk this list instead of the
 * strict `<a:r>`-only `getShapeRunCount` / `getShapeRunText` pair when
 * they need to reproduce the paragraph's full visible content.
 *
 * `text` is the cached value (`<a:t>` content for `r` and `fld`; `''`
 * for `br`). `format` is the literal `<a:rPr>` on the element when
 * present; use `getShapeRunFormatEffective` to walk inheritance.
 *
 * Field kinds (`fld.type`): typical ECMA-376 `ST_TextFieldType` tokens
 * are `slidenum`, `datetime` (variants `1`..`13`), `presentationDate`,
 * `headerfooter`, `footer`, etc. Unrecognised tokens come through
 * unchanged so renderers can decide whether to substitute live values.
 */
export type ShapeParagraphElement =
  | { readonly kind: 'r'; readonly text: string; readonly format: TextFormat | null }
  | {
      readonly kind: 'fld';
      readonly text: string;
      readonly format: TextFormat | null;
      readonly type: string | null;
    }
  | { readonly kind: 'br'; readonly format: TextFormat | null };

/**
 * Returns the inline children of a paragraph in document order — runs,
 * field placeholders, and line breaks. Used by renderers that need to
 * reproduce the paragraph faithfully (the `<a:r>`-only run accessors
 * silently drop fields and breaks).
 */
export const getShapeParagraphElements = (
  shape: SlideShapeData,
  paragraphIndex: number,
): ReadonlyArray<ShapeParagraphElement> => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const out: ShapeParagraphElement[] = [];
  const readT = (parent: XmlElement): string => {
    const tEl = firstChildElement(parent, NAME_A_T);
    if (!tEl) return '';
    let acc = '';
    for (const c of tEl.children) {
      if (c.kind === 'text' || c.kind === 'cdata') acc += c.data;
    }
    return acc;
  };
  const readFmt = (parent: XmlElement): TextFormat | null => {
    const rPr = firstChildElement(parent, NAME_A_RPR);
    if (!rPr) return null;
    return parseRPrLikeElement(rPr) as TextFormat;
  };
  for (const child of paragraph.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.dml) continue;
    if (child.name.localName === 'r') {
      out.push({ kind: 'r', text: readT(child), format: readFmt(child) });
    } else if (child.name.localName === 'fld') {
      const type = getAttrValue(child, qname('', 'type', ''));
      out.push({ kind: 'fld', text: readT(child), format: readFmt(child), type });
    } else if (child.name.localName === 'br') {
      out.push({ kind: 'br', format: readFmt(child) });
    }
  }
  return out;
};

/**
 * Number of text runs in the given paragraph. Throws on out-of-range
 * paragraph index or non-text shapes.
 */
export const getShapeRunCount = (shape: SlideShapeData, paragraphIndex: number): number =>
  runsOf(requireParagraph(shape, paragraphIndex)).length;

/** Visible text of a single run. */
export const getShapeRunText = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
): string => readRunText(requireRun(shape, paragraphIndex, runIndex));

/**
 * Sets `<a:hlinkClick>` on a single run. Per-run counterpart to
 * `setShapeHyperlink` (which targets every run in the shape). Pass
 * `null` to clear the link on that run alone — other runs are
 * untouched. Allocates or reuses a hyperlink rel on the slide
 * exactly like the shape-level setter.
 */
export const setShapeRunHyperlink = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
  url: string | null,
): void => {
  const run = requireRun(shape, paragraphIndex, runIndex);
  let rPr = firstChildElement(run, qname('a', 'rPr', NS.dml));
  if (rPr === null) {
    rPr = elem(qname('a', 'rPr', NS.dml));
    run.children.unshift(rPr);
  }
  rPr.children = rPr.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        c.name.localName === 'hlinkClick'
      ),
  );
  if (url !== null) {
    const slide = shape[SHAPE_SLIDE];
    const pkg = slide[INTERNAL_PACKAGE];
    const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
    const existing = rels.items.find(
      (r) => r.type === REL_TYPES.hyperlink && r.target === url && r.targetMode === 'External',
    );
    let rId: string;
    if (existing) {
      rId = existing.id;
    } else {
      rId = nextRelId(rels.items.map((r) => r.id));
      rels.items.push({
        id: rId,
        type: REL_TYPES.hyperlink,
        target: url,
        targetMode: 'External',
      });
      pkg.setRels(slide[SLIDE_PART_NAME], rels);
    }
    rPr.children.push(
      elem(qname('a', 'hlinkClick', NS.dml), {
        attrs: [attr(qname('r', 'id', NS.officeDocRels), rId)],
      }),
    );
  }
  commitAndRefresh(shape);
};

/**
 * Reads the external URL on a single run's `<a:hlinkClick>`. Per-run
 * counterpart to `getShapeHyperlink` (which only surfaces the first
 * link it finds). Returns `null` when this run has no link, or the
 * link's `r:id` resolves to a non-hyperlink rel.
 */
export const getShapeRunHyperlink = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
): string | null => {
  const run = requireRun(shape, paragraphIndex, runIndex);
  const rPr = firstChildElement(run, qname('a', 'rPr', NS.dml));
  if (!rPr) return null;
  const hlink = firstChildElement(rPr, qname('a', 'hlinkClick', NS.dml));
  if (!hlink) return null;
  const rId = getAttrValue(hlink, qname('r', 'id', NS.officeDocRels));
  if (!rId) return null;
  const slide = shape[SHAPE_SLIDE];
  const rels = slide[INTERNAL_PACKAGE].getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return null;
  const rel = rels.items.find((x) => x.id === rId);
  if (rel?.type === REL_TYPES.hyperlink && rel.targetMode === 'External') return rel.target;
  return null;
};

const NAME_A_PPR = qname('a', 'pPr', NS.dml);
const ATTR_LVL = qname('', 'lvl', '');
const ATTR_ALGN_FN = qname('', 'algn', '');

const ensurePPr = (paragraph: XmlElement): XmlElement => {
  const existing = firstChildElement(paragraph, NAME_A_PPR);
  if (existing !== null) return existing;
  const fresh = elem(NAME_A_PPR);
  // <a:pPr> must be the first child of <a:p>.
  paragraph.children.unshift(fresh);
  return fresh;
};

const alignTokenForFn = (a: ParagraphAlignment): string => {
  switch (a) {
    case 'left':
    case 'l':
      return 'l';
    case 'center':
    case 'ctr':
      return 'ctr';
    case 'right':
    case 'r':
      return 'r';
    case 'justify':
    case 'just':
      return 'just';
    case 'distribute':
    case 'dist':
      return 'dist';
    default:
      return a;
  }
};

/**
 * Sets the horizontal alignment of a single paragraph. Same token set
 * as `setShapeAlignment`. Other paragraphs are untouched.
 */
export const setParagraphAlignment = (
  shape: SlideShapeData,
  paragraphIndex: number,
  align: ParagraphAlignment,
): void => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = ensurePPr(paragraph);
  pPr.attrs = pPr.attrs.filter((a) => a.name.localName !== 'algn');
  pPr.attrs.push(attr(ATTR_ALGN_FN, alignTokenForFn(align)));
  commitAndRefresh(shape);
};

/**
 * Sets the paragraph's nesting level (`<a:pPr lvl="N"/>`). Levels are
 * 0-indexed; PowerPoint accepts 0 through 8. Pass `0` to clear an
 * existing level — `<a:pPr lvl="0"/>` is the same as omitting the attr.
 *
 * Used in tandem with bullets to author nested lists:
 *
 *   setShapeText(shape, 'Item 1\nNested\nItem 2');
 *   setShapeBullets(shape, 'bullet');
 *   setParagraphLevel(shape, 1, 1);  // indent the second line
 */
export const setParagraphLevel = (
  shape: SlideShapeData,
  paragraphIndex: number,
  level: number,
): void => {
  if (!Number.isInteger(level) || level < 0 || level > 8) {
    throw new RangeError(`paragraph level must be an integer in [0, 8], got ${level}`);
  }
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = ensurePPr(paragraph);
  pPr.attrs = pPr.attrs.filter((a) => a.name.localName !== 'lvl');
  if (level > 0) pPr.attrs.push(attr(ATTR_LVL, String(level)));
  commitAndRefresh(shape);
};

/**
 * Reads the paragraph's horizontal alignment. Returns `null` when no
 * `algn` attribute is present (inherits from layout / master).
 */
export const getParagraphAlignment = (
  shape: SlideShapeData,
  paragraphIndex: number,
): ParagraphAlignment | null => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (pPr === null) return null;
  const v = getAttrValue(pPr, ATTR_ALGN_FN);
  return (v as ParagraphAlignment | null) ?? null;
};

/**
 * Reads the paragraph's nesting level (`lvl` attribute), or `0` when
 * absent — PowerPoint's default. Returns `null` for non-existent
 * paragraphs.
 */
export const getParagraphLevel = (shape: SlideShapeData, paragraphIndex: number): number => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (pPr === null) return 0;
  const v = getAttrValue(pPr, ATTR_LVL);
  if (v === null) return 0;
  const n = Number.parseInt(v, 10);
  return Number.isFinite(n) ? n : 0;
};

/**
 * Sets the spacing before and/or after a paragraph, in points (where
 * a "point" is 1/72 inch). PowerPoint stores these as hundredths of a
 * point inside `<a:pPr><a:spcBef>/<a:spcAft><a:spcPts val="…"/>` —
 * the helper converts.
 *
 *   setParagraphSpacing(shape, 0, { beforePts: 6, afterPts: 3 });
 *
 * Omitting a side keeps the existing value (or layout default).
 * Passing a side as `null` removes that spacing element.
 */
export const setParagraphSpacing = (
  shape: SlideShapeData,
  paragraphIndex: number,
  opts: { beforePts?: number | null; afterPts?: number | null },
): void => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = ensurePPr(paragraph);

  const writeSide = (localName: 'spcBef' | 'spcAft', value: number | null | undefined): void => {
    if (value === undefined) return;
    pPr.children = pPr.children.filter(
      (c) =>
        !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === localName),
    );
    if (value === null) return;
    if (!Number.isFinite(value) || value < 0) {
      throw new RangeError(`paragraph ${localName} must be a non-negative number, got ${value}`);
    }
    const spcEl = elem(qname('a', localName, NS.dml), {
      children: [
        elem(qname('a', 'spcPts', NS.dml), {
          attrs: [attr(qname('', 'val', ''), String(Math.round(value * 100)))],
        }),
      ],
    });
    pPr.children.push(spcEl);
  };

  writeSide('spcBef', opts.beforePts);
  writeSide('spcAft', opts.afterPts);
  commitAndRefresh(shape);
};

/**
 * Reads back paragraph spacing in points. Returns `{ beforePts,
 * afterPts }`; each side is `null` when no `<a:spcBef>` / `<a:spcAft>`
 * is present or when the inner element isn't `<a:spcPts>` (percentage
 * spacing is reported as `null` for now).
 */
export const getParagraphSpacing = (
  shape: SlideShapeData,
  paragraphIndex: number,
): { readonly beforePts: number | null; readonly afterPts: number | null } => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (!pPr) return { beforePts: null, afterPts: null };
  const readSide = (localName: 'spcBef' | 'spcAft'): number | null => {
    const side = firstChildElement(pPr, qname('a', localName, NS.dml));
    if (!side) return null;
    const spcPts = firstChildElement(side, qname('a', 'spcPts', NS.dml));
    if (!spcPts) return null;
    const v = getAttrValue(spcPts, qname('', 'val', ''));
    if (v === null) return null;
    const n = Number.parseInt(v, 10);
    return Number.isFinite(n) ? n / 100 : null;
  };
  return { beforePts: readSide('spcBef'), afterPts: readSide('spcAft') };
};

/**
 * Reads the paragraph's left / right / first-line indents from
 * `<a:pPr marL="…" marR="…" indent="…"/>`. Each is in EMU (matching
 * PowerPoint's internal storage); positive means a positive indent,
 * negative `indent` is a hanging indent (typical for bullets).
 *
 * Returns `null` for sides the paragraph doesn't set (those inherit
 * from the layout / master).
 */
export const getParagraphIndent = (
  shape: SlideShapeData,
  paragraphIndex: number,
): { leftEmu: number | null; rightEmu: number | null; firstLineEmu: number | null } => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (!pPr) return { leftEmu: null, rightEmu: null, firstLineEmu: null };
  const read = (name: string): number | null => {
    const raw = getAttrValue(pPr, qname('', name, ''));
    if (raw === null) return null;
    const n = Number.parseInt(raw, 10);
    return Number.isFinite(n) ? n : null;
  };
  return {
    leftEmu: read('marL'),
    rightEmu: read('marR'),
    firstLineEmu: read('indent'),
  };
};

/**
 * Reads the paragraph's `<a:lnSpc>` line spacing. PowerPoint stores
 * line spacing two ways:
 *
 *   - Multiple of the natural line height — `<a:spcPct val="150000"/>`
 *     (= 1.5×). Returns `{ kind: 'pct', value }` with value as the unit
 *     fraction (1.5).
 *   - Fixed points — `<a:spcPts val="2400"/>` (= 24pt). Returns
 *     `{ kind: 'pts', value }` with value in points.
 *
 * Returns `null` when no `<a:lnSpc>` is present (the paragraph
 * inherits line spacing from the layout / master).
 */
export const getParagraphLineSpacing = (
  shape: SlideShapeData,
  paragraphIndex: number,
):
  | { readonly kind: 'pct'; readonly value: number }
  | { readonly kind: 'pts'; readonly value: number }
  | null => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (!pPr) return null;
  const lnSpc = firstChildElement(pPr, qname('a', 'lnSpc', NS.dml));
  if (!lnSpc) return null;
  const pct = firstChildElement(lnSpc, qname('a', 'spcPct', NS.dml));
  if (pct) {
    const v = getAttrValue(pct, qname('', 'val', ''));
    if (v !== null) {
      let n = Number.parseFloat(v);
      if (Number.isFinite(n)) {
        if (Math.abs(n) > 1) n = n / 100000;
        return { kind: 'pct', value: n };
      }
    }
  }
  const pts = firstChildElement(lnSpc, qname('a', 'spcPts', NS.dml));
  if (pts) {
    const v = getAttrValue(pts, qname('', 'val', ''));
    if (v !== null) {
      const n = Number.parseInt(v, 10);
      if (Number.isFinite(n)) return { kind: 'pts', value: n / 100 };
    }
  }
  return null;
};

/**
 * Reads back the bullet style on a single paragraph, or `null` when
 * no `<a:buChar>` / `<a:buAutoNum>` / `<a:buNone>` is present (the
 * paragraph inherits its bullet from the layout / master).
 */
export const getParagraphBullet = (
  shape: SlideShapeData,
  paragraphIndex: number,
): BulletStyle | null => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (pPr === null) return null;
  for (const c of pPr.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (c.name.localName === 'buNone') return 'none';
    if (c.name.localName === 'buChar') {
      const char = getAttrValue(c, qname('', 'char', ''));
      if (char === '•') return 'bullet';
      if (char !== null) return { char };
    }
    if (c.name.localName === 'buAutoNum') {
      const t = getAttrValue(c, qname('', 'type', ''));
      if (t === 'arabicPeriod') return 'number';
      if (t !== null) return { autoNum: t };
    }
  }
  return null;
};

/**
 * Returns `true` when the paragraph uses an image as its bullet
 * (`<a:pPr><a:buBlip r:embed="…"/>`). Renderers without image
 * support should fall back to a generic bullet glyph.
 *
 * The underlying rId / image bytes aren't surfaced here — resolving
 * that would need the rels of the layout / master the paragraph
 * inherits from, which can be cumbersome. Knowing that the bullet
 * *is* an image is usually enough for the UI to pick a fallback.
 */
export const isParagraphBulletPicture = (
  shape: SlideShapeData,
  paragraphIndex: number,
): boolean => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (!pPr) return false;
  return firstChildElement(pPr, qname('a', 'buBlip', NS.dml)) !== null;
};

/**
 * Reads the bullet's per-paragraph color, size, and font overrides —
 * `<a:buClr>` (theme-resolved hex), `<a:buSzPct>` / `<a:buSzPts>`
 * (size relative to run or fixed pt), and `<a:buFont typeface="…"/>`.
 *
 * Returns `{ color: null, sizePct: null, sizePts: null, font: null }`
 * when the paragraph doesn't override any of them (the bullet inherits
 * from the run / layout).
 */
export const getParagraphBulletStyle = (
  pres: PresentationData,
  shape: SlideShapeData,
  paragraphIndex: number,
): {
  color: string | null;
  sizePct: number | null;
  sizePts: number | null;
  font: string | null;
} => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (!pPr) return { color: null, sizePct: null, sizePts: null, font: null };
  const theme = getPresentationTheme(pres);
  let color: string | null = null;
  let sizePct: number | null = null;
  let sizePts: number | null = null;
  let font: string | null = null;
  const buClr = firstChildElement(pPr, qname('a', 'buClr', NS.dml));
  if (buClr) {
    for (const c of buClr.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      color = resolveDrawingColor(c, theme);
      break;
    }
  }
  const buSzPct = firstChildElement(pPr, qname('a', 'buSzPct', NS.dml));
  if (buSzPct) {
    const v = getAttrValue(buSzPct, qname('', 'val', ''));
    if (v !== null) {
      let n = Number.parseFloat(v);
      if (Number.isFinite(n)) {
        if (Math.abs(n) > 1) n = n / 100000;
        sizePct = n;
      }
    }
  }
  const buSzPts = firstChildElement(pPr, qname('a', 'buSzPts', NS.dml));
  if (buSzPts) {
    const v = getAttrValue(buSzPts, qname('', 'val', ''));
    if (v !== null) {
      const n = Number.parseInt(v, 10);
      if (Number.isFinite(n)) sizePts = n / 100;
    }
  }
  const buFont = firstChildElement(pPr, qname('a', 'buFont', NS.dml));
  if (buFont) {
    const t = getAttrValue(buFont, qname('', 'typeface', ''));
    if (t !== null) font = t;
  }
  return { color, sizePct, sizePts, font };
};

/**
 * Sets the bullet style on a single paragraph. Same `BulletStyle` shape
 * as `setShapeBullets` — pass `'bullet'` / `'number'` / `'none'` or an
 * object like `{ char: '◆' }` / `{ autoNum: 'romanLcPeriod' }`.
 */
export const setParagraphBullet = (
  shape: SlideShapeData,
  paragraphIndex: number,
  style: BulletStyle,
): void => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  applyBulletToParagraph(paragraph, style);
  commitAndRefresh(shape);
};

/**
 * Sets the text of a single run. Existing rPr (font, size, color, ...)
 * is preserved — only the visible characters change.
 */
export const setShapeRunText = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
  text: string,
): void => {
  const run = requireRun(shape, paragraphIndex, runIndex);
  writeRunText(run, text);
  commitAndRefresh(shape);
};

// -- Color transforms (ECMA-376 §20.1.2.3.x) --------------------------------
//
// DrawingML color elements (`<a:srgbClr>`, `<a:schemeClr>`, `<a:sysClr>`,
// `<a:prstClr>`) may carry one or more transform children — `lumMod`,
// `lumOff`, `shade`, `tint`, `satMod`, `hueMod`, `alpha`, `gray`, `inv`,
// `comp`, etc. — that adjust the base color before it's painted. Real
// templates use them heavily for "tinted accent" backgrounds and "shaded
// hover" states, so any visual-fidelity story has to apply them.
//
// Percentages in the spec use the `ST_Percentage` style — `100000`
// represents 100% — though some third-party tools emit bare floats; we
// accept both forms.

type ColorTransformOp =
  | {
      readonly kind:
        | 'lumMod'
        | 'lumOff'
        | 'shade'
        | 'tint'
        | 'satMod'
        | 'satOff'
        | 'hueMod'
        | 'hueOff'
        | 'alpha'
        | 'alphaMod'
        | 'alphaOff';
      readonly val: number;
    }
  | { readonly kind: 'gray' | 'inv' | 'comp' };

const COLOR_TRANSFORM_LOCALS: ReadonlySet<string> = new Set([
  'lumMod',
  'lumOff',
  'shade',
  'tint',
  'satMod',
  'satOff',
  'hueMod',
  'hueOff',
  'alpha',
  'alphaMod',
  'alphaOff',
  'gray',
  'inv',
  'comp',
]);

const parseColorTransforms = (colorEl: XmlElement): readonly ColorTransformOp[] => {
  const out: ColorTransformOp[] = [];
  for (const child of colorEl.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.dml) continue;
    const local = child.name.localName;
    if (!COLOR_TRANSFORM_LOCALS.has(local)) continue;
    if (local === 'gray' || local === 'inv' || local === 'comp') {
      out.push({ kind: local });
      continue;
    }
    const raw = getAttrValue(child, qname('', 'val', ''));
    if (raw === null) continue;
    let n = Number.parseFloat(raw);
    if (!Number.isFinite(n)) continue;
    // PowerPoint emits ST_Percentage (`100000` = 100%); tolerate the
    // bare-float form some third-party tools emit.
    if (Math.abs(n) > 1) n = n / 100000;
    out.push({ kind: local as Exclude<ColorTransformOp['kind'], 'gray' | 'inv' | 'comp'>, val: n });
  }
  return out;
};

const hexToRgb01 = (hex: string): [number, number, number] => {
  const h = hex.startsWith('#') ? hex.slice(1) : hex;
  return [
    Number.parseInt(h.slice(0, 2), 16) / 255,
    Number.parseInt(h.slice(2, 4), 16) / 255,
    Number.parseInt(h.slice(4, 6), 16) / 255,
  ];
};

const rgb01ToHex = (r: number, g: number, b: number): string => {
  const clamp = (v: number): number => Math.max(0, Math.min(255, Math.round(v * 255)));
  const part = (n: number): string => n.toString(16).padStart(2, '0').toUpperCase();
  return `#${part(clamp(r))}${part(clamp(g))}${part(clamp(b))}`;
};

const rgbToHsl = (r: number, g: number, b: number): [number, number, number] => {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, s, l];
};

const hueToRgb = (p: number, q: number, t: number): number => {
  if (t < 0) t += 1;
  if (t > 1) t -= 1;
  if (t < 1 / 6) return p + (q - p) * 6 * t;
  if (t < 1 / 2) return q;
  if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
  return p;
};

const hslToRgb = (h: number, s: number, l: number): [number, number, number] => {
  if (s === 0) return [l, l, l];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return [hueToRgb(p, q, h + 1 / 3), hueToRgb(p, q, h), hueToRgb(p, q, h - 1 / 3)];
};

const applyColorTransforms = (hex: string, transforms: readonly ColorTransformOp[]): string => {
  if (transforms.length === 0) return hex;
  let [r, g, b] = hexToRgb01(hex);
  for (const t of transforms) {
    switch (t.kind) {
      case 'inv':
        r = 1 - r;
        g = 1 - g;
        b = 1 - b;
        break;
      case 'gray': {
        const y = 0.3 * r + 0.59 * g + 0.11 * b;
        r = g = b = y;
        break;
      }
      case 'comp': {
        const [h, s, l] = rgbToHsl(r, g, b);
        [r, g, b] = hslToRgb((h + 0.5) % 1, s, l);
        break;
      }
      case 'shade':
        // Mix toward black: out = base * val
        r *= t.val;
        g *= t.val;
        b *= t.val;
        break;
      case 'tint':
        // Mix toward white: out = base * val + (1 - val)
        r = r * t.val + (1 - t.val);
        g = g * t.val + (1 - t.val);
        b = b * t.val + (1 - t.val);
        break;
      case 'lumMod':
      case 'lumOff': {
        const [h, s, l] = rgbToHsl(r, g, b);
        const newL = Math.max(0, Math.min(1, t.kind === 'lumMod' ? l * t.val : l + t.val));
        [r, g, b] = hslToRgb(h, s, newL);
        break;
      }
      case 'satMod':
      case 'satOff': {
        const [h, s, l] = rgbToHsl(r, g, b);
        const newS = Math.max(0, Math.min(1, t.kind === 'satMod' ? s * t.val : s + t.val));
        [r, g, b] = hslToRgb(h, newS, l);
        break;
      }
      case 'hueMod':
      case 'hueOff': {
        const [h, s, l] = rgbToHsl(r, g, b);
        const newH = (((t.kind === 'hueMod' ? h * t.val : h + t.val / 360) % 1) + 1) % 1;
        [r, g, b] = hslToRgb(newH, s, l);
        break;
      }
      // alpha / alphaMod / alphaOff intentionally don't touch RGB — they
      // surface as `fill-opacity`, not as a tinted color.
    }
  }
  return rgb01ToHex(r, g, b);
};

const SCHEME_TOKEN_TO_THEME_KEY: Record<string, keyof Omit<PresentationTheme, 'name'>> = {
  tx1: 'dark1',
  dk1: 'dark1',
  bg1: 'light1',
  lt1: 'light1',
  tx2: 'dark2',
  dk2: 'dark2',
  bg2: 'light2',
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

const resolveSchemeToken = (token: string, theme: PresentationTheme | null): string | null => {
  if (!theme) return null;
  const key = SCHEME_TOKEN_TO_THEME_KEY[token];
  if (!key) return null;
  const hex = theme[key];
  if (typeof hex !== 'string') return null;
  const normalized = hex.startsWith('#') ? hex : `#${hex}`;
  return /^#[0-9A-Fa-f]{6}$/.test(normalized) ? normalized.toUpperCase() : null;
};

/**
 * Resolves a DrawingML color element (`<a:srgbClr>` / `<a:schemeClr>` /
 * `<a:sysClr>` / `<a:prstClr>`) with all its `<a:lumMod>` / `<a:tint>` /
 * `<a:shade>` / `<a:satMod>` etc. transform children applied. Returns
 * `null` when the color is a scheme token and no theme is supplied to
 * resolve it.
 *
 * Exposed because both run-format and fill-format code paths need to
 * apply the same transform pipeline; keeping a single implementation
 * means future spec-coverage additions only have to land in one place.
 */
export const resolveDrawingColor = (
  colorEl: XmlElement,
  theme: PresentationTheme | null,
): string | null => {
  if (colorEl.name.namespaceURI !== NS.dml) return null;
  const local = colorEl.name.localName;
  let baseHex: string | null = null;
  if (local === 'srgbClr') {
    const v = getAttrValue(colorEl, qname('', 'val', ''));
    if (v) baseHex = `#${v.toUpperCase()}`;
  } else if (local === 'schemeClr') {
    const v = getAttrValue(colorEl, qname('', 'val', ''));
    if (v) baseHex = resolveSchemeToken(v, theme);
  } else if (local === 'sysClr') {
    const last = getAttrValue(colorEl, qname('', 'lastClr', ''));
    if (last) baseHex = `#${last.toUpperCase()}`;
  } else if (local === 'prstClr') {
    // Preset colors aren't worth a full lookup table in this pass —
    // black / white cover most cases anyone reaches for in PresentationML.
    const v = getAttrValue(colorEl, qname('', 'val', ''));
    if (v === 'black') baseHex = '#000000';
    else if (v === 'white') baseHex = '#FFFFFF';
  }
  if (!baseHex) return null;
  return applyColorTransforms(baseHex, parseColorTransforms(colorEl));
};

// Reads any element shaped like `CT_TextCharacterProperties` (the schema
// shared by `<a:rPr>`, `<a:defRPr>`, and `<a:endParaRPr>`) into a partial
// TextFormat. Used by both the literal-only `getShapeRunFormat` and the
// inheritance-aware `getShapeRunFormatEffective`.
//
// When `ctx.theme` is provided, scheme tokens are resolved to concrete
// `#RRGGBB` and color transforms (`<a:lumMod>` etc.) are applied. Without
// a theme, transforms are not applied and theme tokens are passed through
// verbatim — this preserves the legacy `getShapeRunFormat` behavior.
const parseRPrLikeElement = (
  rPr: XmlElement,
  ctx?: { readonly theme: PresentationTheme | null },
): Partial<TextFormat> => {
  const out: Partial<TextFormat> = {};
  const sz = getAttrValue(rPr, qname('', 'sz', ''));
  if (sz !== null) {
    const n = Number.parseInt(sz, 10);
    if (Number.isFinite(n)) out.size = n / 100;
  }
  const b = getAttrValue(rPr, qname('', 'b', ''));
  if (b !== null) out.bold = b !== '0';
  const i = getAttrValue(rPr, qname('', 'i', ''));
  if (i !== null) out.italic = i !== '0';
  const u = getAttrValue(rPr, qname('', 'u', ''));
  if (u !== null) {
    if (u === 'none') out.underline = false;
    else if (u === 'sng') out.underline = true;
    else out.underline = u;
  }
  const strike = getAttrValue(rPr, qname('', 'strike', ''));
  if (strike !== null) {
    if (strike === 'noStrike') out.strike = false;
    else if (strike === 'sngStrike') out.strike = true;
    else out.strike = strike;
  }
  const spc = getAttrValue(rPr, qname('', 'spc', ''));
  if (spc !== null) {
    const n = Number.parseInt(spc, 10);
    if (Number.isFinite(n)) out.spc = n;
  }
  const kern = getAttrValue(rPr, qname('', 'kern', ''));
  if (kern !== null) {
    const n = Number.parseInt(kern, 10);
    if (Number.isFinite(n)) out.kern = n;
  }
  const baselineAttr = getAttrValue(rPr, qname('', 'baseline', ''));
  if (baselineAttr !== null) {
    // ST_Percentage: 100000 = 100%; tolerate bare floats.
    let n = Number.parseFloat(baselineAttr);
    if (Number.isFinite(n)) {
      if (Math.abs(n) > 1) n = n / 100000;
      out.baseline = n;
    }
  }
  const cap = getAttrValue(rPr, qname('', 'cap', ''));
  if (cap === 'none' || cap === 'small' || cap === 'all') {
    out.cap = cap;
  }
  // <a:highlight><a:srgbClr val="…"/></a:highlight>
  const highlight = firstChildElement(rPr, qname('a', 'highlight', NS.dml));
  if (highlight !== null) {
    let hlChild: XmlElement | null = null;
    for (const c of highlight.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      hlChild = c;
      break;
    }
    if (hlChild) {
      if (ctx) {
        const hex = resolveDrawingColor(hlChild, ctx.theme);
        if (hex !== null) out.highlight = hex;
      } else if (hlChild.name.localName === 'srgbClr') {
        const v = getAttrValue(hlChild, qname('', 'val', ''));
        if (v !== null) out.highlight = `#${v.toUpperCase()}`;
      } else if (hlChild.name.localName === 'schemeClr') {
        const v = getAttrValue(hlChild, qname('', 'val', ''));
        if (v !== null) out.highlight = v;
      }
    }
  }
  const solidFill = firstChildElement(rPr, qname('a', 'solidFill', NS.dml));
  if (solidFill !== null) {
    // Find the inner color element (srgbClr / schemeClr / sysClr / prstClr).
    // CT_SolidColorFillProperties holds exactly one EG_ColorChoice child.
    let colorChild: XmlElement | null = null;
    for (const c of solidFill.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      colorChild = c;
      break;
    }
    if (colorChild) {
      if (ctx) {
        // Apply transforms + resolve scheme tokens to hex.
        const hex = resolveDrawingColor(colorChild, ctx.theme);
        if (hex !== null) out.color = hex;
        else if (colorChild.name.localName === 'schemeClr') {
          // Theme not provided / token not in scheme — surface the raw token.
          const v = getAttrValue(colorChild, qname('', 'val', ''));
          if (v !== null) out.color = v;
        }
      } else {
        // Legacy `getShapeRunFormat` path: no transforms, scheme tokens
        // emitted as bare strings to match prior public behavior.
        if (colorChild.name.localName === 'srgbClr') {
          const v = getAttrValue(colorChild, qname('', 'val', ''));
          if (v !== null) out.color = `#${v.toUpperCase()}`;
        } else if (colorChild.name.localName === 'schemeClr') {
          const v = getAttrValue(colorChild, qname('', 'val', ''));
          if (v !== null) out.color = v;
        }
      }
    }
  }
  const latin = firstChildElement(rPr, qname('a', 'latin', NS.dml));
  if (latin !== null) {
    const t = getAttrValue(latin, qname('', 'typeface', ''));
    if (t !== null) out.font = t;
  }
  return out;
};

/**
 * Reads back the format of a single run. Returns `null` when the run
 * has no `<a:rPr>` (it inherits its format from the paragraph /
 * layout / master). Boolean attributes that are explicitly `"0"`
 * decode to `false`.
 *
 * Use `getShapeRunFormatEffective` if you want the resolved format
 * after walking the placeholder / lstStyle / master inheritance chain.
 */
export const getShapeRunFormat = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
): TextFormat | null => {
  const run = requireRun(shape, paragraphIndex, runIndex);
  const rPr = firstChildElement(run, NAME_A_RPR);
  if (rPr === null) return null;
  return parseRPrLikeElement(rPr) as TextFormat;
};

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
): { element: XmlElement } | undefined => {
  let match = phIdx !== null ? shapes.find((s) => s.placeholderIdx === phIdx) : undefined;
  if (!match && phType !== null) {
    match = shapes.find((s) => s.placeholderType === phType);
  }
  return match;
};

const extractPlaceholderLstStyle = (placeholderEl: XmlElement): XmlElement | null => {
  const txBody = firstChildElement(placeholderEl, NAME_P_TX_BODY_PML);
  if (!txBody) return null;
  return firstChildElement(txBody, NAME_A_LST_STYLE);
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

  const slide = shape[SHAPE_SLIDE];
  const layout = getSlideLayout(slide);

  if (layout) {
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
  const slide = shape[SHAPE_SLIDE];
  const layout = getSlideLayout(slide);

  if (layout) {
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
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY_FN);
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
export const setShapeHyperlink = (shape: SlideShapeData, url: string | null): void => {
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
    applyHyperlinkToAllRuns(txBody, rId);
  }
  commitAndRefresh(shape);
};

// ---------------------------------------------------------------------------
// Shape mutation — removal.

/**
 * Copies a shape into `targetSlide`. The source XML is cloned and
 * appended to the target's `<p:spTree>`. Image rels on the source
 * shape are followed: the linked media part is referenced from the
 * target slide via a freshly allocated rId (no media bytes are
 * copied — both slides share the underlying part).
 *
 * v1 requires source and target to live in the same package
 * (`sourceShape`'s slide and `targetSlide` must share the same
 * `OpcPackage`). Cross-package copy is `importSlide` territory.
 *
 * Returns the new `SlideShapeData` on `targetSlide`.
 */
export const copyShape = (targetSlide: SlideData, sourceShape: SlideShapeData): SlideShapeData => {
  const sourceSlide = sourceShape[SHAPE_SLIDE];
  if (sourceSlide[INTERNAL_PACKAGE] !== targetSlide[INTERNAL_PACKAGE]) {
    throw new Error(
      'copyShape: source and target must be in the same package. Use importSlide for cross-deck copies.',
    );
  }
  const pkg = targetSlide[INTERNAL_PACKAGE];
  const sourceEl = sourceShape[SHAPE_ELEMENT];

  // Deep-clone the XML by serializing + re-parsing one element.
  // We wrap in a temporary parent so we can extract the cloned element
  // back out without ambient namespaces leaking from the slide root.
  const cloned = cloneXmlElement(sourceEl);

  // Allocate a fresh shape id on the target slide and overwrite the
  // cNvPr/cNvPr id attribute.
  const newId = nextShapeId(targetSlide);
  rewriteCNvPrId(cloned, newId);

  // Walk the cloned element for r:embed / r:link references. For each
  // referenced rId in the source slide's rels, copy the rel onto the
  // target slide's rels (allocating a fresh rId) and update the cloned
  // attribute. This covers picture blips + media references.
  const sourceRels = pkg.getRels(sourceSlide[SLIDE_PART_NAME]);
  if (sourceRels) {
    const targetRels = pkg.getRels(targetSlide[SLIDE_PART_NAME]) ?? emptyRels();
    const usedIds = new Set(targetRels.items.map((r) => r.id));
    rewriteRIdReferences(cloned, (oldRId) => {
      const sourceRel = sourceRels.items.find((r) => r.id === oldRId);
      if (!sourceRel) return oldRId;
      // Look for an existing rel on target with the same type+target;
      // reuse if found to avoid duplicates.
      const existing = targetRels.items.find(
        (r) =>
          r.type === sourceRel.type &&
          r.target === sourceRel.target &&
          r.targetMode === sourceRel.targetMode,
      );
      if (existing) return existing.id;
      const newRId = nextRelId([...usedIds]);
      usedIds.add(newRId);
      targetRels.items.push({ ...sourceRel, id: newRId });
      return newRId;
    });
    pkg.setRels(targetSlide[SLIDE_PART_NAME], targetRels);
  }

  return appendAndReturnNewShape(targetSlide, cloned);
};

/** Recursively clone an XML element (no parent, deep). */
const cloneXmlElement = (el: XmlElement): XmlElement => ({
  kind: 'element',
  name: el.name,
  attrs: el.attrs.map((a) => ({ name: a.name, value: a.value })),
  prefixDecls: new Map(el.prefixDecls),
  children: el.children.map((c) => {
    if (c.kind === 'element') return cloneXmlElement(c);
    return { ...c };
  }),
});

const rewriteCNvPrId = (root: XmlElement, newId: number): void => {
  const walk = (el: XmlElement): boolean => {
    if (
      el.name.namespaceURI === NS.pml &&
      el.name.localName === 'cNvPr' &&
      el.attrs.some((a) => a.name.namespaceURI === '' && a.name.localName === 'id')
    ) {
      el.attrs = el.attrs.map((a) =>
        a.name.namespaceURI === '' && a.name.localName === 'id'
          ? { name: a.name, value: String(newId) }
          : a,
      );
      return true;
    }
    for (const c of el.children) {
      if (c.kind === 'element' && walk(c)) return true;
    }
    return false;
  };
  walk(root);
};

const rewriteRIdReferences = (root: XmlElement, map: (oldRId: string) => string): void => {
  const walk = (el: XmlElement): void => {
    el.attrs = el.attrs.map((a) => {
      if (
        a.name.namespaceURI === NS.officeDocRels &&
        (a.name.localName === 'id' || a.name.localName === 'embed' || a.name.localName === 'link')
      ) {
        return { name: a.name, value: map(a.value) };
      }
      return a;
    });
    for (const c of el.children) {
      if (c.kind === 'element') walk(c);
    }
  };
  walk(root);
};

// ---------------------------------------------------------------------------
// Z-order — move shapes forward / backward inside the slide's spTree.
//
// OOXML shape z-order is just the document order of children of
// `<p:spTree>`: the first child renders behind, the last in front.
// PowerPoint's "Bring to Front" / "Send to Back" affordances translate
// directly to reordering those children.
//
// Each function targets only "real" shape children — `<p:sp>`, `<p:pic>`,
// `<p:cxnSp>`, `<p:graphicFrame>`, `<p:grpSp>`. The required
// `<p:nvGrpSpPr>` / `<p:grpSpPr>` preface stays at the top.

const SHAPE_CHILD_LOCALS = new Set(['sp', 'pic', 'cxnSp', 'graphicFrame', 'grpSp']);

const isShapeChild = (node: {
  kind: string;
  name?: { namespaceURI: string; localName: string };
}): boolean =>
  node.kind === 'element' &&
  node.name?.namespaceURI === NS.pml &&
  SHAPE_CHILD_LOCALS.has(node.name.localName);

/** Move `shape` to the end of its spTree (render in front of all others). */
export const bringShapeToFront = (shape: SlideShapeData): void => {
  const slide = shape[SHAPE_SLIDE];
  const spTree = requireSpTree(slide);
  const target = shape[SHAPE_ELEMENT];
  const idx = spTree.children.indexOf(target);
  if (idx < 0) return;
  if (idx === spTree.children.length - 1) return; // already at front
  spTree.children.splice(idx, 1);
  spTree.children.push(target);
  commitSlideData(slide);
  rebuildShapesFromDocument(slide);
};

/**
 * Move `shape` behind every other shape on the slide. The
 * `<p:nvGrpSpPr>` / `<p:grpSpPr>` preface — required by the schema —
 * stays at the top.
 */
export const sendShapeToBack = (shape: SlideShapeData): void => {
  const slide = shape[SHAPE_SLIDE];
  const spTree = requireSpTree(slide);
  const target = shape[SHAPE_ELEMENT];
  const idx = spTree.children.indexOf(target);
  if (idx < 0) return;

  // First "shape child" position — after nvGrpSpPr / grpSpPr.
  let firstShapeAt = spTree.children.length;
  for (let i = 0; i < spTree.children.length; i++) {
    const c = spTree.children[i];
    if (c && isShapeChild(c)) {
      firstShapeAt = i;
      break;
    }
  }
  if (idx <= firstShapeAt) return;
  spTree.children.splice(idx, 1);
  spTree.children.splice(firstShapeAt, 0, target);
  commitSlideData(slide);
  rebuildShapesFromDocument(slide);
};

/** Swap `shape` with the next shape sibling (move one step forward). */
export const bringShapeForward = (shape: SlideShapeData): void => {
  const slide = shape[SHAPE_SLIDE];
  const spTree = requireSpTree(slide);
  const target = shape[SHAPE_ELEMENT];
  const idx = spTree.children.indexOf(target);
  if (idx < 0) return;
  // Find next shape sibling.
  for (let i = idx + 1; i < spTree.children.length; i++) {
    const c = spTree.children[i];
    if (c && isShapeChild(c)) {
      const next = c;
      spTree.children[idx] = next;
      spTree.children[i] = target;
      commitSlideData(slide);
      rebuildShapesFromDocument(slide);
      return;
    }
  }
};

/**
 * Returns the shape's z-index among the slide's "real" shape children
 * (`<p:sp>` / `<p:pic>` / `<p:cxnSp>` / `<p:graphicFrame>` / `<p:grpSp>`),
 * skipping the required `<p:nvGrpSpPr>` / `<p:grpSpPr>` preface.
 * Higher numbers render in front.
 */
export const getShapeZIndex = (shape: SlideShapeData): number => {
  const slide = shape[SHAPE_SLIDE];
  const spTree = requireSpTree(slide);
  let i = 0;
  for (const c of spTree.children) {
    if (!isShapeChild(c)) continue;
    if (c === shape[SHAPE_ELEMENT]) return i;
    i++;
  }
  return -1;
};

/**
 * Moves the shape to a specific z-index among the slide's "real"
 * shape children. Index is clamped to the available range. Higher
 * numbers render in front. The required preface elements stay at the
 * top of `<p:spTree>`.
 */
export const setShapeZIndex = (shape: SlideShapeData, toIndex: number): void => {
  const slide = shape[SHAPE_SLIDE];
  const spTree = requireSpTree(slide);
  const target = shape[SHAPE_ELEMENT];
  const allShapeChildren = spTree.children.filter((c): c is XmlElement => isShapeChild(c));
  const clamped = Math.max(0, Math.min(toIndex, allShapeChildren.length - 1));

  // Remove the target from the tree, then re-insert at the position
  // corresponding to z-index `clamped` among the remaining shapes.
  spTree.children = spTree.children.filter((c) => c !== target);
  const remainingShapes = spTree.children.filter((c): c is XmlElement => isShapeChild(c));
  if (clamped >= remainingShapes.length) {
    spTree.children.push(target);
  } else {
    const anchor = remainingShapes[clamped]!;
    const anchorIdx = spTree.children.indexOf(anchor);
    spTree.children.splice(anchorIdx, 0, target);
  }
  commitSlideData(slide);
  rebuildShapesFromDocument(slide);
};

/** Swap `shape` with the previous shape sibling (move one step backward). */
export const sendShapeBackward = (shape: SlideShapeData): void => {
  const slide = shape[SHAPE_SLIDE];
  const spTree = requireSpTree(slide);
  const target = shape[SHAPE_ELEMENT];
  const idx = spTree.children.indexOf(target);
  if (idx < 0) return;
  for (let i = idx - 1; i >= 0; i--) {
    const c = spTree.children[i];
    if (c && isShapeChild(c)) {
      const prev = c;
      spTree.children[idx] = prev;
      spTree.children[i] = target;
      commitSlideData(slide);
      rebuildShapesFromDocument(slide);
      return;
    }
  }
};

/**
 * Removes the shape from its slide's shape tree. Subsequent property
 * reads on this handle reflect the stale snapshot — discard it after.
 *
 * Removing a picture does NOT delete the underlying media part — it
 * may be referenced from other slides.
 */
/**
 * Removes every shape (sp / pic / cxnSp / graphicFrame / grpSp) from
 * the slide's `<p:spTree>`. The required `<p:nvGrpSpPr>` and
 * `<p:grpSpPr>` preface stays in place, so the slide is still valid
 * and re-applies its layout's placeholders on the next open.
 *
 * Useful for "start this slide over but keep its layout binding."
 */
export const clearSlideShapes = (slide: SlideData): void => {
  const spTree = requireSpTree(slide);
  spTree.children = spTree.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.pml &&
        SHAPE_CHILD_LOCALS.has(c.name.localName)
      ),
  );
  commitSlideData(slide);
  rebuildShapesFromDocument(slide);
};

export const removeShape = (shape: SlideShapeData): void => {
  const slide = shape[SHAPE_SLIDE];
  const doc = slide[SLIDE_DOCUMENT];
  const cSld = firstChildElement(doc.root, qname('p', 'cSld', NS.pml));
  if (!cSld) return;
  const spTree = firstChildElement(cSld, qname('p', 'spTree', NS.pml));
  if (!spTree) return;
  const idx = spTree.children.indexOf(shape[SHAPE_ELEMENT]);
  if (idx < 0) return;
  spTree.children.splice(idx, 1);
  commitSlideData(slide);
  rebuildShapesFromDocument(slide);
};

// ---------------------------------------------------------------------------
// Slide-level shape authoring.
//
// Each `addXxx` builds an XML element via an internal builder, appends
// it to the slide's `<p:spTree>`, commits, rebuilds the typed view, and
// returns the new SlideShapeData.

const requireSpTree = (slide: SlideData): XmlElement => {
  const cSld = firstChildElement(slide[SLIDE_DOCUMENT].root, NAME_CSLD);
  if (!cSld) throw new Error('slide has no <p:cSld>');
  const spTree = firstChildElement(cSld, NAME_SP_TREE);
  if (!spTree) throw new Error('slide has no <p:spTree>');
  return spTree;
};

const nextShapeId = (slide: SlideData): number => {
  let maxId = 0;
  for (const s of slide[SLIDE_PART].shapes) {
    if (s.id > maxId) maxId = s.id;
  }
  return Math.max(maxId, 1) + 1;
};

const appendAndReturnNewShape = (slide: SlideData, child: XmlElement): SlideShapeData => {
  const spTree = requireSpTree(slide);
  spTree.children.push(child);
  commitSlideData(slide);
  const previousLength = slide[SLIDE_SHAPES].length;
  rebuildShapesFromDocument(slide);
  const created = slide[SLIDE_SHAPES][previousLength];
  if (!created) throw new Error('appendShape: post-condition failed');
  return created;
};

/**
 * Adds a free-form text box to the slide. Returns the new shape.
 *
 * The box is a plain rectangle with no fill or outline carrying one
 * paragraph with one run. The shape id is allocated as one more than
 * the current max id.
 */
export const addSlideTextBox = (
  slide: SlideData,
  opts: { x: Emu; y: Emu; w: Emu; h: Emu; text: string; name?: string },
): SlideShapeData => {
  const sp = buildTextBox({
    id: nextShapeId(slide),
    ...(opts.name !== undefined ? { name: opts.name } : {}),
    x: opts.x,
    y: opts.y,
    w: opts.w,
    h: opts.h,
    text: opts.text,
  });
  return appendAndReturnNewShape(slide, sp);
};

/**
 * Adds a preset shape (rectangle, ellipse, arrow, ...) to the slide.
 * Optional `text` seeds a single run.
 */
export const addSlideShape = (
  slide: SlideData,
  opts: {
    preset: PresetShape | string;
    x: Emu;
    y: Emu;
    w: Emu;
    h: Emu;
    text?: string;
    textAnchor?: 'l' | 'ctr' | 'r' | 't' | 'b';
    name?: string;
  },
): SlideShapeData => {
  const sp = buildShape({
    id: nextShapeId(slide),
    ...(opts.name !== undefined ? { name: opts.name } : {}),
    preset: opts.preset,
    x: opts.x,
    y: opts.y,
    w: opts.w,
    h: opts.h,
    ...(opts.text !== undefined ? { text: opts.text } : {}),
    ...(opts.textAnchor !== undefined ? { textAnchor: opts.textAnchor } : {}),
  });
  return appendAndReturnNewShape(slide, sp);
};

/** Adds a straight-line connector between two points. */
export const addSlideLine = (
  slide: SlideData,
  opts: {
    from: { x: Emu; y: Emu };
    to: { x: Emu; y: Emu };
    color?: string;
    widthEmu?: number;
    name?: string;
  },
): SlideShapeData => {
  const cxn = buildConnector({
    id: nextShapeId(slide),
    ...(opts.name !== undefined ? { name: opts.name } : {}),
    from: opts.from,
    to: opts.to,
    ...(opts.color !== undefined ? { color: opts.color } : {}),
    ...(opts.widthEmu !== undefined ? { widthEmu: opts.widthEmu } : {}),
  });
  return appendAndReturnNewShape(slide, cxn);
};

/**
 * Adds a table to the slide. Cells render as plain text with default
 * theme-aware styling; `firstRow` / `bandRow` flags drive PowerPoint's
 * banded-header look unless options say otherwise.
 */
export const addSlideTable = (
  slide: SlideData,
  opts: {
    x: Emu;
    y: Emu;
    w: Emu;
    h: Emu;
    rows: ReadonlyArray<ReadonlyArray<string>>;
    colWidths?: ReadonlyArray<Emu>;
    rowHeights?: ReadonlyArray<Emu>;
    firstRow?: boolean;
    bandRow?: boolean;
    name?: string;
  },
): SlideShapeData => {
  const frame = buildTable({
    id: nextShapeId(slide),
    ...(opts.name !== undefined ? { name: opts.name } : {}),
    x: opts.x,
    y: opts.y,
    w: opts.w,
    h: opts.h,
    rows: opts.rows,
    ...(opts.colWidths !== undefined ? { colWidths: opts.colWidths } : {}),
    ...(opts.rowHeights !== undefined ? { rowHeights: opts.rowHeights } : {}),
    ...(opts.firstRow !== undefined ? { firstRow: opts.firstRow } : {}),
    ...(opts.bandRow !== undefined ? { bandRow: opts.bandRow } : {}),
  });
  return appendAndReturnNewShape(slide, frame);
};

/**
 * Adds a picture to the slide from raw bytes. Returns the new shape.
 *
 * Allocates a `/ppt/media/imageN.<ext>` part, registers a Content_Types
 * Default if the extension isn't yet covered, allocates a slide→image
 * rel, and appends a `<p:pic>` element to the slide's `<p:spTree>`.
 *
 * Format is detected from magic bytes; pass `opts.format` to override.
 */
export const addSlideImage = (
  slide: SlideData,
  bytes: Uint8Array,
  opts: { x: Emu; y: Emu; w: Emu; h: Emu; format?: ImageFormat; name?: string },
): SlideShapeData => {
  const pkg = slide[INTERNAL_PACKAGE];
  const format = opts.format ?? detectImageFormat(bytes);
  if (format === null) {
    throw new Error(
      'addSlideImage: could not detect image format. Pass options.format explicitly.',
    );
  }
  const contentType = contentTypeForFormat(format);
  const extension = extensionForFormat(format);

  let nextN = 1;
  const mediaPattern = /^\/ppt\/media\/image(\d+)\./;
  for (const p of pkg.parts) {
    const m = p.name.match(mediaPattern);
    if (m?.[1] !== undefined) {
      const n = Number.parseInt(m[1], 10);
      if (Number.isFinite(n) && n >= nextN) nextN = n + 1;
    }
  }
  const newMediaName = partName(`/ppt/media/image${nextN}.${extension}`);

  const hasDefault = pkg.contentTypes.defaults.some((d) => d.extension.toLowerCase() === extension);
  if (!hasDefault) {
    pkg.contentTypes.defaults.push({ extension, contentType });
  }
  pkg.addPart(newMediaName, contentType, bytes);

  const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
  const newRId = nextRelId(rels.items.map((r) => r.id));
  rels.items.push({
    id: newRId,
    type: REL_TYPES.image,
    target: `../media/image${nextN}.${extension}`,
    targetMode: 'Internal',
  });
  pkg.setRels(slide[SLIDE_PART_NAME], rels);

  const pic = buildPicture({
    id: nextShapeId(slide),
    ...(opts.name !== undefined ? { name: opts.name } : {}),
    rEmbed: newRId,
    x: opts.x,
    y: opts.y,
    w: opts.w,
    h: opts.h,
  });
  return appendAndReturnNewShape(slide, pic);
};

// ---------------------------------------------------------------------------
// Slide-level background + transition.

const removeTransition = (slide: SlideData): void => {
  slide[SLIDE_DOCUMENT].root.children = slide[SLIDE_DOCUMENT].root.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.pml &&
        c.name.localName === 'transition'
      ),
  );
};

const insertAfterClrMapOvr = (slide: SlideData, t: XmlElement): void => {
  const children = slide[SLIDE_DOCUMENT].root.children;
  let insertAt = children.length;
  for (let i = 0; i < children.length; i++) {
    const c = children[i];
    if (c?.kind !== 'element' || c.name.namespaceURI !== NS.pml) continue;
    if (c.name.localName === 'clrMapOvr') {
      insertAt = i + 1;
    } else if (c.name.localName === 'cSld' && insertAt === children.length) {
      insertAt = i + 1;
    }
  }
  children.splice(insertAt, 0, t);
};

/**
 * Reads back the slide's current transition (or `null` if no
 * `<p:transition>` is present). The returned shape mirrors what
 * `setSlideTransition` accepts.
 */
export const getSlideTransition = (slide: SlideData): TransitionOptions | null => {
  const transition = slide[SLIDE_DOCUMENT].root.children.find(
    (c): c is XmlElement =>
      c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'transition',
  );
  if (!transition) return null;
  const speed = getAttrValue(transition, qname('', 'spd', '')) as 'slow' | 'med' | 'fast' | null;
  const advClick = getAttrValue(transition, qname('', 'advClick', ''));
  const advTm = getAttrValue(transition, qname('', 'advTm', ''));
  // First child element identifies the effect (`p:fade`, `p:wipe`, ...).
  let effect: string | null = null;
  let direction: string | null = null;
  let orientation: 'horz' | 'vert' | null = null;
  let thruBlack: boolean | undefined;
  for (const child of transition.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.pml) continue;
    effect = child.name.localName;
    direction = getAttrValue(child, qname('', 'dir', ''));
    const o = getAttrValue(child, qname('', 'orient', ''));
    if (o === 'horz' || o === 'vert') orientation = o;
    const tb = getAttrValue(child, qname('', 'thruBlk', ''));
    if (tb !== null) thruBlack = tb === '1';
    break;
  }
  if (effect === null) return null;
  return {
    effect,
    ...(speed !== null ? { speed } : {}),
    ...(direction !== null ? { direction } : {}),
    ...(orientation !== null ? { orientation } : {}),
    ...(thruBlack !== undefined ? { thruBlack } : {}),
    ...(advClick !== null ? { advanceOnClick: advClick !== '0' } : {}),
    ...(advTm !== null ? { advanceAfterMs: Number.parseInt(advTm, 10) } : {}),
  };
};

/** Sets the slide's transition effect. */
export const setSlideTransition = (slide: SlideData, options: TransitionOptions): void => {
  removeTransition(slide);
  insertAfterClrMapOvr(slide, buildTransition(options));
  commitSlideData(slide);
  refreshSlideData(slide);
};

/** Removes any existing transition on the slide. */
export const clearSlideTransition = (slide: SlideData): void => {
  removeTransition(slide);
  commitSlideData(slide);
  refreshSlideData(slide);
};

const setSlideBackgroundXml = (slide: SlideData, configure: (bgPr: XmlElement) => void): void => {
  const cSld = firstChildElement(slide[SLIDE_DOCUMENT].root, NAME_CSLD);
  if (!cSld) throw new Error('slide has no <p:cSld>');
  const bgName = qname('p', 'bg', NS.pml);
  const bgPrName = qname('p', 'bgPr', NS.pml);
  let bg = firstChildElement(cSld, bgName);
  if (bg === null) {
    bg = { kind: 'element', name: bgName, attrs: [], prefixDecls: new Map(), children: [] };
    cSld.children.unshift(bg);
  }
  bg.children = [];
  const bgPr: XmlElement = {
    kind: 'element',
    name: bgPrName,
    attrs: [],
    prefixDecls: new Map(),
    children: [],
  };
  bg.children.push(bgPr);
  configure(bgPr);
  commitSlideData(slide);
  refreshSlideData(slide);
};

/**
 * Reads back the slide's current background. Returns a discriminated
 * union mirroring `getShapeFill`'s shape, plus `inherit` when no
 * `<p:bg>` element is present (the slide picks up its background from
 * the layout / master).
 */
export type SlideBackground =
  | { readonly kind: 'solid'; readonly color: string }
  | { readonly kind: 'gradient' }
  | { readonly kind: 'pattern' }
  | { readonly kind: 'image' }
  | { readonly kind: 'inherit' };

/**
 * Reads the slide's color-map override (`<p:clrMapOvr><p:overrideClrMapping/>`).
 * The mapping remaps the eight stable ECMA-376 color tokens (`bg1`,
 * `tx1`, `bg2`, `tx2`, `accent1`–`accent6`, `hlink`, `folHlink`) to
 * different theme positions. Returns `null` when the slide uses the
 * master's color map unchanged (the overwhelming common case).
 */
export const getSlideColorMapOverride = (slide: SlideData): Record<string, string> | null => {
  const root = slide[SLIDE_DOCUMENT].root;
  let ovr: XmlElement | null = null;
  for (const c of root.children) {
    if (c.kind !== 'element') continue;
    if (c.name.namespaceURI === NS.pml && c.name.localName === 'clrMapOvr') {
      ovr = c;
      break;
    }
  }
  if (!ovr) return null;
  const mapping = firstChildElement(ovr, qname('a', 'overrideClrMapping', NS.dml));
  if (!mapping) return null;
  // overrideClrMapping carries 12 attributes — bg1..folHlink — each
  // pointing to the index the token is remapped to in the theme.
  const out: Record<string, string> = {};
  for (const a of mapping.attrs) {
    if (a.name.namespaceURI !== '') continue;
    out[a.name.localName] = a.value;
  }
  return Object.keys(out).length > 0 ? out : null;
};

export const getSlideBackground = (slide: SlideData): SlideBackground => {
  const cSld = firstChildElement(slide[SLIDE_DOCUMENT].root, NAME_CSLD);
  if (!cSld) return { kind: 'inherit' };
  const bg = firstChildElement(cSld, qname('p', 'bg', NS.pml));
  if (!bg) return { kind: 'inherit' };
  const bgPr = firstChildElement(bg, qname('p', 'bgPr', NS.pml));
  if (!bgPr) return { kind: 'inherit' };
  for (const c of bgPr.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    switch (c.name.localName) {
      case 'solidFill': {
        for (const inner of c.children) {
          if (inner.kind !== 'element' || inner.name.namespaceURI !== NS.dml) continue;
          if (inner.name.localName === 'srgbClr') {
            const val = getAttrValue(inner, qname('', 'val', ''));
            if (val !== null) return { kind: 'solid', color: `#${val.toUpperCase()}` };
          }
          if (inner.name.localName === 'schemeClr') {
            const val = getAttrValue(inner, qname('', 'val', ''));
            if (val !== null) return { kind: 'solid', color: `scheme:${val}` };
          }
        }
        return { kind: 'solid', color: '' };
      }
      case 'gradFill':
        return { kind: 'gradient' };
      case 'pattFill':
        return { kind: 'pattern' };
      case 'blipFill':
        return { kind: 'image' };
    }
  }
  return { kind: 'inherit' };
};

/**
 * A simplified, render-ready view of one of the layout's non-placeholder
 * shapes. Resolves the bounds, preset, fill, and stroke without going
 * through the slide-bound `SlideShapeData` symbols. Returned by
 * `getSlideLayoutBackgroundShapes` for renderers that want to paint the
 * layout's brand-template decoration (corner bars, divider lines, logos
 * as solid rects, etc.) behind the slide's own shapes.
 */
export interface SlideLayoutBackgroundShape {
  readonly kind: 'shape' | 'connector' | 'picture' | 'group' | 'graphicFrame';
  /** Bounds in EMU, or `null` when the shape inherits from its master. */
  readonly bounds: ShapeBounds | null;
  /** Preset geometry token for shapes (`'rect'`, `'roundRect'`, `'ellipse'`, …). */
  readonly preset: string | null;
  /** Fill color resolved to `#RRGGBB` (transforms + theme applied), or `null`. */
  readonly fillHex: string | null;
  /** Stroke color resolved to `#RRGGBB`, or `null`. */
  readonly strokeHex: string | null;
  /** Stroke width in EMU, or `null` when no explicit width is set. */
  readonly strokeWidthEmu: number | null;
  /** Rotation in degrees. */
  readonly rotation: number;
  /** Flip state. */
  readonly flip: { horizontal: boolean; vertical: boolean };
}

/**
 * Returns the non-placeholder shapes on a layout as a render-ready
 * view. Useful for previewing brand-template decoration (corner bars,
 * background rectangles, divider lines) that would otherwise be hidden
 * because they aren't reachable through `getSlideLayoutPlaceholders`.
 *
 * Placeholders are excluded — they're better rendered through the
 * slide's own placeholder bounds (which already cascade through the
 * layout). Picture and group shapes are omitted; their bytes / nested
 * children would need the layout's relationship table to resolve.
 */
export const getSlideLayoutBackgroundShapes = (
  pres: PresentationData,
  layout: SlideLayoutData,
): ReadonlyArray<SlideLayoutBackgroundShape> => {
  const theme = getPresentationTheme(pres);
  const out: SlideLayoutBackgroundShape[] = [];
  for (const shape of layout[LAYOUT_PART].shapes) {
    if (shape.placeholderType !== null || shape.placeholderIdx !== null) continue;
    if (shape.kind !== 'shape' && shape.kind !== 'connector') continue;
    const el = shape.element;
    const pos = readPosition(el, shape.kind);
    const size = readSize(el, shape.kind);
    const bounds: ShapeBounds | null =
      pos !== null && size !== null
        ? { x: pos.x as Emu, y: pos.y as Emu, w: size.w as Emu, h: size.h as Emu }
        : null;
    const spPr = firstChildElement(el, qname('p', 'spPr', NS.pml));
    let preset: string | null = null;
    let fillHex: string | null = null;
    let strokeHex: string | null = null;
    let strokeWidthEmu: number | null = null;
    if (spPr) {
      const prstGeom = firstChildElement(spPr, qname('a', 'prstGeom', NS.dml));
      if (prstGeom) preset = getAttrValue(prstGeom, qname('', 'prst', ''));
      const solid = firstChildElement(spPr, qname('a', 'solidFill', NS.dml));
      if (solid) {
        for (const c of solid.children) {
          if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
          fillHex = resolveDrawingColor(c, theme);
          break;
        }
      }
      const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
      if (ln) {
        const w = getAttrValue(ln, qname('', 'w', ''));
        if (w !== null) {
          const n = Number.parseInt(w, 10);
          if (Number.isFinite(n)) strokeWidthEmu = n;
        }
        const lnSolid = firstChildElement(ln, qname('a', 'solidFill', NS.dml));
        if (lnSolid) {
          for (const c of lnSolid.children) {
            if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
            strokeHex = resolveDrawingColor(c, theme);
            break;
          }
        }
      }
    }
    const rotation = readRotation(el, shape.kind);
    const flip = readFlip(el, shape.kind) ?? { horizontal: false, vertical: false };
    out.push({
      kind: shape.kind,
      bounds,
      preset,
      fillHex,
      strokeHex,
      strokeWidthEmu,
      rotation,
      flip,
    });
  }
  return out;
};

/**
 * Reads the slide layout's background. Same discriminated union as
 * `getSlideBackground` for slides — renderers fall back to this when
 * the slide's own background reports `'inherit'`. Walking one further
 * level to the master is left to callers (the same shape applies).
 */
export const getSlideLayoutBackground = (layout: SlideLayoutData): SlideBackground => {
  const cSld = firstChildElement(layout[LAYOUT_PART].root, NAME_CSLD);
  if (!cSld) return { kind: 'inherit' };
  const bg = firstChildElement(cSld, qname('p', 'bg', NS.pml));
  if (!bg) return { kind: 'inherit' };
  const bgPr = firstChildElement(bg, qname('p', 'bgPr', NS.pml));
  if (!bgPr) return { kind: 'inherit' };
  for (const c of bgPr.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    switch (c.name.localName) {
      case 'solidFill': {
        for (const inner of c.children) {
          if (inner.kind !== 'element' || inner.name.namespaceURI !== NS.dml) continue;
          if (inner.name.localName === 'srgbClr') {
            const val = getAttrValue(inner, qname('', 'val', ''));
            if (val !== null) return { kind: 'solid', color: `#${val.toUpperCase()}` };
          }
          if (inner.name.localName === 'schemeClr') {
            const val = getAttrValue(inner, qname('', 'val', ''));
            if (val !== null) return { kind: 'solid', color: `scheme:${val}` };
          }
        }
        return { kind: 'solid', color: '' };
      }
      case 'gradFill':
        return { kind: 'gradient' };
      case 'pattFill':
        return { kind: 'pattern' };
      case 'blipFill':
        return { kind: 'image' };
    }
  }
  return { kind: 'inherit' };
};

/**
 * Returns the gradient stops + path when the slide carries a
 * `<p:bgPr><a:gradFill>` background. Returns `null` for any other
 * background kind. Shape identical to `getShapeGradientFill` so renderers
 * can use the same projection logic for slide backgrounds.
 */
export const getSlideBackgroundGradientFill = (slide: SlideData): GradientFillOptions | null => {
  const cSld = firstChildElement(slide[SLIDE_DOCUMENT].root, NAME_CSLD);
  if (!cSld) return null;
  const bg = firstChildElement(cSld, qname('p', 'bg', NS.pml));
  if (!bg) return null;
  const bgPr = firstChildElement(bg, qname('p', 'bgPr', NS.pml));
  if (!bgPr) return null;
  const gradFill = firstChildElement(bgPr, NAME_A_GRAD_FILL);
  if (!gradFill) return null;
  // Reuse the same algorithm `getShapeGradientFill` does. The gradFill
  // element shape is identical between shape and slide backgrounds.
  const gsLst = firstChildElement(gradFill, NAME_A_GS_LST);
  if (!gsLst) return null;
  const stops: Array<{ offset: number; color: string }> = [];
  for (const c of gsLst.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml || c.name.localName !== 'gs')
      continue;
    const posRaw = getAttrValue(c, qname('', 'pos', ''));
    if (posRaw === null) continue;
    const pos = Number.parseInt(posRaw, 10);
    if (!Number.isFinite(pos)) continue;
    const color = readColorFromContainer(c);
    if (color === null) continue;
    stops.push({ offset: pos / 100_000, color });
  }
  if (stops.length === 0) return null;
  let angleDeg = 0;
  const lin = firstChildElement(gradFill, NAME_A_LIN);
  if (lin) {
    const angRaw = getAttrValue(lin, qname('', 'ang', ''));
    if (angRaw !== null) {
      const ang = Number.parseInt(angRaw, 10);
      if (Number.isFinite(ang)) angleDeg = ang / 60_000;
    }
  }
  const pathEl = firstChildElement(gradFill, qname('a', 'path', NS.dml));
  if (pathEl) {
    const p = getAttrValue(pathEl, qname('', 'path', ''));
    const pathVal: 'circle' | 'rect' | 'shape' | null =
      p === 'circle' || p === 'rect' || p === 'shape' ? p : null;
    if (pathVal) return { stops, angleDeg, path: pathVal };
  }
  return { stops, angleDeg };
};

/**
 * Returns the pattern preset + theme-resolved colors when the slide
 * carries a `<p:bgPr><a:pattFill>` background. Returns `null` for any
 * other background kind. Shape mirrors `getShapePatternFill`.
 */
export const getSlideBackgroundPatternFill = (
  pres: PresentationData,
  slide: SlideData,
): { preset: string; foreground: string; background: string } | null => {
  const cSld = firstChildElement(slide[SLIDE_DOCUMENT].root, NAME_CSLD);
  if (!cSld) return null;
  const bg = firstChildElement(cSld, qname('p', 'bg', NS.pml));
  if (!bg) return null;
  const bgPr = firstChildElement(bg, qname('p', 'bgPr', NS.pml));
  if (!bgPr) return null;
  const pattFill = firstChildElement(bgPr, qname('a', 'pattFill', NS.dml));
  if (!pattFill) return null;
  const preset = getAttrValue(pattFill, qname('', 'prst', '')) ?? 'pct50';
  const theme = getPresentationTheme(pres);
  const colorFrom = (parentName: string, fallback: string): string => {
    const parent = firstChildElement(pattFill, qname('a', parentName, NS.dml));
    if (!parent) return fallback;
    for (const c of parent.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      const hex = resolveDrawingColor(c, theme);
      if (hex) return hex;
    }
    return fallback;
  };
  return {
    preset,
    foreground: colorFrom('fgClr', '#000000'),
    background: colorFrom('bgClr', '#FFFFFF'),
  };
};

/**
 * Returns the embedded image bytes when the slide carries a
 * `<p:bgPr><a:blipFill>` background, or `null` for any other background
 * kind (solid / gradient / pattern / inherit) or when the `r:embed`
 * relationship points at an external `r:link` target.
 *
 * Companion to `getSlideBackground`, which only reports the
 * discriminated `kind`. Renderers that want to actually paint the
 * background image (preview generators, snapshot tools) call this.
 */
export const getSlideBackgroundImageBytes = (slide: SlideData): Uint8Array | null => {
  const cSld = firstChildElement(slide[SLIDE_DOCUMENT].root, NAME_CSLD);
  if (!cSld) return null;
  const bg = firstChildElement(cSld, qname('p', 'bg', NS.pml));
  if (!bg) return null;
  const bgPr = firstChildElement(bg, qname('p', 'bgPr', NS.pml));
  if (!bgPr) return null;
  const blipFill = firstChildElement(bgPr, qname('a', 'blipFill', NS.dml));
  if (!blipFill) return null;
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) return null;
  const rEmbed = getAttrValue(blip, qname('r', 'embed', NS.officeDocRels));
  if (rEmbed === null) return null;
  const pkg = slide[INTERNAL_PACKAGE];
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return null;
  const rel = rels.items.find((r) => r.id === rEmbed);
  if (!rel || rel.targetMode === 'External') return null;
  const mediaName = rel.target.startsWith('/')
    ? partName(rel.target)
    : resolveTarget(slide[SLIDE_PART_NAME], rel.target);
  const part = pkg.getPart(mediaName);
  return part?.data ?? null;
};

/** Sets a solid fill on the slide's background. */
export const setSlideBackground = (slide: SlideData, color: string): void => {
  setSlideBackgroundXml(slide, (bgPr) => setSolidFill(bgPr, color));
};

/**
 * Sets a picture as the slide's background. Embeds `bytes` as a new
 * media part, wires a slide → image rel, and replaces any prior
 * background with a `<p:bgPr><a:blipFill><a:blip r:embed="…"/>
 * <a:stretch><a:fillRect/></a:stretch></a:blipFill></p:bgPr>` payload.
 *
 * Format is detected from magic bytes; pass `options.format` to
 * override.
 */
export const setSlideBackgroundImage = (
  slide: SlideData,
  bytes: Uint8Array,
  options: { format?: ImageFormat } = {},
): void => {
  const format = options.format ?? detectImageFormat(bytes);
  if (format === null) {
    throw new Error(
      'setSlideBackgroundImage: could not detect image format. Pass options.format explicitly.',
    );
  }
  const contentType = contentTypeForFormat(format);
  const extension = extensionForFormat(format);
  const pkg = slide[INTERNAL_PACKAGE];

  // Allocate media part name.
  let nextN = 1;
  const mediaPattern = /^\/ppt\/media\/image(\d+)\./;
  for (const p of pkg.parts) {
    const m = p.name.match(mediaPattern);
    if (m?.[1] !== undefined) {
      const n = Number.parseInt(m[1], 10);
      if (Number.isFinite(n) && n >= nextN) nextN = n + 1;
    }
  }
  const newMediaName = partName(`/ppt/media/image${nextN}.${extension}`);
  setOpcDefault(pkg, extension, contentType);
  pkg.addPart(newMediaName, contentType, bytes);

  // Slide → image rel.
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
  const newRId = nextRelId(rels.items.map((r) => r.id));
  rels.items.push({
    id: newRId,
    type: REL_TYPES.image,
    target: `../media/image${nextN}.${extension}`,
    targetMode: 'Internal',
  });
  pkg.setRels(slide[SLIDE_PART_NAME], rels);

  setSlideBackgroundXml(slide, (bgPr) => {
    const blip = elem(qname('a', 'blip', NS.dml), {
      attrs: [attr(qname('r', 'embed', NS.officeDocRels), newRId)],
    });
    const stretch = elem(qname('a', 'stretch', NS.dml), {
      children: [elem(qname('a', 'fillRect', NS.dml))],
    });
    bgPr.children.push(elem(qname('a', 'blipFill', NS.dml), { children: [blip, stretch] }));
  });
};

/** Clears any explicit slide background, restoring layout inheritance. */
export const clearSlideBackground = (slide: SlideData): void => {
  const cSld = firstChildElement(slide[SLIDE_DOCUMENT].root, NAME_CSLD);
  if (!cSld) return;
  cSld.children = cSld.children.filter(
    (c) => !(c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'bg'),
  );
  commitSlideData(slide);
  refreshSlideData(slide);
};

// ---------------------------------------------------------------------------
// Speaker notes.

const findNotesPartName = (slide: SlideData): PartName | null => {
  const rels = slide[INTERNAL_PACKAGE].getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return null;
  const notesRel = rels.items.find((r) => r.type === REL_TYPES.notesSlide);
  if (!notesRel) return null;
  return notesRel.target.startsWith('/')
    ? partName(notesRel.target)
    : resolveTarget(slide[SLIDE_PART_NAME], notesRel.target);
};

/**
 * Returns the slide's speaker notes (`null` if none). Pulls plain text
 * from the `body` placeholder; multi-line notes use `\n`.
 */
export const getSlideNotes = (slide: SlideData): string | null => {
  const notesPartName = findNotesPartName(slide);
  if (notesPartName === null) return null;
  const part = slide[INTERNAL_PACKAGE].getPart(notesPartName);
  if (part === null) return null;
  const root = parseXml(decode(part.data)).root;
  const cSld = firstChildElement(root, NAME_CSLD);
  if (!cSld) return null;
  const spTree = firstChildElement(cSld, NAME_SP_TREE);
  if (!spTree) return null;
  for (const child of spTree.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.pml) continue;
    if (child.name.localName !== 'sp') continue;
    const nvSpPr = firstChildElement(child, qname('p', 'nvSpPr', NS.pml));
    if (!nvSpPr) continue;
    const nvPr = firstChildElement(nvSpPr, qname('p', 'nvPr', NS.pml));
    if (!nvPr) continue;
    const ph = firstChildElement(nvPr, qname('p', 'ph', NS.pml));
    if (!ph) continue;
    const txBody = firstChildElement(child, qname('p', 'txBody', NS.pml));
    if (!txBody) continue;
    const lines: string[] = [];
    for (const p of txBody.children) {
      if (p.kind !== 'element' || p.name.namespaceURI !== NS.dml || p.name.localName !== 'p') {
        continue;
      }
      let line = '';
      for (const r of p.children) {
        if (r.kind !== 'element' || r.name.namespaceURI !== NS.dml || r.name.localName !== 'r') {
          continue;
        }
        for (const tElement of r.children) {
          if (
            tElement.kind === 'element' &&
            tElement.name.namespaceURI === NS.dml &&
            tElement.name.localName === 't'
          ) {
            for (const tc of tElement.children) {
              if (tc.kind === 'text') line += tc.data;
            }
          }
        }
      }
      lines.push(line);
    }
    return lines.join('\n');
  }
  return null;
};

/**
 * Sets the slide's speaker notes. Creates the `notesSlide` part and
 * wires up the rels (slide ↔ notesSlide ↔ notesMaster) on first call;
 * subsequent calls just replace the body placeholder text.
 */
/**
 * Returns every slide whose `<p:sldId show="0">` flag is set —
 * complement of `getVisibleSlides`. Useful for audit UIs and
 * batch-unhide operations.
 */
export const getHiddenSlides = (pres: PresentationData): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    if (isSlideHidden(slide)) out.push(slide);
  }
  return out;
};

/**
 * Returns every slide in document order whose `<p:sldId show="0">`
 * flag is *not* set. Convenience over `getSlides(pres).filter(s =>
 * !isSlideHidden(s))` — useful when an export pipeline needs to
 * skip hidden slides without touching the `show` attribute itself.
 */
export const getVisibleSlides = (pres: PresentationData): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    if (!isSlideHidden(slide)) out.push(slide);
  }
  return out;
};

/**
 * Returns every slide carrying at least one chart graphic frame.
 * Built on `isChartShape`.
 */
export const getSlidesWithCharts = (pres: PresentationData): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    if (slide[SLIDE_SHAPES].some((s) => isChartShape(s))) out.push(slide);
  }
  return out;
};

/**
 * Returns every slide where at least two shapes have overlapping
 * bounding boxes. Built on `findOverlappingShapePairs`. Useful for
 * deck-wide layout audits — surfacing slides that may have stacked
 * or accidentally-colliding content for human review.
 */
export const getSlidesWithOverlap = (pres: PresentationData): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    if (findOverlappingShapePairs(slide).length > 0) out.push(slide);
  }
  return out;
};

/**
 * Returns every slide carrying at least one table graphic frame.
 * Built on `isTableShape`.
 */
export const getSlidesWithTables = (pres: PresentationData): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    if (slide[SLIDE_SHAPES].some((s) => isTableShape(s))) out.push(slide);
  }
  return out;
};

/**
 * Returns every slide carrying at least one image-bearing shape
 * (a `<p:pic>` picture or a regular shape with `<a:blipFill>`).
 * Built on `hasShapeImage`.
 */
export const getSlidesWithImages = (pres: PresentationData): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    if (slide[SLIDE_SHAPES].some((s) => hasShapeImage(s))) out.push(slide);
  }
  return out;
};

/**
 * Returns every slide that has at least one comment attached.
 * Convenience over `getSlides(pres).filter(s =>
 * getSlideComments(s).length > 0)`.
 */
export const getSlidesWithComments = (pres: PresentationData): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    if (getSlideComments(slide).length > 0) out.push(slide);
  }
  return out;
};

/**
 * One entry per slide with non-empty notes, carrying its 0-based
 * slide index and the notes text. Useful for "export speaker
 * notes to a separate document" workflows that need both the
 * notes and the slide they belong to.
 */
export interface PresentationNotesEntry {
  readonly slideIndex: number;
  readonly notes: string;
}

/**
 * One entry per chart in the deck, carrying both the chart and the
 * 0-based slide it was attached to.
 */
export interface PresentationChartEntry {
  readonly slideIndex: number;
  readonly chart: SlideChartData;
}

/**
 * One entry per image-bearing shape in the deck, carrying the
 * shape (picture or image-filled) and the 0-based slide it lives
 * on. Sibling of `getAllCharts` / `getAllTables`.
 */
export interface PresentationImageEntry {
  readonly slideIndex: number;
  readonly shape: SlideShapeData;
}

/**
 * Returns every image-bearing shape across the deck (pictures and
 * shapes with `<a:blipFill>`), paired with its 0-based slide
 * index. Built on `hasShapeImage`.
 */
export const getAllImages = (pres: PresentationData): ReadonlyArray<PresentationImageEntry> => {
  const out: PresentationImageEntry[] = [];
  const slides = getSlides(pres);
  for (let i = 0; i < slides.length; i++) {
    for (const shape of slides[i]![SLIDE_SHAPES]) {
      if (hasShapeImage(shape)) out.push({ slideIndex: i, shape });
    }
  }
  return out;
};

/**
 * One entry per table in the deck, carrying the table shape and
 * the 0-based slide it sits on. Sibling of `getAllCharts`.
 */
export interface PresentationTableEntry {
  readonly slideIndex: number;
  readonly table: SlideShapeData;
}

/**
 * Returns every table across every slide in the deck, paired with
 * the 0-based index of its slide. Built on `isTableShape`.
 */
export const getAllTables = (pres: PresentationData): ReadonlyArray<PresentationTableEntry> => {
  const out: PresentationTableEntry[] = [];
  const slides = getSlides(pres);
  for (let i = 0; i < slides.length; i++) {
    for (const shape of slides[i]![SLIDE_SHAPES]) {
      if (isTableShape(shape)) out.push({ slideIndex: i, table: shape });
    }
  }
  return out;
};

/**
 * Returns every chart across every slide in the deck, paired with
 * the 0-based index of its slide. Useful for chart-inventory UIs
 * and bulk chart-update pipelines.
 */
export const getAllCharts = (pres: PresentationData): ReadonlyArray<PresentationChartEntry> => {
  const out: PresentationChartEntry[] = [];
  const slides = getSlides(pres);
  for (let i = 0; i < slides.length; i++) {
    for (const c of getSlideCharts(slides[i]!)) {
      out.push({ slideIndex: i, chart: c });
    }
  }
  return out;
};

/**
 * One entry per external hyperlink found in a shape's text body,
 * carrying the URL, the linked shape, and the 0-based slide index.
 * Each hyperlinked shape is reported once (the URL of its first
 * `<a:hlinkClick>` run).
 */
export interface PresentationHyperlinkEntry {
  readonly slideIndex: number;
  readonly shape: SlideShapeData;
  readonly url: string;
}

/**
 * Returns every external hyperlink in the deck — one entry per
 * shape whose text body carries an `<a:hlinkClick>`. Useful for
 * "link audit" passes before publishing, and for building a
 * deck-wide table of contents of outbound URLs.
 */
export const getAllHyperlinks = (
  pres: PresentationData,
): ReadonlyArray<PresentationHyperlinkEntry> => {
  const out: PresentationHyperlinkEntry[] = [];
  const slides = getSlides(pres);
  for (let i = 0; i < slides.length; i++) {
    for (const shape of slides[i]![SLIDE_SHAPES]) {
      const url = getShapeHyperlink(shape);
      if (url !== null) out.push({ slideIndex: i, shape, url });
    }
  }
  return out;
};

/**
 * Returns every distinct external URL referenced by any shape in
 * the deck, in first-seen order. Sibling of `getAllHyperlinks`
 * (which keeps duplicates and slide indices). Useful for "are
 * these URLs all live?" audits where checking each URL once is
 * enough.
 */
export const getDistinctHyperlinkUrls = (pres: PresentationData): ReadonlyArray<string> => {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const slide of getSlides(pres)) {
    for (const shape of slide[SLIDE_SHAPES]) {
      const url = getShapeHyperlink(shape);
      if (url !== null && !seen.has(url)) {
        seen.add(url);
        out.push(url);
      }
    }
  }
  return out;
};

/**
 * Returns every slide carrying at least one shape with an external
 * hyperlink. Built on `findHyperlinkedShapes`. Useful for navigation
 * UIs that want to surface only the slides containing outbound URLs.
 */
export const getSlidesWithHyperlinks = (pres: PresentationData): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    if (slide[SLIDE_SHAPES].some((s) => getShapeHyperlink(s) !== null)) {
      out.push(slide);
    }
  }
  return out;
};

/**
 * Returns every slide containing at least one shape whose external
 * hyperlink matches `needle` (substring or `RegExp`). Sibling of
 * `findSlidesByText` for outbound-URL audits — e.g. "every slide
 * that links to old.docs.example.com".
 */
export const findSlidesByHyperlink = (
  pres: PresentationData,
  needle: string | RegExp,
): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    for (const shape of slide[SLIDE_SHAPES]) {
      const url = getShapeHyperlink(shape);
      if (url === null) continue;
      const hit = typeof needle === 'string' ? url.includes(needle) : needle.test(url);
      if (hit) {
        out.push(slide);
        break;
      }
    }
  }
  return out;
};

/**
 * Bulk URL migration. Re-points every shape across the deck whose
 * first hyperlink exactly equals `from` to instead point at `to`.
 * Returns the number of shapes updated. Built on
 * `setShapeHyperlink`, so each update goes through the standard
 * rels-allocation path and stays schema-valid.
 *
 * Matching is exact (case-sensitive). To migrate by pattern, use
 * `findSlidesByHyperlink` to locate slides and rewrite each shape
 * yourself.
 */
export const replaceHyperlink = (pres: PresentationData, from: string, to: string): number => {
  let n = 0;
  for (const slide of getSlides(pres)) {
    for (const shape of slide[SLIDE_SHAPES]) {
      if (getShapeHyperlink(shape) === from) {
        setShapeHyperlink(shape, to);
        n++;
      }
    }
  }
  return n;
};

/**
 * Removes every external hyperlink across the deck — useful for
 * sanitizing a template before sharing, or for stripping outbound
 * URLs from an exported preview. Returns the number of shapes
 * cleared. Each call goes through `setShapeHyperlink(_, null)`.
 */
export const clearAllHyperlinks = (pres: PresentationData): number => {
  let n = 0;
  for (const slide of getSlides(pres)) {
    for (const shape of slide[SLIDE_SHAPES]) {
      if (getShapeHyperlink(shape) !== null) {
        setShapeHyperlink(shape, null);
        n++;
      }
    }
  }
  return n;
};

/**
 * Slide-scoped sibling of `clearAllHyperlinks`. Removes every
 * external hyperlink on this slide and returns the number of
 * shapes cleared.
 */
export const clearSlideHyperlinks = (slide: SlideData): number => {
  let n = 0;
  for (const shape of slide[SLIDE_SHAPES]) {
    if (getShapeHyperlink(shape) !== null) {
      setShapeHyperlink(shape, null);
      n++;
    }
  }
  return n;
};

/**
 * One entry per comment in the deck, carrying both the comment and
 * the 0-based slide it was attached to.
 */
export interface PresentationCommentEntry {
  readonly slideIndex: number;
  readonly comment: SlideCommentData;
}

/**
 * Returns every comment across every slide in the deck, each paired
 * with the 0-based index of its slide. Useful for review-summary
 * UIs that show all annotations in one chronological list.
 */
export const getAllComments = (pres: PresentationData): ReadonlyArray<PresentationCommentEntry> => {
  const out: PresentationCommentEntry[] = [];
  const slides = getSlides(pres);
  for (let i = 0; i < slides.length; i++) {
    for (const c of getSlideComments(slides[i]!)) {
      out.push({ slideIndex: i, comment: c });
    }
  }
  return out;
};

/**
 * Returns every slide's speaker notes alongside its 0-based index.
 * Skips slides whose notes are empty / unset.
 */
export const getAllNotes = (pres: PresentationData): ReadonlyArray<PresentationNotesEntry> => {
  const out: PresentationNotesEntry[] = [];
  const slides = getSlides(pres);
  for (let i = 0; i < slides.length; i++) {
    const notes = getSlideNotes(slides[i]!);
    if (notes !== null && notes.length > 0) out.push({ slideIndex: i, notes });
  }
  return out;
};

/**
 * Returns every slide in the presentation that carries non-empty
 * speaker notes. Convenience over `getSlides(pres).filter(s =>
 * getSlideNotes(s) !== null && getSlideNotes(s) !== '')`.
 */
export const getSlidesWithNotes = (pres: PresentationData): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    const notes = getSlideNotes(slide);
    if (notes !== null && notes.length > 0) out.push(slide);
  }
  return out;
};

/**
 * Predicate sibling of `getSlideNotes`. Returns `true` when the
 * slide carries a non-empty `notesSlide` body — i.e. whatever
 * `getSlideNotes(slide)` would surface is a non-empty string.
 *
 * Cheap to call in hot loops where the caller only needs to know
 * "are there any notes here?" without materializing the text.
 */
export const hasSlideNotes = (slide: SlideData): boolean => {
  const notes = getSlideNotes(slide);
  return notes !== null && notes.length > 0;
};

/**
 * Code-point length of the slide's speaker notes, or `0` when the
 * slide has no notes. Counts via `Array.from`, so surrogate-pair
 * characters (emoji, supplementary CJK) count as 1 — matches
 * `getSlideTextLength`.
 */
export const getSlideNotesLength = (slide: SlideData): number => {
  const notes = getSlideNotes(slide);
  return notes === null ? 0 : Array.from(notes).length;
};

/**
 * Concatenated speaker notes from every slide, joined with the
 * given `separator` (defaults to a form-feed, `\f`). Slides with
 * no notes contribute the empty string. Sibling of
 * `getPresentationText` — useful for search-indexing notes
 * across a whole deck.
 */
export const getPresentationNotesText = (
  pres: PresentationData,
  separator: string = '\f',
): string => {
  const parts: string[] = [];
  for (const slide of getSlides(pres)) parts.push(getSlideNotes(slide) ?? '');
  return parts.join(separator);
};

/**
 * Total code-point length of speaker notes across every slide.
 * Sibling of `getPresentationTextLength`; counts surrogate-pair
 * characters as 1 each. Cheaper than `getPresentationNotesText`
 * when the caller only needs the size.
 */
export const getPresentationNotesLength = (pres: PresentationData): number => {
  let n = 0;
  for (const slide of getSlides(pres)) n += getSlideNotesLength(slide);
  return n;
};

/**
 * Appends `text` to the slide's existing notes on its own line.
 * Equivalent to `setSlideNotes(slide, (getSlideNotes(slide) ?? '') + '\n' + text)`,
 * minus the leading newline when there were no notes yet.
 */
export const appendSlideNotes = (slide: SlideData, text: string): void => {
  const existing = getSlideNotes(slide);
  const value = existing === null || existing.length === 0 ? text : `${existing}\n${text}`;
  setSlideNotes(slide, value);
};

export const setSlideNotes = (slide: SlideData, value: string): void => {
  const pkg = slide[INTERNAL_PACKAGE];
  const notesPartName = findNotesPartName(slide);
  if (notesPartName !== null) {
    const part = pkg.getPart(notesPartName);
    if (part === null) throw new Error(`notes rel points at missing part ${notesPartName}`);
    const doc = parseXml(decode(part.data));
    const cSld = firstChildElement(doc.root, NAME_CSLD);
    if (!cSld) throw new Error('notesSlide has no <p:cSld>');
    const spTree = firstChildElement(cSld, NAME_SP_TREE);
    if (!spTree) throw new Error('notesSlide has no <p:spTree>');
    for (const child of spTree.children) {
      if (child.kind !== 'element' || child.name.namespaceURI !== NS.pml) continue;
      if (child.name.localName !== 'sp') continue;
      const nvSpPr = firstChildElement(child, qname('p', 'nvSpPr', NS.pml));
      if (!nvSpPr) continue;
      const nvPr = firstChildElement(nvSpPr, qname('p', 'nvPr', NS.pml));
      if (!nvPr) continue;
      const ph = firstChildElement(nvPr, qname('p', 'ph', NS.pml));
      if (!ph) continue;
      const txBody = firstChildElement(child, qname('p', 'txBody', NS.pml));
      if (!txBody) continue;
      setTextBody(txBody, value);
      part.data = encode(serializeXml(doc));
      return;
    }
    throw new Error('notesSlide has no body placeholder to fill');
  }

  // Create a new notesSlide part.
  const notesMasterPart = pkg.parts.find((p) => p.contentType.endsWith('notesMaster+xml'));
  let nextN = 1;
  const pattern = /^\/ppt\/notesSlides\/notesSlide(\d+)\.xml$/;
  for (const p of pkg.parts) {
    const m = p.name.match(pattern);
    if (m?.[1] !== undefined) {
      const n = Number.parseInt(m[1], 10);
      if (Number.isFinite(n) && n >= nextN) nextN = n + 1;
    }
  }
  const notesName = partName(`/ppt/notesSlides/notesSlide${nextN}.xml`);
  const doc = buildEmptyNotesSlide(value);
  pkg.addPart(
    notesName,
    'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
    encode(serializeXml(doc)),
  );

  const notesRels = emptyRels();
  const slideBase = slide[SLIDE_PART_NAME].split('/').pop() ?? 'slide.xml';
  notesRels.items.push({
    id: 'rId1',
    type: REL_TYPES.slide,
    target: `../slides/${slideBase}`,
    targetMode: 'Internal',
  });
  if (notesMasterPart) {
    const notesMasterBase = notesMasterPart.name.split('/').pop() ?? 'notesMaster1.xml';
    notesRels.items.push({
      id: 'rId2',
      type: REL_TYPES.notesMaster,
      target: `../notesMasters/${notesMasterBase}`,
      targetMode: 'Internal',
    });
  }
  pkg.setRels(notesName, notesRels);

  const slideRels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
  const existingIds = slideRels.items.map((r) => r.id);
  let n = 1;
  while (existingIds.includes(`rId${n}`)) n++;
  slideRels.items.push({
    id: `rId${n}`,
    type: REL_TYPES.notesSlide,
    target: `../notesSlides/notesSlide${nextN}.xml`,
    targetMode: 'Internal',
  });
  pkg.setRels(slide[SLIDE_PART_NAME], slideRels);
};

/**
 * Removes the slide's speaker-notes part entirely. Drops the
 * `notesSlide` part + its `.rels`, and unwires the slide → notesSlide
 * relationship. No-op when the slide has no notes.
 *
 * The shared `notesMaster` part is left alone; other slides may still
 * reference it.
 */
export const removeSlideNotes = (slide: SlideData): void => {
  const notesPartName = findNotesPartName(slide);
  if (notesPartName === null) return;
  const pkg = slide[INTERNAL_PACKAGE];
  pkg.removePart(notesPartName);
  pkg.removePart(relsPartNameFor(notesPartName));
  const slideRels = pkg.getRels(slide[SLIDE_PART_NAME]);
  if (slideRels === null) return;
  slideRels.items = slideRels.items.filter((r) => r.type !== REL_TYPES.notesSlide);
  pkg.setRels(slide[SLIDE_PART_NAME], slideRels);
};

/**
 * Removes the speaker-notes part from every slide that has one.
 * Built on `removeSlideNotes`. Returns the number of slides
 * stripped. Useful as a privacy/sharing helper before exporting a
 * deck whose notes contain internal commentary.
 */
export const clearAllSlideNotes = (pres: PresentationData): number => {
  let n = 0;
  for (const slide of getSlides(pres)) {
    if (findNotesPartName(slide) === null) continue;
    removeSlideNotes(slide);
    n++;
  }
  return n;
};

// ---------------------------------------------------------------------------
// Shape image replacement.

// ---------------------------------------------------------------------------
// Slide size.

/**
 * Width × height of the slide canvas, in EMU. `type` is PowerPoint's
 * aspect-ratio hint (`screen4x3`, `screen16x9`, ...); the actual size
 * is always `width` × `height`.
 */
export interface SlideSize {
  readonly width: Emu;
  readonly height: Emu;
  readonly type?: string;
}

/** Returns the slide canvas size, or `null` if `presentation.xml` omits it. */
export const getSlideSize = (pres: PresentationData): SlideSize | null => {
  const pkg = pres[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (presPart === null) return null;
  const root = parseXml(decode(presPart.data)).root;
  const model = readPresentationPart(root);
  if (model.slideSize === null) return null;
  return {
    width: model.slideSize.cx as Emu,
    height: model.slideSize.cy as Emu,
    ...(model.slideSize.type !== undefined ? { type: model.slideSize.type } : {}),
  };
};

const NAME_SLD_SZ_FN = qname('p', 'sldSz', NS.pml);
const ATTR_CX = qname('', 'cx', '');
const ATTR_CY = qname('', 'cy', '');
const ATTR_TYPE = qname('', 'type', '');
const NAME_SLD_ID_LST_FN = qname('p', 'sldIdLst', NS.pml);

/**
 * Sets the slide canvas size. Creates `<p:sldSz>` when absent, replaces
 * its attributes when present. The `type` hint is preserved as given.
 *
 * Schema ordering: `<p:sldSz>` follows `<p:sldIdLst>` per ECMA-376
 * §19.2.1.26; we insert at the correct position when bootstrapping.
 */
export const setSlideSize = (pres: PresentationData, opts: SlideSize): void => {
  const pkg = pres[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (!presPart) throw new Error('presentation.xml is missing');
  const doc = parseXml(decode(presPart.data));

  let sldSz = firstChildElement(doc.root, NAME_SLD_SZ_FN);
  if (sldSz === null) {
    sldSz = elem(NAME_SLD_SZ_FN);
    const sldIdLst = firstChildElement(doc.root, NAME_SLD_ID_LST_FN);
    if (sldIdLst !== null) {
      const idx = doc.root.children.indexOf(sldIdLst);
      doc.root.children.splice(idx + 1, 0, sldSz);
    } else {
      doc.root.children.push(sldSz);
    }
  }

  sldSz.attrs = [attr(ATTR_CX, String(opts.width)), attr(ATTR_CY, String(opts.height))];
  if (opts.type !== undefined) sldSz.attrs.push(attr(ATTR_TYPE, opts.type));

  presPart.data = encode(serializeXml(doc));
};

import { emu as emuValue } from './units.ts';

/** 10in × 7.5in (`screen4x3`). */
export const SLIDE_SIZE_4_3: SlideSize = {
  width: emuValue(9144000),
  height: emuValue(6858000),
  type: 'screen4x3',
};

/** 13.333in × 7.5in (`screen16x9`) — Office 2013+ default. */
export const SLIDE_SIZE_16_9: SlideSize = {
  width: emuValue(12192000),
  height: emuValue(6858000),
  type: 'screen16x9',
};

/** 13.333in × 8.33in (`screen16x10`). */
export const SLIDE_SIZE_16_10: SlideSize = {
  width: emuValue(12192000),
  height: emuValue(7620000),
  type: 'screen16x10',
};

// ---------------------------------------------------------------------------
// Comments.
//
// Legacy schema (ECMA-376 Part 1 §19.4):
//   * One package-level `/ppt/commentAuthors.xml` holds every author.
//   * One `/ppt/comments/comment{N}.xml` per slide that has comments;
//     N matches the slide's part name.
//   * Slide rels reference the slide's comments part; presentation rels
//     reference the author list.
//
// Authors are deduped by (name, initials). `idx` allocation is per-author
// monotonic; we read each author's `lastIdx` and bump it on add.

const COMMENT_AUTHORS_PART_NAME = partName('/ppt/commentAuthors.xml');
const COMMENT_AUTHORS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml';
const COMMENTS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.comments+xml';

const slideNumberFromPartName = (name: PartName): number => {
  const m = name.match(/^\/ppt\/slides\/slide(\d+)\.xml$/);
  if (!m?.[1]) {
    throw new Error(`comments: cannot derive slide number from ${name}`);
  }
  return Number.parseInt(m[1], 10);
};

const commentsPartNameForSlide = (slide: SlideData): PartName => {
  const slideN = slideNumberFromPartName(slide[SLIDE_PART_NAME]);
  return partName(`/ppt/comments/comment${slideN}.xml`);
};

const loadAuthorList = (pkg: OpcPackage): CommentAuthor[] => {
  const part = pkg.getPart(COMMENT_AUTHORS_PART_NAME);
  if (part === null) return [];
  const list = readCommentAuthorList(parseXml(decode(part.data)).root);
  return list.authors.slice();
};

const writeAuthorList = (pkg: OpcPackage, authors: ReadonlyArray<CommentAuthor>): void => {
  const doc = buildCommentAuthorListDoc(authors);
  const bytes = encode(serializeXml(doc));
  const existing = pkg.getPart(COMMENT_AUTHORS_PART_NAME);
  if (existing !== null) {
    existing.data = bytes;
    return;
  }
  pkg.addPart(COMMENT_AUTHORS_PART_NAME, COMMENT_AUTHORS_CONTENT_TYPE, bytes);
  // presentation → commentAuthors rel.
  const presRels = pkg.getRels(PRES_PART_NAME) ?? emptyRels();
  const exists = presRels.items.some(
    (r) => r.type === REL_TYPES.commentAuthors && r.target.endsWith('commentAuthors.xml'),
  );
  if (!exists) {
    presRels.items.push({
      id: nextRelId(presRels.items.map((r) => r.id)),
      type: REL_TYPES.commentAuthors,
      target: 'commentAuthors.xml',
      targetMode: 'Internal',
    });
    pkg.setRels(PRES_PART_NAME, presRels);
  }
};

const loadCommentsForSlide = (slide: SlideData): SlideComment[] => {
  const pkg = slide[INTERNAL_PACKAGE];
  const partNameValue = commentsPartNameForSlide(slide);
  const part = pkg.getPart(partNameValue);
  if (part === null) return [];
  const list = readCommentList(parseXml(decode(part.data)).root);
  return list.comments.slice();
};

const writeCommentsForSlide = (slide: SlideData, comments: ReadonlyArray<SlideComment>): void => {
  const pkg = slide[INTERNAL_PACKAGE];
  const commentsName = commentsPartNameForSlide(slide);

  if (comments.length === 0) {
    // Drop the comments part + slide → comments rel when no comments
    // remain. Leaves an empty part orphaned otherwise.
    if (pkg.getPart(commentsName) !== null) {
      pkg.removePart(commentsName);
    }
    const slideRels = pkg.getRels(slide[SLIDE_PART_NAME]);
    if (slideRels !== null) {
      const before = slideRels.items.length;
      slideRels.items = slideRels.items.filter((r) => r.type !== REL_TYPES.comments);
      if (slideRels.items.length !== before) {
        pkg.setRels(slide[SLIDE_PART_NAME], slideRels);
      }
    }
    return;
  }

  const doc = buildCommentListDoc(comments);
  const bytes = encode(serializeXml(doc));
  const existing = pkg.getPart(commentsName);
  if (existing !== null) {
    existing.data = bytes;
    return;
  }
  pkg.addPart(commentsName, COMMENTS_CONTENT_TYPE, bytes);

  const slideRels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
  const hasRel = slideRels.items.some((r) => r.type === REL_TYPES.comments);
  if (!hasRel) {
    const slideN = slideNumberFromPartName(slide[SLIDE_PART_NAME]);
    slideRels.items.push({
      id: nextRelId(slideRels.items.map((r) => r.id)),
      type: REL_TYPES.comments,
      target: `../comments/comment${slideN}.xml`,
      targetMode: 'Internal',
    });
    pkg.setRels(slide[SLIDE_PART_NAME], slideRels);
  }
};

const asCommentData = (
  slide: SlideData,
  snap: SlideComment,
  author: CommentAuthor,
): SlideCommentData => ({
  [COMMENT_SLIDE]: slide,
  [COMMENT_SNAPSHOT]: snap,
  author,
});

/**
 * Every author known to the package's `commentAuthors.xml`.
 * Returns an empty array when no author list exists.
 */
export const getCommentAuthors = (pres: PresentationData): ReadonlyArray<CommentAuthor> =>
  loadAuthorList(pres[INTERNAL_PACKAGE]);

/**
 * Returns every slide that has at least one comment by the given
 * author name. Sibling of `findCommentsByAuthor` (which returns
 * the comments themselves). Case-sensitive equality.
 */
export const findSlidesWithCommentsByAuthor = (
  pres: PresentationData,
  authorName: string,
): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    const hit = getSlideComments(slide).some((c) => c.author.name === authorName);
    if (hit) out.push(slide);
  }
  return out;
};

/**
 * Returns every distinct author who has at least one comment
 * anywhere in the deck. Deduplicates by author id; preserves
 * first-seen order. Differs from `getCommentAuthors(pres)`, which
 * surfaces every author registered in `commentAuthors.xml` even
 * when no comments reference them.
 */
export const getPresentationCommenters = (pres: PresentationData): ReadonlyArray<CommentAuthor> => {
  const seen = new Set<number>();
  const out: CommentAuthor[] = [];
  for (const slide of getSlides(pres)) {
    for (const c of getSlideComments(slide)) {
      if (seen.has(c.author.id)) continue;
      seen.add(c.author.id);
      out.push(c.author);
    }
  }
  return out;
};

/**
 * Looks up a `CommentAuthor` from `commentAuthors.xml` by display
 * name (case-sensitive equality). Returns `null` when no author has
 * that name. Sibling of `findCommentsByAuthor` — the latter returns
 * the matching comments; this returns the author handle for
 * downstream metadata reads (id, initials, color).
 */
export const findCommentAuthorByName = (
  pres: PresentationData,
  authorName: string,
): CommentAuthor | null => {
  for (const a of getCommentAuthors(pres)) {
    if (a.name === authorName) return a;
  }
  return null;
};

/**
 * Returns every comment whose author name matches `authorName`
 * exactly, across every slide in the deck. Useful for reviewer-
 * specific filters ("show me all of Alice's notes").
 *
 * Case-sensitive equality. Use `findShapesByText` (with a slide
 * predicate of your choosing) for fuzzier text matching against
 * comment bodies.
 */
export const findCommentsByAuthor = (
  pres: PresentationData,
  authorName: string,
): ReadonlyArray<SlideCommentData> => {
  const out: SlideCommentData[] = [];
  for (const slide of getSlides(pres)) {
    for (const c of getSlideComments(slide)) {
      if (c.author.name === authorName) out.push(c);
    }
  }
  return out;
};

/**
 * Returns every comment whose text matches `needle` (substring or
 * `RegExp`) across the whole deck. Sibling of `findCommentsByAuthor`
 * — useful for "find every comment that mentions X" reviewer flows.
 */
export const findCommentsByText = (
  pres: PresentationData,
  needle: string | RegExp,
): ReadonlyArray<SlideCommentData> => {
  const out: SlideCommentData[] = [];
  for (const slide of getSlides(pres)) {
    for (const c of getSlideComments(slide)) {
      const text = c[COMMENT_SNAPSHOT].text;
      const hit = typeof needle === 'string' ? text.includes(needle) : needle.test(text);
      if (hit) out.push(c);
    }
  }
  return out;
};

/**
 * Returns every comment whose `@dt` timestamp is **strictly after**
 * `since` (an ISO-8601 string or `Date`). Comments missing a `dt`
 * are skipped. Sibling of `findCommentsByText` / `findCommentsByAuthor`
 * — useful for "what's new since my last review?" surfaces.
 */
export const findCommentsAfter = (
  pres: PresentationData,
  since: string | Date,
): ReadonlyArray<SlideCommentData> => {
  const threshold = typeof since === 'string' ? Date.parse(since) : since.getTime();
  if (Number.isNaN(threshold)) return [];
  const out: SlideCommentData[] = [];
  for (const slide of getSlides(pres)) {
    for (const c of getSlideComments(slide)) {
      const dt = c[COMMENT_SNAPSHOT].dt;
      if (dt === null) continue;
      const t = Date.parse(dt);
      if (Number.isNaN(t)) continue;
      if (t > threshold) out.push(c);
    }
  }
  return out;
};

/**
 * Returns every comment whose `@dt` timestamp is **strictly before**
 * `until` (an ISO-8601 string or `Date`). Comments missing a `dt`
 * are skipped. Sibling of `findCommentsAfter`.
 */
export const findCommentsBefore = (
  pres: PresentationData,
  until: string | Date,
): ReadonlyArray<SlideCommentData> => {
  const threshold = typeof until === 'string' ? Date.parse(until) : until.getTime();
  if (Number.isNaN(threshold)) return [];
  const out: SlideCommentData[] = [];
  for (const slide of getSlides(pres)) {
    for (const c of getSlideComments(slide)) {
      const dt = c[COMMENT_SNAPSHOT].dt;
      if (dt === null) continue;
      const t = Date.parse(dt);
      if (Number.isNaN(t)) continue;
      if (t < threshold) out.push(c);
    }
  }
  return out;
};

const datedComments = (
  pres: PresentationData,
): ReadonlyArray<{ comment: SlideCommentData; t: number }> => {
  const out: { comment: SlideCommentData; t: number }[] = [];
  for (const slide of getSlides(pres)) {
    for (const c of getSlideComments(slide)) {
      const dt = c[COMMENT_SNAPSHOT].dt;
      if (dt === null) continue;
      const t = Date.parse(dt);
      if (!Number.isNaN(t)) out.push({ comment: c, t });
    }
  }
  return out;
};

/**
 * Returns every comment carrying a parseable `@dt`, sorted oldest
 * to newest. Comments without a date are omitted. Use `.at(0)` for
 * the oldest and `.at(-1)` for the newest.
 */
export const getCommentsSortedByDate = (
  pres: PresentationData,
): ReadonlyArray<SlideCommentData> => {
  const dated = [...datedComments(pres)];
  dated.sort((a, b) => a.t - b.t);
  return dated.map((d) => d.comment);
};

/**
 * Returns the distinct authors who commented on this slide, in
 * first-seen order. Dedupes by author id. Sibling of
 * `getPresentationCommenters` for a slide-scoped reviewer roster.
 */
export const getSlideCommentAuthors = (slide: SlideData): ReadonlyArray<CommentAuthor> => {
  const seen = new Set<number>();
  const out: CommentAuthor[] = [];
  for (const c of getSlideComments(slide)) {
    if (seen.has(c.author.id)) continue;
    seen.add(c.author.id);
    out.push(c.author);
  }
  return out;
};

/**
 * Returns every comment attached to the slide, with the author already
 * resolved. The list is read-only — use `addSlideComment` /
 * `removeSlideComment` to mutate.
 */
export const getSlideComments = (slide: SlideData): ReadonlyArray<SlideCommentData> => {
  const pkg = slide[INTERNAL_PACKAGE];
  const authors = loadAuthorList(pkg);
  const authorById = new Map<number, CommentAuthor>();
  for (const a of authors) authorById.set(a.id, a);

  const comments = loadCommentsForSlide(slide);
  const out: SlideCommentData[] = [];
  for (const snap of comments) {
    const author = authorById.get(snap.authorId);
    if (!author) {
      // Comment references an unknown author — surface a synthetic
      // placeholder rather than dropping the comment silently.
      out.push(
        asCommentData(slide, snap, {
          id: snap.authorId,
          name: '',
          initials: '',
          lastIdx: snap.idx,
          clrIdx: null,
        }),
      );
      continue;
    }
    out.push(asCommentData(slide, snap, author));
  }
  return out;
};

/**
 * Adds a comment to the slide. Returns the new comment handle.
 *
 * Author handling: if an author with the given `name`+`initials` already
 * exists in `commentAuthors.xml`, the existing record is reused (and its
 * `lastIdx` is bumped). Otherwise a new author is allocated. `initials`
 * defaults to the first character of `name`.
 *
 * `position` is in EMU; pass `null` to omit the `<p:pos>` element.
 * `date` defaults to the current time.
 */
export const addSlideComment = (
  slide: SlideData,
  opts: {
    author: { name: string; initials?: string };
    text: string;
    position?: CommentPosition | null;
    date?: Date;
  },
): SlideCommentData => {
  const pkg = slide[INTERNAL_PACKAGE];
  const initials =
    opts.author.initials ?? (opts.author.name.length > 0 ? opts.author.name.charAt(0) : '?');

  const authors = loadAuthorList(pkg);
  let author = authors.find((a) => a.name === opts.author.name && a.initials === initials);
  if (!author) {
    let maxId = -1;
    for (const a of authors) if (a.id > maxId) maxId = a.id;
    author = {
      id: maxId + 1,
      name: opts.author.name,
      initials,
      lastIdx: 0,
      clrIdx: null,
    };
    authors.push(author);
  }
  const newIdx = author.lastIdx + 1;
  // Bump lastIdx on the author for the persisted list.
  const updatedAuthor: CommentAuthor = { ...author, lastIdx: newIdx };
  const persistedAuthors = authors.map((a) => (a.id === author!.id ? updatedAuthor : a));
  writeAuthorList(pkg, persistedAuthors);

  const dt = (opts.date ?? new Date()).toISOString();
  const snap: SlideComment = {
    authorId: updatedAuthor.id,
    idx: newIdx,
    dt,
    text: opts.text,
    position: opts.position ?? null,
  };

  const comments = loadCommentsForSlide(slide);
  comments.push(snap);
  writeCommentsForSlide(slide, comments);

  return asCommentData(slide, snap, updatedAuthor);
};

/**
 * Removes the comment from its slide's comments part. If the comment
 * was the last one on the slide, the comments part and the
 * slide → comments rel are also removed. The author entry in
 * `commentAuthors.xml` is left intact (an author may have comments on
 * other slides).
 */
export const removeSlideComment = (comment: SlideCommentData): void => {
  const slide = comment[COMMENT_SLIDE];
  const target = comment[COMMENT_SNAPSHOT];
  const remaining = loadCommentsForSlide(slide).filter(
    (c) => !(c.authorId === target.authorId && c.idx === target.idx),
  );
  writeCommentsForSlide(slide, remaining);
};

/**
 * Strips every comment from every slide in the deck. Returns the
 * number of comments removed. Built on `writeCommentsForSlide`,
 * so each slide's modern comment part is rewritten with an empty
 * list. The `commentAuthors.xml` registry is left intact for any
 * caller that still needs author identity.
 *
 * Useful as a sanitizer before sharing a draft externally — pairs
 * with `clearAllSlideNotes` for a "remove reviewer chatter" pass.
 */
export const clearAllSlideComments = (pres: PresentationData): number => {
  let n = 0;
  for (const slide of getSlides(pres)) n += clearSlideComments(slide);
  return n;
};

/**
 * Slide-scoped sibling of `clearAllSlideComments`. Removes every
 * comment on the given slide and returns the number removed.
 */
export const clearSlideComments = (slide: SlideData): number => {
  const comments = loadCommentsForSlide(slide);
  if (comments.length === 0) return 0;
  writeCommentsForSlide(slide, []);
  return comments.length;
};

// Accessors over CommentAuthor / SlideCommentData for tree-shake convenience.

export const getCommentAuthor = (comment: SlideCommentData): CommentAuthor => comment.author;
export const getCommentText = (comment: SlideCommentData): string => comment[COMMENT_SNAPSHOT].text;
export const getCommentDate = (comment: SlideCommentData): string | null =>
  comment[COMMENT_SNAPSHOT].dt;
export const getCommentPosition = (comment: SlideCommentData): CommentPosition | null =>
  comment[COMMENT_SNAPSHOT].position;

/**
 * Returns the slide that owns this comment. Counterpart to
 * `getShapeSlide(shape)` — handy when filtering across the deck
 * with `findCommentsByAuthor` and the caller needs to know which
 * slide each hit came from.
 */
export const getCommentSlide = (comment: SlideCommentData): SlideData => comment[COMMENT_SLIDE];

// ---------------------------------------------------------------------------

/**
 * Replaces a picture's media with `bytes`. Same-format replacements
 * write in place; cross-format replacements allocate a new media part
 * and repoint the rel. The original geometry — crop, sizing, transform —
 * is preserved.
 */
export const setShapeImage = (
  shape: SlideShapeData,
  bytes: Uint8Array,
  options: { format?: ImageFormat } = {},
): void => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') {
    throw new Error(
      `setShapeImage only works on picture shapes; ${shape[SHAPE_SNAPSHOT].kind} is not one`,
    );
  }
  const format = options.format ?? detectImageFormat(bytes);
  if (format === null) {
    throw new Error(
      'setShapeImage: could not detect image format. Pass options.format explicitly.',
    );
  }
  const rEmbed = getPictureEmbedRId(shape[SHAPE_ELEMENT]);
  if (rEmbed === null) {
    throw new Error(`picture "${shape[SHAPE_SNAPSHOT].name}" has no r:embed`);
  }
  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
  if (rels === null) throw new Error(`slide ${slide[SLIDE_PART_NAME]} has no rels`);
  const rel = rels.items.find((r) => r.id === rEmbed);
  if (!rel) throw new Error(`slide rels missing entry for r:embed="${rEmbed}"`);

  const mediaName = rel.target.startsWith('/')
    ? partName(rel.target)
    : resolveTarget(slide[SLIDE_PART_NAME], rel.target);
  const newExtension = extensionForFormat(format);
  const newContentType = contentTypeForFormat(format);
  const dotIdx = mediaName.lastIndexOf('.');
  const currentExtension = dotIdx >= 0 ? mediaName.slice(dotIdx + 1).toLowerCase() : '';

  if (currentExtension === newExtension) {
    const part = pkg.getPart(mediaName);
    if (!part) throw new Error(`media part missing: ${mediaName}`);
    part.data = bytes;
    part.contentType = newContentType;
    return;
  }

  let nextN = 1;
  const mediaPathRegex = /^\/ppt\/media\/image(\d+)\./;
  for (const p of pkg.parts) {
    const m = p.name.match(mediaPathRegex);
    if (m?.[1] !== undefined) {
      const num = Number.parseInt(m[1], 10);
      if (Number.isFinite(num) && num >= nextN) nextN = num + 1;
    }
  }
  const newPartName = partName(`/ppt/media/image${nextN}.${newExtension}`);
  const hasDefault = pkg.contentTypes.defaults.some(
    (d) => d.extension.toLowerCase() === newExtension,
  );
  if (!hasDefault) {
    pkg.contentTypes.defaults.push({ extension: newExtension, contentType: newContentType });
  }
  pkg.addPart(newPartName, newContentType, bytes);
  rel.target = `../media/image${nextN}.${newExtension}`;
  pkg.setRels(slide[SLIDE_PART_NAME], rels);
};

// ---------------------------------------------------------------------------
// Shape click action — `<a:hlinkClick>` on the shape's cNvPr.
//
// Two flavors today: open a URL (External rel) or jump to another slide
// in this deck (Internal rel + `action="ppaction://hlinksldjump"`).
//
// PowerPoint also supports preset actions like `nextslide`, `prevslide`,
// `firstslide`, `lastslide`, but they're niche enough to defer until a
// concrete user need shows up.

/** What clicking the shape should do. */
export type ShapeClickAction =
  | { readonly kind: 'url'; readonly url: string }
  | { readonly kind: 'slide'; readonly slide: SlideData }
  | { readonly kind: 'nextSlide' }
  | { readonly kind: 'prevSlide' }
  | { readonly kind: 'firstSlide' }
  | { readonly kind: 'lastSlide' };

const NAME_HLINK_CLICK_FN = qname('a', 'hlinkClick', NS.dml);

// cNvPr lives at different paths depending on shape kind. Returns null
// for kinds we don't know how to navigate yet (groups, etc.).
const findCNvPr = (shape: SlideShapeData): XmlElement | null => {
  const root = shape[SHAPE_ELEMENT];
  const kind = shape[SHAPE_SNAPSHOT].kind;
  const wrapperName =
    kind === 'shape'
      ? 'nvSpPr'
      : kind === 'picture'
        ? 'nvPicPr'
        : kind === 'connector'
          ? 'nvCxnSpPr'
          : kind === 'graphicFrame'
            ? 'nvGraphicFramePr'
            : null;
  if (wrapperName === null) return null;
  const wrapper = firstChildElement(root, qname('p', wrapperName, NS.pml));
  if (!wrapper) return null;
  return firstChildElement(wrapper, qname('p', 'cNvPr', NS.pml));
};

const removeExistingHlinkClick = (cNvPr: XmlElement): void => {
  cNvPr.children = cNvPr.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        c.name.localName === 'hlinkClick'
      ),
  );
};

const findExistingHyperlinkRel = (
  rels: ReturnType<OpcPackage['getRels']>,
  url: string,
): string | null => {
  if (rels === null) return null;
  const existing = rels.items.find(
    (rl) => rl.type === REL_TYPES.hyperlink && rl.target === url && rl.targetMode === 'External',
  );
  return existing?.id ?? null;
};

/**
 * Reads the click action attached to the shape's cNvPr, or `null` if
 * none. Mirrors `setShapeClickAction`:
 *
 *   - `{ kind: 'url', url }`     — `hyperlink` rel + targetMode=External
 *   - `{ kind: 'slide', slide }` — `slide` rel + `ppaction://hlinksldjump`
 *   - `{ kind: 'nextSlide' | 'prevSlide' | 'firstSlide' | 'lastSlide' }`
 *     — preset show-navigation `ppaction`.
 *
 * For `kind: 'slide'`, the matching slide is resolved by part name.
 * Returns `null` for unknown `ppaction` strings.
 */
export const getShapeClickAction = (shape: SlideShapeData): ShapeClickAction | null => {
  const cNvPr = findCNvPr(shape);
  if (!cNvPr) return null;
  const hlink = firstChildElement(cNvPr, NAME_HLINK_CLICK_FN);
  if (!hlink) return null;
  const action = getAttrValue(hlink, qname('', 'action', ''));
  const rId = getAttrValue(hlink, qname('r', 'id', NS.officeDocRels));

  if (action === 'ppaction://hlinkshowjump?jump=nextslide') return { kind: 'nextSlide' };
  if (action === 'ppaction://hlinkshowjump?jump=previousslide') return { kind: 'prevSlide' };
  if (action === 'ppaction://hlinkshowjump?jump=firstslide') return { kind: 'firstSlide' };
  if (action === 'ppaction://hlinkshowjump?jump=lastslide') return { kind: 'lastSlide' };

  if (rId !== null && rId !== '') {
    const slide = shape[SHAPE_SLIDE];
    const pkg = slide[INTERNAL_PACKAGE];
    const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
    if (!rels) return null;
    const rel = rels.items.find((r) => r.id === rId);
    if (!rel) return null;

    if (action === 'ppaction://hlinksldjump' && rel.type === REL_TYPES.slide) {
      // Resolve to the SlideData of the target slide.
      const targetPartName = rel.target.startsWith('/')
        ? partName(rel.target)
        : resolveTarget(slide[SLIDE_PART_NAME], rel.target);
      const pres: PresentationData = { [INTERNAL_PACKAGE]: pkg, _slidesCache: null };
      for (const candidate of getSlides(pres)) {
        if (candidate[SLIDE_PART_NAME] === targetPartName) {
          return { kind: 'slide', slide: candidate };
        }
      }
      return null;
    }
    if (rel.type === REL_TYPES.hyperlink && rel.targetMode === 'External') {
      return { kind: 'url', url: rel.target };
    }
  }
  return null;
};

/**
 * Sets (or clears) the click action on the shape. Side effects:
 *
 *   - For `kind: 'url'`, a `hyperlink` rel is added (or reused) on the
 *     slide's rels with `targetMode="External"`. `<a:hlinkClick r:id=…/>`
 *     points at it.
 *   - For `kind: 'slide'`, a `slide` rel is added pointing at the
 *     target slide's part. The `<a:hlinkClick>` carries
 *     `action="ppaction://hlinksldjump"`.
 *   - For the preset navigations (`nextSlide`, `prevSlide`, ...), no rel
 *     is allocated; just the `action` attribute carries the preset.
 *   - `null` removes any existing `<a:hlinkClick>`.
 *
 * The shape must be one of `shape | picture | connector | graphicFrame`.
 * Groups don't carry their own click action in our model.
 */
export const setShapeClickAction = (
  shape: SlideShapeData,
  action: ShapeClickAction | null,
): void => {
  const cNvPr = findCNvPr(shape);
  if (!cNvPr) {
    throw new Error(
      `setShapeClickAction: ${shape[SHAPE_SNAPSHOT].kind} shape has no cNvPr to attach to`,
    );
  }

  removeExistingHlinkClick(cNvPr);

  if (action === null) {
    commitAndRefresh(shape);
    return;
  }

  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];

  let rId: string | null = null;
  let actionAttr: string | null = null;

  switch (action.kind) {
    case 'url': {
      const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
      const reused = findExistingHyperlinkRel(rels, action.url);
      if (reused !== null) {
        rId = reused;
      } else {
        const newId = nextRelId(rels.items.map((r) => r.id));
        rels.items.push({
          id: newId,
          type: REL_TYPES.hyperlink,
          target: action.url,
          targetMode: 'External',
        });
        pkg.setRels(slide[SLIDE_PART_NAME], rels);
        rId = newId;
      }
      break;
    }
    case 'slide': {
      const target = action.slide[SLIDE_PART_NAME];
      const targetBase = basename(target);
      const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
      const existing = rels.items.find(
        (rl) =>
          rl.type === REL_TYPES.slide &&
          rl.target === `../slides/${targetBase}` &&
          rl.targetMode === 'Internal',
      );
      if (existing) {
        rId = existing.id;
      } else {
        const newId = nextRelId(rels.items.map((r) => r.id));
        rels.items.push({
          id: newId,
          type: REL_TYPES.slide,
          target: `../slides/${targetBase}`,
          targetMode: 'Internal',
        });
        pkg.setRels(slide[SLIDE_PART_NAME], rels);
        rId = newId;
      }
      actionAttr = 'ppaction://hlinksldjump';
      break;
    }
    case 'nextSlide':
      actionAttr = 'ppaction://hlinkshowjump?jump=nextslide';
      break;
    case 'prevSlide':
      actionAttr = 'ppaction://hlinkshowjump?jump=previousslide';
      break;
    case 'firstSlide':
      actionAttr = 'ppaction://hlinkshowjump?jump=firstslide';
      break;
    case 'lastSlide':
      actionAttr = 'ppaction://hlinkshowjump?jump=lastslide';
      break;
  }

  const attrs = [] as Array<ReturnType<typeof attr>>;
  if (rId !== null) attrs.push(attr(qname('r', 'id', NS.officeDocRels), rId));
  else attrs.push(attr(qname('r', 'id', NS.officeDocRels), ''));
  if (actionAttr !== null) attrs.push(attr(qname('', 'action', ''), actionAttr));

  cNvPr.children.push(
    elem(NAME_HLINK_CLICK_FN, {
      attrs,
    }),
  );

  commitAndRefresh(shape);
};

// ---------------------------------------------------------------------------
// Charts.
//
// Authoring path for ChartML (`/ppt/charts/chart{N}.xml`) + the embedded
// `/ppt/embeddings/Microsoft_Excel_Worksheet{N}.xlsx` workbook that
// PowerPoint requires for the "Edit data" action to work. See plan §P9
// and §Risks for the scope constraints.
//
// Public surface is intentionally narrow: one `addSlideChart` entry point
// that takes a typed `ChartSpec`. The internal layer handles the chart
// XML, the embedded xlsx ZIP, and all the relationship wiring.

const CHART_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const EMBEDDED_XLSX_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const allocateChartIndex = (pkg: OpcPackage): number => {
  let next = 1;
  const re = /^\/ppt\/charts\/chart(\d+)\.xml$/;
  for (const p of pkg.parts) {
    const m = p.name.match(re);
    if (m?.[1] !== undefined) {
      const n = Number.parseInt(m[1], 10);
      if (Number.isFinite(n) && n >= next) next = n + 1;
    }
  }
  return next;
};

const NAME_GRAPHIC_FRAME = qname('p', 'graphicFrame', NS.pml);
const NAME_NV_GRAPHIC_FRAME_PR = qname('p', 'nvGraphicFramePr', NS.pml);
const NAME_C_NV_PR_FN = qname('p', 'cNvPr', NS.pml);
const NAME_C_NV_GRAPHIC_FRAME_PR = qname('p', 'cNvGraphicFramePr', NS.pml);
const NAME_NV_PR = qname('p', 'nvPr', NS.pml);
const NAME_XFRM = qname('p', 'xfrm', NS.pml);
const NAME_OFF = qname('a', 'off', NS.dml);
const NAME_EXT = qname('a', 'ext', NS.dml);
const NAME_GRAPHIC = qname('a', 'graphic', NS.dml);
const NAME_GRAPHIC_DATA = qname('a', 'graphicData', NS.dml);
const NAME_C_CHART = qname('c', 'chart', NS.chart);

const buildChartGraphicFrame = (opts: {
  id: number;
  name: string;
  x: Emu;
  y: Emu;
  w: Emu;
  h: Emu;
  rEmbed: string;
}): XmlElement => {
  const cNvPr = elem(NAME_C_NV_PR_FN, {
    attrs: [attr(qname('', 'id', ''), String(opts.id)), attr(qname('', 'name', ''), opts.name)],
  });
  const nvGraphicFramePr = elem(NAME_NV_GRAPHIC_FRAME_PR, {
    children: [cNvPr, elem(NAME_C_NV_GRAPHIC_FRAME_PR), elem(NAME_NV_PR)],
  });
  const off = elem(NAME_OFF, {
    attrs: [attr(qname('', 'x', ''), String(opts.x)), attr(qname('', 'y', ''), String(opts.y))],
  });
  const ext = elem(NAME_EXT, {
    attrs: [attr(qname('', 'cx', ''), String(opts.w)), attr(qname('', 'cy', ''), String(opts.h))],
  });
  const xfrm = elem(NAME_XFRM, { children: [off, ext] });
  const chartRef = elem(NAME_C_CHART, {
    prefixDecls: new Map([
      ['c', NS.chart],
      ['r', NS.officeDocRels],
    ]),
    attrs: [attr(qname('r', 'id', NS.officeDocRels), opts.rEmbed)],
  });
  const graphicData = elem(NAME_GRAPHIC_DATA, {
    attrs: [attr(qname('', 'uri', ''), NS.chart)],
    children: [chartRef],
  });
  const graphic = elem(NAME_GRAPHIC, { children: [graphicData] });
  return elem(NAME_GRAPHIC_FRAME, { children: [nvGraphicFramePr, xfrm, graphic] });
};

const setOpcDefault = (pkg: OpcPackage, extension: string, contentType: string): void => {
  const has = pkg.contentTypes.defaults.some((d) => d.extension.toLowerCase() === extension);
  if (!has) pkg.contentTypes.defaults.push({ extension, contentType });
};

/**
 * Adds a chart to the slide. Returns the new shape handle (kind
 * `graphicFrame`). Supported chart kinds today: `bar`, `column`,
 * `line`, `pie` — see `ChartSpec.kind`.
 *
 * Side effects:
 *
 *   - Allocates `/ppt/charts/chart{N}.xml` for the chart definition.
 *   - Allocates `/ppt/embeddings/Microsoft_Excel_Worksheet{N}.xlsx` as
 *     a placeholder workbook (single sheet, header row + one row per
 *     category). PowerPoint reads the inline `<c:strCache>` /
 *     `<c:numCache>` so the workbook is for "Edit data" only.
 *   - Slide → chart and chart → workbook rels are wired with fresh rIds.
 *   - `<a:graphicFrame>` is appended to the slide's `<p:spTree>`.
 *
 * Constraints:
 *
 *   - `pie` charts require exactly one series.
 *   - All series should have at most `categories.length` values; missing
 *     values are treated as blanks (gaps in the visualization).
 */
export const addSlideChart = (
  slide: SlideData,
  opts: {
    spec: ChartSpec;
    x: Emu;
    y: Emu;
    w: Emu;
    h: Emu;
    name?: string;
  },
): SlideShapeData => {
  const pkg = slide[INTERNAL_PACKAGE];
  const chartN = allocateChartIndex(pkg);
  const chartPartName = partName(`/ppt/charts/chart${chartN}.xml`);
  const xlsxPartName = partName(`/ppt/embeddings/Microsoft_Excel_Worksheet${chartN}.xlsx`);

  // Build the embedded xlsx bytes. Each row in the sheet corresponds to
  // one category; header row carries the series names.
  const xlsxRows = opts.spec.categories.map((label, i) => ({
    label,
    values: opts.spec.series.map((s) => s.values[i] ?? null),
  }));
  const xlsxBytes = buildEmbeddedXlsx(
    opts.spec.series.map((s) => s.name),
    xlsxRows,
  );

  // Build the chart XML and serialize.
  const chartDoc = buildChartSpaceDoc(opts.spec);
  const chartBytes = encode(serializeXml(chartDoc));

  // Add the chart part + its rel → embedded xlsx.
  pkg.addPart(chartPartName, CHART_CONTENT_TYPE, chartBytes);

  // The xlsx is a binary part; xlsx is already an OPC zip so we add a
  // Content_Types override (no Default, since `.xlsx` shouldn't override
  // unrelated archive entries even though there's only one such part
  // here in practice).
  pkg.addPart(xlsxPartName, EMBEDDED_XLSX_CONTENT_TYPE, xlsxBytes);

  // Make sure `.rels` is a recognized Default (it always is by the time
  // we get here, but be defensive for new packages).
  setOpcDefault(pkg, 'rels', 'application/vnd.openxmlformats-package.relationships+xml');

  const chartRels = emptyRels();
  chartRels.items.push({
    id: 'rId1',
    type: REL_TYPES.package,
    target: `../embeddings/Microsoft_Excel_Worksheet${chartN}.xlsx`,
    targetMode: 'Internal',
  });
  pkg.setRels(chartPartName, chartRels);

  // Slide → chart rel.
  const slideRels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
  const slideChartRId = nextRelId(slideRels.items.map((r) => r.id));
  slideRels.items.push({
    id: slideChartRId,
    type: REL_TYPES.chart,
    target: `../charts/chart${chartN}.xml`,
    targetMode: 'Internal',
  });
  pkg.setRels(slide[SLIDE_PART_NAME], slideRels);

  // Build and append the <p:graphicFrame> wrapper.
  const frame = buildChartGraphicFrame({
    id: nextShapeId(slide),
    name: opts.name ?? `Chart ${chartN}`,
    x: opts.x,
    y: opts.y,
    w: opts.w,
    h: opts.h,
    rEmbed: slideChartRId,
  });
  return appendAndReturnNewShape(slide, frame);
};

// Re-export chart types for consumers.
export type { ChartKind, ChartSeries, ChartSpec };

// ---------------------------------------------------------------------------
// Table cell access.
//
// `addSlideTable` builds the cell tree under `<a:graphic>/<a:graphicData>/
// <a:tbl>`. These helpers let callers reach individual `<a:tc>` cells to
// re-fill text, paint backgrounds, or align contents without rebuilding
// the table.

export type { TableCellData };

const NAME_A_GRAPHIC_TBL = qname('a', 'graphic', NS.dml);
const NAME_A_GRAPHIC_DATA_TBL = qname('a', 'graphicData', NS.dml);
const NAME_A_TBL = qname('a', 'tbl', NS.dml);
const NAME_A_TC_PR = qname('a', 'tcPr', NS.dml);
const NAME_A_TX_BODY_TBL = qname('a', 'txBody', NS.dml);

const findTblElement = (shape: SlideShapeData): XmlElement | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'graphicFrame') return null;
  const graphic = firstChildElement(shape[SHAPE_ELEMENT], NAME_A_GRAPHIC_TBL);
  if (!graphic) return null;
  const graphicData = firstChildElement(graphic, NAME_A_GRAPHIC_DATA_TBL);
  if (!graphicData) return null;
  return firstChildElement(graphicData, NAME_A_TBL);
};

/**
 * `true` when the shape is a `<p:graphicFrame>` wrapping `<a:tbl>` —
 * i.e. a table. Sharper than `getShapeKind(shape) === 'graphicFrame'`,
 * which also matches charts and SmartArt frames.
 */
export const isTableShape = (shape: SlideShapeData): boolean => findTblElement(shape) !== null;

/**
 * `true` when the shape is a `<p:graphicFrame>` wrapping a chart
 * reference (`<c:chart>`). Charts, tables, and SmartArt all share
 * the graphic-frame kind; this predicate filters down to charts only.
 */
export const isChartShape = (shape: SlideShapeData): boolean =>
  resolveChartPartName(shape[SHAPE_SLIDE], shape) !== null;

const NAME_A_GRID_COL = qname('a', 'gridCol', NS.dml);
const ATTR_W_TBL = qname('', 'w', '');
const ATTR_H_TBL = qname('', 'h', '');

const tableRows = (tbl: XmlElement): XmlElement[] =>
  tbl.children.filter(
    (c): c is XmlElement =>
      c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'tr',
  );

const rowCells = (tr: XmlElement): XmlElement[] =>
  tr.children.filter(
    (c): c is XmlElement =>
      c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'tc',
  );

const buildCellHandle = (
  table: SlideShapeData,
  tc: XmlElement,
  row: number,
  col: number,
): TableCellData => ({
  [CELL_TABLE]: table,
  [CELL_ELEMENT]: tc,
  [CELL_ROW]: row,
  [CELL_COL]: col,
});

/**
 * Returns a 2D array of cell handles for the given table shape, in
 * row-major order. Throws if the shape isn't a table graphic frame.
 */
export const getTableCells = (
  table: SlideShapeData,
): ReadonlyArray<ReadonlyArray<TableCellData>> => {
  const tbl = findTblElement(table);
  if (!tbl) throw new Error('getTableCells: shape is not a table graphic frame');
  const rows = tableRows(tbl);
  return rows.map((tr, rowIdx) =>
    rowCells(tr).map((tc, colIdx) => buildCellHandle(table, tc, rowIdx, colIdx)),
  );
};

/**
 * Reads the table-style GUID from `<a:tbl><a:tblPr><a:tableStyleId>`.
 * PowerPoint uses GUIDs to reference built-in table styles
 * (`{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}` = "Medium Style 2 -
 * Accent 1", etc.) and theme-local styles. Returns `null` when the
 * table doesn't reference one (uses the slide's default style).
 */
export const getTableStyleId = (table: SlideShapeData): string | null => {
  const tbl = findTblElement(table);
  if (!tbl) return null;
  const tblPr = firstChildElement(tbl, qname('a', 'tblPr', NS.dml));
  if (!tblPr) return null;
  const idEl = firstChildElement(tblPr, qname('a', 'tableStyleId', NS.dml));
  if (!idEl) return null;
  let acc = '';
  for (const c of idEl.children) {
    if (c.kind === 'text' || c.kind === 'cdata') acc += c.data;
  }
  return acc.trim() || null;
};

/**
 * Reads the boolean style flags off `<a:tblPr>` — which header /
 * footer rows + columns are banded or emphasised. Mirrors the
 * `firstRow` / `bandRow` flags exposed by `addSlideTable`.
 *
 * Per ECMA-376 §17.18.95 / §21.1.3.15:
 *
 *   - `firstRow` — header row is styled differently.
 *   - `lastRow` — total row is styled differently.
 *   - `firstCol` — first column is styled differently.
 *   - `lastCol` — last column is styled differently.
 *   - `bandRow` — alternating row shading.
 *   - `bandCol` — alternating column shading.
 *
 * Renderers use these to switch on the corresponding table style
 * variant. Returns all-`false` for tables that don't author `<a:tblPr>`.
 */
export const getTableStyleFlags = (
  table: SlideShapeData,
): {
  firstRow: boolean;
  lastRow: boolean;
  firstCol: boolean;
  lastCol: boolean;
  bandRow: boolean;
  bandCol: boolean;
} => {
  const empty = {
    firstRow: false,
    lastRow: false,
    firstCol: false,
    lastCol: false,
    bandRow: false,
    bandCol: false,
  };
  const tbl = findTblElement(table);
  if (!tbl) return empty;
  const tblPr = firstChildElement(tbl, qname('a', 'tblPr', NS.dml));
  if (!tblPr) return empty;
  const readBool = (attr: string): boolean => {
    const v = getAttrValue(tblPr, qname('', attr, ''));
    return v === '1' || v === 'true';
  };
  return {
    firstRow: readBool('firstRow'),
    lastRow: readBool('lastRow'),
    firstCol: readBool('firstCol'),
    lastCol: readBool('lastCol'),
    bandRow: readBool('bandRow'),
    bandCol: readBool('bandCol'),
  };
};

/**
 * Returns the table's row + column counts. Throws when the shape
 * isn't a table graphic frame.
 */
export const getTableDimensions = (
  table: SlideShapeData,
): { readonly rows: number; readonly cols: number } => {
  const tbl = findTblElement(table);
  if (!tbl) throw new Error('getTableDimensions: shape is not a table graphic frame');
  const rows = tableRows(tbl);
  const cols = rows[0] !== undefined ? rowCells(rows[0]).length : 0;
  return { rows: rows.length, cols };
};

/**
 * Returns each column's width in EMU, in left-to-right order, as
 * declared on `<a:tblGrid>/<a:gridCol w="...">`. Missing or
 * unparseable widths default to 0. Throws when the shape isn't a
 * table graphic frame.
 */
export const getTableColumnWidths = (table: SlideShapeData): ReadonlyArray<Emu> => {
  const tbl = findTblElement(table);
  if (!tbl) throw new Error('getTableColumnWidths: shape is not a table graphic frame');
  const grid = firstChildElement(tbl, qname('a', 'tblGrid', NS.dml));
  if (!grid) return [];
  const out: Emu[] = [];
  for (const col of allChildElements(grid, NAME_A_GRID_COL)) {
    const v = getAttrValue(col, ATTR_W_TBL);
    const n = v !== null ? Number.parseInt(v, 10) : 0;
    out.push((Number.isFinite(n) ? n : 0) as Emu);
  }
  return out;
};

/**
 * Returns the table's nominal `(width, height)` derived from
 * summing the `<a:gridCol w>` and `<a:tr h>` attributes (both in
 * EMU). Useful for layout pipelines that want to know how big a
 * table really is without dereferencing the shape's `<a:xfrm>`.
 *
 * Throws when the shape isn't a table graphic frame.
 */
export const getTableSize = (
  table: SlideShapeData,
): { readonly width: Emu; readonly height: Emu } => {
  const widths = getTableColumnWidths(table);
  const heights = getTableRowHeights(table);
  const width = widths.reduce((sum, w) => sum + w, 0) as Emu;
  const height = heights.reduce((sum, h) => sum + h, 0) as Emu;
  return { width, height };
};

/**
 * Returns each row's height in EMU, in top-to-bottom order, from
 * `<a:tr h="...">`. Missing or unparseable heights default to 0.
 * Throws when the shape isn't a table graphic frame.
 */
export const getTableRowHeights = (table: SlideShapeData): ReadonlyArray<Emu> => {
  const tbl = findTblElement(table);
  if (!tbl) throw new Error('getTableRowHeights: shape is not a table graphic frame');
  const out: Emu[] = [];
  for (const tr of tableRows(tbl)) {
    const v = getAttrValue(tr, ATTR_H_TBL);
    const n = v !== null ? Number.parseInt(v, 10) : 0;
    out.push((Number.isFinite(n) ? n : 0) as Emu);
  }
  return out;
};

/**
 * Sets a single column's width on the table grid. Throws on
 * out-of-range column indices or non-table shapes. The total table
 * width is not auto-adjusted — callers are responsible for keeping
 * the sum consistent with the table's `<a:xfrm>` extent if PowerPoint
 * is to render the table without clipping.
 */
export const setTableColumnWidth = (table: SlideShapeData, col: number, width: Emu): void => {
  const tbl = requireTbl(table);
  const grid = firstChildElement(tbl, qname('a', 'tblGrid', NS.dml));
  if (!grid) throw new Error('table has no <a:tblGrid>');
  const cols = allChildElements(grid, NAME_A_GRID_COL);
  const target = cols[col];
  if (!target) throw new RangeError(`table column ${col} out of range (have ${cols.length})`);
  target.attrs = target.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'w'),
  );
  target.attrs.push(attr(ATTR_W_TBL, String(Math.round(width))));
  commitSlideData(table[SHAPE_SLIDE]);
  refreshSlideData(table[SHAPE_SLIDE]);
};

/**
 * Sets a single row's height. Throws on out-of-range row indices or
 * non-table shapes. As with `setTableColumnWidth`, the table's
 * `<a:xfrm>` extent is left to the caller.
 */
export const setTableRowHeight = (table: SlideShapeData, row: number, height: Emu): void => {
  const tbl = requireTbl(table);
  const rows = tableRows(tbl);
  const target = rows[row];
  if (!target) throw new RangeError(`table row ${row} out of range (have ${rows.length})`);
  target.attrs = target.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'h'),
  );
  target.attrs.push(attr(ATTR_H_TBL, String(Math.round(height))));
  commitSlideData(table[SHAPE_SLIDE]);
  refreshSlideData(table[SHAPE_SLIDE]);
};

/**
 * Returns the cell at `(row, col)`. Throws on out-of-range coordinates
 * or non-table shapes.
 */
export const getTableCell = (table: SlideShapeData, row: number, col: number): TableCellData => {
  const cells = getTableCells(table);
  const r = cells[row];
  if (!r) throw new RangeError(`table row ${row} out of range (have ${cells.length})`);
  const c = r[col];
  if (!c) throw new RangeError(`table column ${col} out of range in row ${row} (have ${r.length})`);
  return c;
};

const commitTableCell = (cell: TableCellData): void => {
  const shape = cell[CELL_TABLE];
  commitSlideData(shape[SHAPE_SLIDE]);
  refreshSlideData(shape[SHAPE_SLIDE]);
};

const ensureCellTxBody = (cell: TableCellData): XmlElement => {
  const tc = cell[CELL_ELEMENT];
  let txBody = firstChildElement(tc, NAME_A_TX_BODY_TBL);
  if (txBody === null) {
    txBody = elem(NAME_A_TX_BODY_TBL);
    // bodyPr lstStyle a:p — keep the canonical ordering.
    txBody.children.push(elem(qname('a', 'bodyPr', NS.dml)));
    txBody.children.push(elem(qname('a', 'lstStyle', NS.dml)));
    // Insert before <a:tcPr> per the schema.
    const tcPrIdx = tc.children.findIndex(
      (c) => c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'tcPr',
    );
    if (tcPrIdx >= 0) tc.children.splice(tcPrIdx, 0, txBody);
    else tc.children.push(txBody);
  }
  return txBody;
};

const ensureCellTcPr = (cell: TableCellData): XmlElement => {
  const tc = cell[CELL_ELEMENT];
  let tcPr = firstChildElement(tc, NAME_A_TC_PR);
  if (tcPr === null) {
    tcPr = elem(NAME_A_TC_PR);
    // <a:tcPr> is the LAST child of <a:tc>.
    tc.children.push(tcPr);
  }
  return tcPr;
};

/** Replaces a cell's text. `\n` starts a new paragraph. */
export const setTableCellText = (cell: TableCellData, text: string): void => {
  const txBody = ensureCellTxBody(cell);
  setTextBody(txBody, text);
  commitTableCell(cell);
};

/**
 * Reads the cell's merge / span attributes:
 *
 *   - `gridSpan` — number of columns this cell spans (≥1, default 1).
 *   - `rowSpan` — number of rows this cell spans (≥1, default 1).
 *   - `hMerge` — `true` when this cell is the right half of a horizontal
 *     span (it's painted by an earlier cell with `gridSpan > 1`).
 *   - `vMerge` — `true` when this cell is the bottom half of a vertical
 *     span (it's painted by an earlier cell with `rowSpan > 1`).
 *
 * Renderers should skip painting cells where `hMerge` or `vMerge` is
 * true; those cells exist only so the row/column grid stays consistent.
 */
export const getTableCellSpan = (
  cell: TableCellData,
): { gridSpan: number; rowSpan: number; hMerge: boolean; vMerge: boolean } => {
  const el = cell[CELL_ELEMENT];
  const gs = getAttrValue(el, qname('', 'gridSpan', ''));
  const rs = getAttrValue(el, qname('', 'rowSpan', ''));
  const hm = getAttrValue(el, qname('', 'hMerge', ''));
  const vm = getAttrValue(el, qname('', 'vMerge', ''));
  const parseSpan = (v: string | null): number => {
    if (v === null) return 1;
    const n = Number.parseInt(v, 10);
    return Number.isFinite(n) && n > 0 ? n : 1;
  };
  return {
    gridSpan: parseSpan(gs),
    rowSpan: parseSpan(rs),
    hMerge: hm === '1' || hm === 'true',
    vMerge: vm === '1' || vm === 'true',
  };
};

/**
 * One side of a cell's border, as read from `<a:tcPr><a:ln{L|R|T|B}>`.
 * `widthEmu` is the line width in EMU; `color` is `#RRGGBB` or `null`
 * when the border is inherited / un-styled.
 */
export interface TableCellBorder {
  readonly color: string | null;
  readonly widthEmu: number | null;
  readonly dash: string | null;
}

export interface TableCellBorders {
  readonly left: TableCellBorder | null;
  readonly right: TableCellBorder | null;
  readonly top: TableCellBorder | null;
  readonly bottom: TableCellBorder | null;
  readonly tlToBr: TableCellBorder | null;
  readonly blToTr: TableCellBorder | null;
}

/**
 * Reads the per-side borders on a cell. Returns `null` for sides with
 * no explicit `<a:ln{Side}>` element (those inherit from the table
 * style / theme). All four cardinal sides plus the two diagonals
 * (TL→BR, BL→TR) are surfaced because real templates use them all.
 */
export const getTableCellBorders = (
  pres: PresentationData,
  cell: TableCellData,
): TableCellBorders => {
  const tcPr = firstChildElement(cell[CELL_ELEMENT], NAME_A_TC_PR);
  const theme = getPresentationTheme(pres);
  const empty: TableCellBorders = {
    left: null,
    right: null,
    top: null,
    bottom: null,
    tlToBr: null,
    blToTr: null,
  };
  if (!tcPr) return empty;
  const readLn = (local: string): TableCellBorder | null => {
    const ln = firstChildElement(tcPr, qname('a', local, NS.dml));
    if (!ln) return null;
    const w = getAttrValue(ln, qname('', 'w', ''));
    const widthEmu = w !== null ? Number.parseInt(w, 10) : null;
    let color: string | null = null;
    const solid = firstChildElement(ln, qname('a', 'solidFill', NS.dml));
    if (solid) {
      for (const c of solid.children) {
        if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
        color = resolveDrawingColor(c, theme);
        break;
      }
    }
    const prstDash = firstChildElement(ln, qname('a', 'prstDash', NS.dml));
    const dash = prstDash ? getAttrValue(prstDash, qname('', 'val', '')) : null;
    return { color, widthEmu, dash };
  };
  return {
    left: readLn('lnL'),
    right: readLn('lnR'),
    top: readLn('lnT'),
    bottom: readLn('lnB'),
    tlToBr: readLn('lnTlToBr'),
    blToTr: readLn('lnBlToTr'),
  };
};

/**
 * Reads the cell's vertical text anchor (`<a:tcPr anchor="t|ctr|b"/>`)
 * — `'top'`, `'center'`, `'bottom'`, or `null` for the default
 * (`ctr` / center per the schema).
 */
export const getTableCellAnchor = (cell: TableCellData): 'top' | 'center' | 'bottom' | null => {
  const tcPr = firstChildElement(cell[CELL_ELEMENT], NAME_A_TC_PR);
  if (!tcPr) return null;
  const v = getAttrValue(tcPr, qname('', 'anchor', ''));
  if (v === 't') return 'top';
  if (v === 'ctr') return 'center';
  if (v === 'b') return 'bottom';
  return null;
};

/**
 * Reads the cell's inset margins (`<a:tcPr marL marR marT marB>`) in
 * EMU. Each side is `null` when the cell doesn't author it (renderers
 * should fall back to PowerPoint's defaults — 91440 EMU / 0.1 inch
 * for the horizontal margins, 45720 EMU for the vertical).
 */
export const getTableCellMargins = (
  cell: TableCellData,
): { left: number | null; right: number | null; top: number | null; bottom: number | null } => {
  const tcPr = firstChildElement(cell[CELL_ELEMENT], NAME_A_TC_PR);
  const empty = { left: null, right: null, top: null, bottom: null };
  if (!tcPr) return empty;
  const read = (name: string): number | null => {
    const v = getAttrValue(tcPr, qname('', name, ''));
    if (v === null) return null;
    const n = Number.parseInt(v, 10);
    return Number.isFinite(n) ? n : null;
  };
  return {
    left: read('marL'),
    right: read('marR'),
    top: read('marT'),
    bottom: read('marB'),
  };
};

/** Reads the cell's plain text (paragraphs joined with `\n`). */
export const getTableCellText = (cell: TableCellData): string => {
  const txBody = firstChildElement(cell[CELL_ELEMENT], NAME_A_TX_BODY_TBL);
  if (!txBody) return '';
  const lines: string[] = [];
  for (const p of txBody.children) {
    if (p.kind !== 'element' || p.name.namespaceURI !== NS.dml || p.name.localName !== 'p')
      continue;
    let line = '';
    for (const r of p.children) {
      if (r.kind !== 'element' || r.name.namespaceURI !== NS.dml || r.name.localName !== 'r')
        continue;
      const tEl = firstChildElement(r, qname('a', 't', NS.dml));
      if (!tEl) continue;
      for (const child of tEl.children) {
        if (child.kind === 'text' || child.kind === 'cdata') line += child.data;
      }
    }
    lines.push(line);
  }
  return lines.join('\n');
};

/** Sets a solid background color on a cell (`<a:tcPr><a:solidFill>`). */
export const setTableCellFill = (cell: TableCellData, color: string): void => {
  const tcPr = ensureCellTcPr(cell);
  setSolidFill(tcPr, color);
  commitTableCell(cell);
};

/** Removes any background fill from a cell. */
export const clearTableCellFill = (cell: TableCellData): void => {
  const tcPr = ensureCellTcPr(cell);
  clearFillImpl(tcPr);
  commitTableCell(cell);
};

/**
 * Reads the cell's solid background color. Returns `#RRGGBB` for
 * `<a:srgbClr>`, `scheme:<token>` for `<a:schemeClr>`, or `null` when
 * the cell has no fill, no `<a:tcPr>`, or the fill is non-solid
 * (gradient / pattern / image).
 */
export const getTableCellFill = (cell: TableCellData): string | null => {
  const tcPr = firstChildElement(cell[CELL_ELEMENT], NAME_A_TC_PR);
  if (!tcPr) return null;
  const solid = firstChildElement(tcPr, qname('a', 'solidFill', NS.dml));
  if (!solid) return null;
  const srgb = firstChildElement(solid, qname('a', 'srgbClr', NS.dml));
  if (srgb) {
    const v = getAttrValue(srgb, qname('', 'val', ''));
    if (v) return `#${v.toUpperCase()}`;
  }
  const scheme = firstChildElement(solid, qname('a', 'schemeClr', NS.dml));
  if (scheme) {
    const v = getAttrValue(scheme, qname('', 'val', ''));
    if (v) return `scheme:${v}`;
  }
  return null;
};

/** Applies a TextFormat to every run in the cell's text. */
export const setTableCellTextFormat = (cell: TableCellData, format: TextFormat): void => {
  const txBody = ensureCellTxBody(cell);
  applyFormatToAllRuns(txBody, format);
  commitTableCell(cell);
};

/** Sets horizontal alignment on every paragraph in the cell. */
export const setTableCellAlignment = (cell: TableCellData, align: ParagraphAlignment): void => {
  const txBody = ensureCellTxBody(cell);
  applyAlignmentToAllParagraphs(txBody, align);
  commitTableCell(cell);
};

/** Zero-based (row, col) of the cell. */
export const getTableCellPosition = (cell: TableCellData): { row: number; col: number } => ({
  row: cell[CELL_ROW],
  col: cell[CELL_COL],
});

/**
 * Reads the horizontal alignment from the cell's first paragraph
 * (`l`, `ctr`, `r`, `just`, `dist`, `justLow`, `thaiDist`). Returns
 * `null` when the cell has no `<a:txBody>` or its first paragraph
 * has no explicit `algn` attribute (PowerPoint then defaults to `l`).
 */
export const getTableCellAlignment = (cell: TableCellData): ParagraphAlignment | null => {
  const txBody = firstChildElement(cell[CELL_ELEMENT], NAME_A_TX_BODY_TBL);
  if (!txBody) return null;
  for (const p of txBody.children) {
    if (p.kind !== 'element' || p.name.namespaceURI !== NS.dml || p.name.localName !== 'p')
      continue;
    const pPr = firstChildElement(p, qname('a', 'pPr', NS.dml));
    if (!pPr) return null;
    const v = getAttrValue(pPr, qname('', 'algn', ''));
    return (v as ParagraphAlignment | null) ?? null;
  }
  return null;
};

const requireTbl = (table: SlideShapeData): XmlElement => {
  const tbl = findTblElement(table);
  if (!tbl) throw new Error('table shape is not a table graphic frame');
  return tbl;
};

const tableColumnCount = (tbl: XmlElement): number => {
  const grid = firstChildElement(tbl, qname('a', 'tblGrid', NS.dml));
  if (!grid) return 0;
  return allChildElements(grid, NAME_A_GRID_COL).length;
};

const rowDefaultHeight = (tbl: XmlElement): number => {
  // Use the average height of existing rows, or 370000 (≈ 0.4in) as a
  // sane default when the table has no rows yet.
  const rows = tableRows(tbl);
  if (rows.length === 0) return 370000;
  let sum = 0;
  let count = 0;
  for (const r of rows) {
    const h = getAttrValue(r, ATTR_H_TBL);
    if (h !== null) {
      const n = Number.parseInt(h, 10);
      if (Number.isFinite(n)) {
        sum += n;
        count++;
      }
    }
  }
  return count > 0 ? Math.round(sum / count) : 370000;
};

/**
 * Inserts a row into the table. `atIndex` is 0-based; `undefined`
 * appends at the end. `cells` supplies cell values; missing cells
 * become blank, extras are dropped. The row's height matches the
 * average of existing rows (or a 0.4in default for empty tables).
 */
export const insertTableRow = (
  table: SlideShapeData,
  atIndex?: number,
  cells: ReadonlyArray<string> = [],
): void => {
  const tbl = requireTbl(table);
  const colCount = tableColumnCount(tbl);
  const padded: string[] = [];
  for (let i = 0; i < colCount; i++) padded.push(cells[i] ?? '');
  const row = buildTableRow(padded, rowDefaultHeight(tbl));

  const rows = tableRows(tbl);
  const insertAt =
    atIndex !== undefined ? Math.max(0, Math.min(atIndex, rows.length)) : rows.length;
  if (insertAt === rows.length) {
    tbl.children.push(row);
  } else {
    const target = rows[insertAt]!;
    const idx = tbl.children.indexOf(target);
    tbl.children.splice(idx, 0, row);
  }
  commitSlideData(table[SHAPE_SLIDE]);
  refreshSlideData(table[SHAPE_SLIDE]);
};

/** Removes the row at `atIndex` from the table. Throws on out-of-range. */
export const removeTableRow = (table: SlideShapeData, atIndex: number): void => {
  const tbl = requireTbl(table);
  const rows = tableRows(tbl);
  if (atIndex < 0 || atIndex >= rows.length) {
    throw new RangeError(`removeTableRow: index ${atIndex} out of range (have ${rows.length})`);
  }
  const target = rows[atIndex]!;
  tbl.children = tbl.children.filter((c) => c !== target);
  commitSlideData(table[SHAPE_SLIDE]);
  refreshSlideData(table[SHAPE_SLIDE]);
};

/**
 * Inserts a column into the table. `atIndex` defaults to the end.
 * `widthEmu` defaults to the average of existing column widths (or
 * 914400 = 1in if the table has no columns). Existing rows get a new
 * blank cell at `atIndex`.
 */
export const insertTableColumn = (
  table: SlideShapeData,
  atIndex?: number,
  widthEmu?: number,
): void => {
  const tbl = requireTbl(table);
  const grid = firstChildElement(tbl, qname('a', 'tblGrid', NS.dml));
  if (!grid) throw new Error('table is missing <a:tblGrid>');
  const cols = allChildElements(grid, NAME_A_GRID_COL);
  const insertAt =
    atIndex !== undefined ? Math.max(0, Math.min(atIndex, cols.length)) : cols.length;

  // Default width: average of existing widths.
  let defaultWidth = widthEmu;
  if (defaultWidth === undefined) {
    let sum = 0;
    let count = 0;
    for (const col of cols) {
      const w = getAttrValue(col, ATTR_W_TBL);
      if (w !== null) {
        const n = Number.parseInt(w, 10);
        if (Number.isFinite(n)) {
          sum += n;
          count++;
        }
      }
    }
    defaultWidth = count > 0 ? Math.round(sum / count) : 914400;
  }
  const newCol = elem(NAME_A_GRID_COL, { attrs: [attr(ATTR_W_TBL, String(defaultWidth))] });
  if (insertAt === cols.length) {
    grid.children.push(newCol);
  } else {
    const target = cols[insertAt]!;
    const idx = grid.children.indexOf(target);
    grid.children.splice(idx, 0, newCol);
  }

  // Add a blank <a:tc> at the same column index in every row.
  for (const tr of tableRows(tbl)) {
    const tcs = rowCells(tr);
    const newCell = buildTableCell('');
    if (insertAt >= tcs.length) {
      tr.children.push(newCell);
    } else {
      const target = tcs[insertAt]!;
      const idx = tr.children.indexOf(target);
      tr.children.splice(idx, 0, newCell);
    }
  }

  commitSlideData(table[SHAPE_SLIDE]);
  refreshSlideData(table[SHAPE_SLIDE]);
};

/** Removes the column at `atIndex` (and the corresponding cell in every row). */
export const removeTableColumn = (table: SlideShapeData, atIndex: number): void => {
  const tbl = requireTbl(table);
  const grid = firstChildElement(tbl, qname('a', 'tblGrid', NS.dml));
  if (!grid) throw new Error('table is missing <a:tblGrid>');
  const cols = allChildElements(grid, NAME_A_GRID_COL);
  if (atIndex < 0 || atIndex >= cols.length) {
    throw new RangeError(`removeTableColumn: index ${atIndex} out of range (have ${cols.length})`);
  }
  const targetCol = cols[atIndex]!;
  grid.children = grid.children.filter((c) => c !== targetCol);
  for (const tr of tableRows(tbl)) {
    const tcs = rowCells(tr);
    if (atIndex < tcs.length) {
      tr.children = tr.children.filter((c) => c !== tcs[atIndex]);
    }
  }
  commitSlideData(table[SHAPE_SLIDE]);
  refreshSlideData(table[SHAPE_SLIDE]);
};

/**
 * A chart sitting on a slide. `shape` is the `<p:graphicFrame>`
 * wrapper; `spec` is the chart definition parsed from the linked
 * `/ppt/charts/chart{N}.xml` part. `null` `spec` means the chart uses
 * a kind we don't model (callers can fall through to pass-through).
 */
export interface SlideChartData {
  readonly shape: SlideShapeData;
  readonly spec: ChartSpec | null;
}

const NAME_A_GRAPHIC_FN = qname('a', 'graphic', NS.dml);
const NAME_A_GRAPHIC_DATA_FN = qname('a', 'graphicData', NS.dml);
const NAME_C_CHART_FN = qname('c', 'chart', NS.chart);

/**
 * Resolves the chart part backing a graphic-frame shape, or `null` if
 * the shape isn't a chart wrapper.
 */
const resolveChartPartName = (
  slide: SlideData,
  shape: SlideShapeData,
): { partName: PartName; rId: string } | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'graphicFrame') return null;
  const graphic = firstChildElement(shape[SHAPE_ELEMENT], NAME_A_GRAPHIC_FN);
  if (!graphic) return null;
  const graphicData = firstChildElement(graphic, NAME_A_GRAPHIC_DATA_FN);
  if (!graphicData) return null;
  const chartRef = firstChildElement(graphicData, NAME_C_CHART_FN);
  if (!chartRef) return null;
  const rId = getAttrValue(chartRef, qname('r', 'id', NS.officeDocRels));
  if (rId === null) return null;
  const slideRels = slide[INTERNAL_PACKAGE].getRels(slide[SLIDE_PART_NAME]);
  if (!slideRels) return null;
  const rel = slideRels.items.find((r) => r.id === rId);
  if (!rel) return null;
  const partNameValue = rel.target.startsWith('/')
    ? partName(rel.target)
    : resolveTarget(slide[SLIDE_PART_NAME], rel.target);
  return { partName: partNameValue, rId };
};

/**
 * Replaces the chart definition on an existing graphic-frame chart
 * shape. Updates the inline `<c:strCache>` / `<c:numCache>` blocks so
 * PowerPoint renders the new data without opening the embedded
 * workbook. The shape's geometry (position / size / rotation) is
 * preserved verbatim.
 *
 * The embedded xlsx is re-written too — it's what the "Edit data"
 * affordance opens. The previous workbook is replaced wholesale (no
 * attempt to preserve styles a user added through Excel).
 *
 * Pass any `ChartSpec`, including a different `kind` from the
 * original; this acts as "change my column chart to a line chart with
 * fresh data."
 */
export const setChartSpec = (chart: SlideChartData, spec: ChartSpec): void => {
  const slide = chart.shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];
  const resolved = resolveChartPartName(slide, chart.shape);
  if (!resolved) {
    throw new Error('setChartSpec: shape is not a chart graphic frame');
  }

  // Rewrite the chart XML.
  const doc = buildChartSpaceDoc(spec);
  const chartBytes = encode(serializeXml(doc));
  const chartPart = pkg.getPart(resolved.partName);
  if (!chartPart) {
    throw new Error(`setChartSpec: chart part ${resolved.partName} not found`);
  }
  chartPart.data = chartBytes;

  // Rewrite the embedded xlsx (the part chart→package rel points at).
  // Reuse the existing rel; fall back to creating a fresh xlsx part if
  // the chart had no package rel (unusual for charts we authored).
  const chartRels = pkg.getRels(resolved.partName);
  if (chartRels) {
    const xlsxRel = chartRels.items.find((r) => r.type === REL_TYPES.package);
    if (xlsxRel) {
      const xlsxName = xlsxRel.target.startsWith('/')
        ? partName(xlsxRel.target)
        : resolveTarget(resolved.partName, xlsxRel.target);
      const xlsxPart = pkg.getPart(xlsxName);
      const rows = spec.categories.map((label, i) => ({
        label,
        values: spec.series.map((s) => s.values[i] ?? null),
      }));
      const xlsxBytes = buildEmbeddedXlsx(
        spec.series.map((s) => s.name),
        rows,
      );
      if (xlsxPart) {
        xlsxPart.data = xlsxBytes;
      }
    }
  }
};

/**
 * Returns every chart on the slide, with its `ChartSpec` parsed from
 * the linked chart part. Skips graphic frames that don't carry a
 * `<c:chart>` reference (e.g. tables, diagrams).
 */
/**
 * For a graphic-frame shape that wraps a chart, returns the parsed
 * `ChartSpec`. Returns `null` when the shape isn't a chart wrapper
 * or the chart uses a kind we don't model (e.g. surface, radar).
 *
 * Convenience over `getSlideCharts(...).find((c) => c.shape === shape)`
 * when the caller already has the shape in hand (e.g. iterating
 * `getSlideShapes`).
 */
/**
 * Convenience over `getShapeChartSpec(shape)?.kind ?? null`. Returns
 * the chart's `ChartKind` ('bar', 'line', 'pie', …) when the shape
 * is a chart wrapper, or `null` for non-charts and charts whose
 * kind isn't modeled yet.
 */
export const getShapeChartKind = (shape: SlideShapeData): ChartKind | null => {
  const spec = getShapeChartSpec(shape);
  return spec === null ? null : spec.kind;
};

/**
 * Returns the categories axis labels of a chart shape, or `null`
 * if the shape isn't a chart wrapper or its kind isn't modeled.
 * Convenience over `getShapeChartSpec(shape)?.categories ?? null`.
 */
export const getShapeChartCategories = (shape: SlideShapeData): ReadonlyArray<string> | null => {
  const spec = getShapeChartSpec(shape);
  return spec === null ? null : spec.categories;
};

/**
 * Returns the chart's series-name list (in spec order). `null`
 * when the shape isn't a chart wrapper or the kind isn't modeled.
 */
export const getShapeChartSeriesNames = (shape: SlideShapeData): ReadonlyArray<string> | null => {
  const spec = getShapeChartSpec(shape);
  return spec === null ? null : spec.series.map((s) => s.name);
};

/**
 * Returns the values for the named series on a chart shape, or
 * `null` when the shape isn't a chart, the kind isn't modeled, or
 * no series matches `seriesName`.
 */
export const getShapeChartSeriesValues = (
  shape: SlideShapeData,
  seriesName: string,
): ReadonlyArray<number | null> | null => {
  const spec = getShapeChartSpec(shape);
  if (spec === null) return null;
  const series = spec.series.find((s) => s.name === seriesName);
  return series ? series.values : null;
};

export const getShapeChartSpec = (shape: SlideShapeData): ChartSpec | null => {
  const slide = shape[SHAPE_SLIDE];
  const resolved = resolveChartPartName(slide, shape);
  if (!resolved) return null;
  const part = slide[INTERNAL_PACKAGE].getPart(resolved.partName);
  if (!part) return null;
  try {
    const root = parseXml(decode(part.data)).root;
    return readChartSpec(root);
  } catch {
    return null;
  }
};

/**
 * Returns every chart on the slide that carries a series whose name
 * equals `seriesName` exactly. Useful for "find the revenue chart"
 * patterns where chart kind alone isn't unique. Skips charts whose
 * kind isn't modeled.
 */
export const findChartsBySeriesName = (
  slide: SlideData,
  seriesName: string,
): ReadonlyArray<SlideChartData> => {
  const out: SlideChartData[] = [];
  for (const chart of getSlideCharts(slide)) {
    if (chart.spec === null) continue;
    if (chart.spec.series.some((s) => s.name === seriesName)) out.push(chart);
  }
  return out;
};

/**
 * Returns the first chart on the slide whose parsed `kind` matches
 * `kind` (e.g. `'bar'`, `'line'`, `'pie'`). Returns `null` when no
 * chart on the slide has that kind, or when every chart on the slide
 * uses a kind this version doesn't yet model.
 */
export const findChartByKind = (slide: SlideData, kind: ChartKind): SlideChartData | null => {
  for (const chart of getSlideCharts(slide)) {
    if (chart.spec !== null && chart.spec.kind === kind) return chart;
  }
  return null;
};

export const getSlideCharts = (slide: SlideData): ReadonlyArray<SlideChartData> => {
  const pkg = slide[INTERNAL_PACKAGE];
  const out: SlideChartData[] = [];
  for (const shape of slide[SLIDE_SHAPES]) {
    const resolved = resolveChartPartName(slide, shape);
    if (!resolved) continue;
    const chartPart = pkg.getPart(resolved.partName);
    if (!chartPart) continue;
    let spec: ChartSpec | null;
    try {
      const root = parseXml(decode(chartPart.data)).root;
      spec = readChartSpec(root);
    } catch {
      spec = null;
    }
    out.push({ shape, spec });
  }
  return out;
};

void textNode;

// ---------------------------------------------------------------------------
// Validator.

export type { IssueSeverity, ValidationIssue };

/**
 * Runs a set of lightweight invariant checks on the package and
 * returns the list of issues found. An empty array means the deck
 * passes every check.
 *
 * Catches the common authoring mistakes — missing presentation.xml,
 * dangling slide rels, slides without a layout, etc. — without
 * depending on a heavyweight XSD engine, so it runs identically in
 * Node and the browser.
 *
 * Use it as a pre-save sanity check, especially after orchestrating
 * lots of mutations against the same package. Higher-fidelity XSD
 * validation lives in the test harness (Layer 1) and stays Node-only.
 */
export const validatePresentation = (pres: PresentationData): ReadonlyArray<ValidationIssue> =>
  validatePresentationPackage(pres[INTERNAL_PACKAGE]);

// ---------------------------------------------------------------------------
// Package introspection escape hatches.
//
// `_internalPackageOf` is the heavy escape hatch for hot-path power
// users; these two helpers cover the 80% case (just enumerate parts
// or read a single part's bytes) without exposing the OpcPackage
// class.

/**
 * Power-user escape hatch. Returns the underlying `OpcPackage`
 * backing `pres`. Use this when you need to manipulate parts /
 * rels directly. Most callers should use the typed helpers above
 * (`listPackageParts`, `readPackagePart`, `getMediaParts`, etc.).
 *
 * @internal — used by `pptx-kit/node` to mount fs-backed helpers.
 */
export const _internalPackageOf = (pres: PresentationData): OpcPackage => pres[INTERNAL_PACKAGE];

/** One entry in the package's parts list. */
export interface PackagePartInfo {
  readonly name: string;
  readonly contentType: string;
  readonly byteLength: number;
}

/**
 * Enumerates every OPC part in the package. Useful for advanced
 * inspection (e.g. "what parts does this template carry?") without
 * dropping to `_internalPackageOf`.
 */
export const listPackageParts = (pres: PresentationData): ReadonlyArray<PackagePartInfo> =>
  pres[INTERNAL_PACKAGE].parts.map((p) => ({
    name: p.name,
    contentType: p.contentType,
    byteLength: p.data.byteLength,
  }));

/**
 * Reads a single OPC part's bytes by part name (e.g.
 * `'/ppt/slides/slide1.xml'`). Returns `null` when no such part
 * exists. The returned `Uint8Array` is a live view into the
 * package — DO NOT mutate it. Use this for read-only inspection
 * (e.g. parsing custom extension parts).
 */
export const readPackagePart = (pres: PresentationData, name: string): Uint8Array | null => {
  const part = pres[INTERNAL_PACKAGE].parts.find((p) => p.name === name);
  return part?.data ?? null;
};

/** A media (image / video / audio) part embedded in the package. */
export interface MediaPart {
  readonly name: string;
  readonly contentType: string;
  readonly data: Uint8Array;
}

/**
 * Returns the total size of the package's parts in bytes
 * (uncompressed). Useful for storage estimation, quota checks,
 * and "how big is this deck before save?" diagnostics. The
 * actual `savePresentation` output is typically smaller after
 * DEFLATE; this is an upper bound.
 */
export const getPackageSize = (pres: PresentationData): number => {
  let total = 0;
  for (const part of pres[INTERNAL_PACKAGE].parts) total += part.data.byteLength;
  return total;
};

/**
 * Returns every `/ppt/media/...` part in the package. Useful for
 * audit / export workflows — e.g. "extract every embedded image."
 */
export const getMediaParts = (pres: PresentationData): ReadonlyArray<MediaPart> => {
  const out: MediaPart[] = [];
  for (const p of pres[INTERNAL_PACKAGE].parts) {
    if (p.name.startsWith('/ppt/media/')) {
      out.push({ name: p.name, contentType: p.contentType, data: p.data });
    }
  }
  return out;
};

/**
 * Returns every media part NOT referenced by any rels in the
 * package — the set `compactPackage` would remove. Non-destructive;
 * the caller decides whether to delete.
 *
 * Useful for audit UIs that want to surface bloat before cleaning,
 * and for "is this asset still used?" checks.
 */
export const getOrphanMediaPartNames = (pres: PresentationData): ReadonlyArray<string> => {
  const pkg = pres[INTERNAL_PACKAGE];
  const referenced = new Set<string>();
  const resolve = (sourcePart: string, target: string): string => {
    if (target.startsWith('/')) return target;
    const dir = sourcePart.split('/').slice(0, -1);
    const segments: string[] = [];
    for (const seg of [...dir, ...target.split('/')]) {
      if (seg === '..') segments.pop();
      else if (seg !== '.' && seg.length > 0) segments.push(seg);
    }
    return `/${segments.join('/')}`;
  };
  for (const part of pkg.parts) {
    if (!part.name.endsWith('.rels')) continue;
    // /ppt/slides/_rels/slide1.xml.rels → /ppt/slides/slide1.xml
    const m = part.name.match(/^(.*)\/_rels\/(.+)\.rels$/);
    let sourceName: string;
    if (part.name === '/_rels/.rels') {
      sourceName = '/';
    } else if (m) {
      sourceName = `${m[1]}/${m[2]}`;
    } else {
      continue;
    }
    const sourceRels = sourceName === '/' ? pkg.rootRels() : pkg.getRels(sourceName as never);
    if (!sourceRels) continue;
    for (const rel of sourceRels.items) {
      if (rel.targetMode === 'External') continue;
      referenced.add(resolve(sourceName, rel.target));
    }
  }
  const out: string[] = [];
  for (const part of pkg.parts) {
    if (!part.name.startsWith('/ppt/media/')) continue;
    if (!referenced.has(part.name)) out.push(part.name);
  }
  return out;
};

/**
 * Returns every media part name the slide's rels reference
 * (typically `/ppt/media/imageN.ext`). Walks the slide's rels
 * graph and resolves each internal target. Useful for "which
 * media files does this slide depend on?" audits.
 */
export const getSlideMediaPartNames = (slide: SlideData): ReadonlyArray<string> => {
  const pkg = slide[INTERNAL_PACKAGE];
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return [];
  const resolve = (sourcePart: string, target: string): string => {
    if (target.startsWith('/')) return target;
    const dir = sourcePart.split('/').slice(0, -1);
    const segments: string[] = [];
    for (const seg of [...dir, ...target.split('/')]) {
      if (seg === '..') segments.pop();
      else if (seg !== '.' && seg.length > 0) segments.push(seg);
    }
    return `/${segments.join('/')}`;
  };
  const out: string[] = [];
  const seen = new Set<string>();
  for (const rel of rels.items) {
    if (rel.targetMode === 'External') continue;
    const resolved = resolve(slide[SLIDE_PART_NAME], rel.target);
    if (!resolved.startsWith('/ppt/media/')) continue;
    if (seen.has(resolved)) continue;
    seen.add(resolved);
    out.push(resolved);
  }
  return out;
};

/**
 * Returns every slide that references the given media part name
 * (typically `/ppt/media/imageN.ext`). Walks each slide's rels and
 * checks whether any internal rel resolves to `mediaPartName`.
 *
 * Useful for image-audit workflows: "before I replace this image,
 * which slides will the change affect?"
 */
export const slidesUsingMediaPart = (
  pres: PresentationData,
  mediaPartName: string,
): ReadonlyArray<SlideData> => {
  const pkg = pres[INTERNAL_PACKAGE];
  const resolve = (sourcePart: string, target: string): string => {
    if (target.startsWith('/')) return target;
    const dir = sourcePart.split('/').slice(0, -1);
    const segments: string[] = [];
    for (const seg of [...dir, ...target.split('/')]) {
      if (seg === '..') segments.pop();
      else if (seg !== '.' && seg.length > 0) segments.push(seg);
    }
    return `/${segments.join('/')}`;
  };
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
    if (!rels) continue;
    const hit = rels.items.some(
      (r) =>
        r.targetMode !== 'External' && resolve(slide[SLIDE_PART_NAME], r.target) === mediaPartName,
    );
    if (hit) out.push(slide);
  }
  return out;
};

/**
 * Removes media parts that no rels graph references. Returns the
 * list of removed part names. Useful after a sequence of slide
 * removals leaves orphan images behind.
 *
 * Only `/ppt/media/...` parts are considered. The check walks every
 * `.rels` part in the package and resolves each internal rel target
 * against its source part name to build the live media set.
 */
export const compactPackage = (
  pres: PresentationData,
): { readonly removed: ReadonlyArray<string> } => {
  const pkg = pres[INTERNAL_PACKAGE];
  const referenced = new Set<string>();

  const resolve = (sourcePart: string, target: string): string => {
    if (target.startsWith('/')) return target;
    const dir = sourcePart.split('/').slice(0, -1);
    const segments: string[] = [];
    for (const seg of [...dir, ...target.split('/')]) {
      if (seg === '..') segments.pop();
      else if (seg !== '.' && seg.length > 0) segments.push(seg);
    }
    return `/${segments.join('/')}`;
  };

  for (const part of pkg.parts) {
    if (!part.name.endsWith('.rels')) continue;
    // /ppt/slides/_rels/slide1.xml.rels → /ppt/slides/slide1.xml
    // /_rels/.rels                       → / (package root)
    let sourceName = part.name.replace('/_rels/', '/').replace(/\.rels$/, '');
    if (sourceName === '/' || sourceName === '') {
      // Root rels — `rel.target` is relative to the package root.
      // We don't need to consult pkg.getRels for it (the only thing it
      // points at that we care about is the presentation.xml, which
      // has its own rels we'll walk). Just parse the part data
      // directly for completeness.
      sourceName = '/';
    }
    const rels = sourceName === '/' ? null : pkg.getRels(partName(sourceName));
    if (!rels) continue;
    for (const rel of rels.items) {
      if (rel.targetMode === 'External') continue;
      referenced.add(resolve(sourceName, rel.target));
    }
  }

  const removed: string[] = [];
  const orphans: string[] = [];
  for (const part of pkg.parts) {
    if (!part.name.startsWith('/ppt/media/')) continue;
    if (!referenced.has(part.name)) orphans.push(part.name);
  }
  for (const name of orphans) {
    pkg.removePart(partName(name));
    removed.push(name);
  }
  return { removed };
};

/**
 * Replaces the bytes of a media part in place. Returns `true` if the
 * part was found and updated, `false` otherwise. The content type is
 * preserved.
 *
 * Useful for the "swap every instance of this logo" workflow — pick
 * the right `partName` via `getMediaParts` and call this once. Every
 * `<a:blip r:embed="…"/>` reference is unaffected because the rels
 * already point at this part name.
 */
export const setMediaPartBytes = (
  pres: PresentationData,
  partName: string,
  bytes: Uint8Array,
): boolean => {
  const part = pres[INTERNAL_PACKAGE].parts.find((p) => p.name === partName);
  if (!part) return false;
  part.data = bytes;
  return true;
};

/**
 * High-level snapshot of the presentation's structure. Useful as a
 * diagnostic checklist when debugging a template or generating audit
 * reports. The numbers reflect what's reachable through the typed
 * API on the current package state.
 */
export interface PresentationSummary {
  readonly slideCount: number;
  readonly hiddenSlideCount: number;
  readonly totalShapes: number;
  readonly shapesByKind: Readonly<Record<ShapeKind, number>>;
  readonly layoutCount: number;
  readonly sectionCount: number;
  readonly partCount: number;
  readonly hasCharts: boolean;
  readonly hasComments: boolean;
  readonly hasAnimations: boolean;
  readonly themeName: string | null;
}

const CHART_CONTENT_TYPE_FN = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const COMMENTS_CONTENT_TYPE_FN =
  'application/vnd.openxmlformats-officedocument.presentationml.comments+xml';

export const getPresentationSummary = (pres: PresentationData): PresentationSummary => {
  const pkg = pres[INTERNAL_PACKAGE];
  const slides = getSlides(pres);
  let hiddenSlideCount = 0;
  let totalShapes = 0;
  const shapesByKind: Record<ShapeKind, number> = {
    shape: 0,
    picture: 0,
    group: 0,
    graphicFrame: 0,
    connector: 0,
  };
  let hasAnimations = false;
  for (const slide of slides) {
    if (isSlideHidden(slide)) hiddenSlideCount++;
    for (const s of slide[SLIDE_SHAPES]) {
      totalShapes++;
      shapesByKind[s[SHAPE_SNAPSHOT].kind]++;
    }
    // <p:timing> presence = at least one animation.
    if (
      !hasAnimations &&
      slide[SLIDE_DOCUMENT].root.children.some(
        (c) =>
          c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'timing',
      )
    ) {
      hasAnimations = true;
    }
  }

  const hasCharts = pkg.parts.some((p) => p.contentType === CHART_CONTENT_TYPE_FN);
  const hasComments = pkg.parts.some((p) => p.contentType === COMMENTS_CONTENT_TYPE_FN);
  const theme = getPresentationTheme(pres);

  return {
    slideCount: slides.length,
    hiddenSlideCount,
    totalShapes,
    shapesByKind,
    layoutCount: getSlideLayouts(pres).length,
    sectionCount: getSlideSections(pres).length,
    partCount: pkg.parts.length,
    hasCharts,
    hasComments,
    hasAnimations,
    themeName: theme?.name ?? null,
  };
};

// ---------------------------------------------------------------------------
// Picture opacity — `<a:alphaModFix>` inside the picture's `<a:blip>`.
//
// `amt` is ECMA-376's ST_PositiveFixedPercentage (0–100000, scale 1/1000
// of a percent). PowerPoint defaults to fully opaque when the element
// is absent. Pass `null` to remove a prior `<a:alphaModFix>`.

const NAME_ALPHA_MOD_FIX_FN = qname('a', 'alphaModFix', NS.dml);
const ATTR_AMT_FN = qname('', 'amt', '');

/**
 * Sets the picture's opacity (0–1 fraction; `1` is fully opaque, `0`
 * fully transparent). Pass `null` to remove an existing opacity
 * override and restore PowerPoint's default behavior.
 *
 * Throws for non-picture shapes and on opacities outside `[0, 1]`.
 */
/**
 * Returns the embedded image bytes for a picture shape, or `null`
 * when the shape isn't a picture or has no `r:embed` reference
 * (external images aren't followed).
 *
 * The returned `Uint8Array` is a live view into the package media
 * part — treat it as read-only; copy if you need an independent
 * buffer.
 */
export const getShapeImageBytes = (shape: SlideShapeData): Uint8Array | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') return null;
  const rEmbed = getPictureEmbedRId(shape[SHAPE_ELEMENT]);
  if (rEmbed === null) return null;
  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return null;
  const rel = rels.items.find((r) => r.id === rEmbed);
  if (!rel || rel.targetMode === 'External') return null;
  const mediaName = rel.target.startsWith('/')
    ? partName(rel.target)
    : resolveTarget(slide[SLIDE_PART_NAME], rel.target);
  const part = pkg.getPart(mediaName);
  return part?.data ?? null;
};

/**
 * `true` when the shape's text body carries any visible characters.
 * Tighter than checking `getShapeText(shape) !== ''` because it
 * doesn't allocate the concatenated string.
 */
export const hasShapeText = (shape: SlideShapeData): boolean => {
  const text = shape[SHAPE_SNAPSHOT].text;
  return typeof text === 'string' && text.length > 0;
};

/**
 * `true` when the shape carries an embedded image — either a
 * `<p:pic>` picture or a `<p:spPr>/<a:blipFill>` image fill on a
 * regular shape. External `r:link` references count too.
 */
export const hasShapeImage = (shape: SlideShapeData): boolean => {
  if (shape[SHAPE_SNAPSHOT].kind === 'picture') {
    return getPictureEmbedRId(shape[SHAPE_ELEMENT]) !== null;
  }
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return false;
  return firstChildElement(spPr, qname('a', 'blipFill', NS.dml)) !== null;
};

/**
 * Returns every shape on the slide that is mirrored — horizontally
 * (`flipH`), vertically (`flipV`), or both.
 */
export const findFlippedShapes = (slide: SlideData): ReadonlyArray<SlideShapeData> =>
  slide[SLIDE_SHAPES].filter((s) => {
    const flip = getShapeFlip(s);
    return flip !== null && (flip.horizontal || flip.vertical);
  });

/**
 * Returns every unordered pair of shapes on the slide whose
 * bounding boxes overlap. Built on `shapesOverlap`. Pairs are
 * returned with `a` strictly preceding `b` in document order, and
 * each pair appears at most once.
 *
 * Useful for layout audits — "do any boxes collide on this slide?"
 * Shapes without `<a:xfrm>` bounds never overlap anything.
 */
export const findOverlappingShapePairs = (
  slide: SlideData,
): ReadonlyArray<readonly [SlideShapeData, SlideShapeData]> => {
  const shapes = slide[SLIDE_SHAPES];
  const out: (readonly [SlideShapeData, SlideShapeData])[] = [];
  for (let i = 0; i < shapes.length; i++) {
    for (let j = i + 1; j < shapes.length; j++) {
      if (shapesOverlap(shapes[i]!, shapes[j]!)) {
        out.push([shapes[i]!, shapes[j]!] as const);
      }
    }
  }
  return out;
};

/**
 * Returns every shape on the slide whose bounding box extends past
 * the slide canvas (`getSlideSize(pres)`). Useful audit helper for
 * catching shapes that PowerPoint will silently render off-screen
 * or clip on export. Shapes without `<a:xfrm>` bounds are skipped.
 *
 * If the presentation has no slide-size declared, every positioned
 * shape is returned (caller can't audit against an absent canvas).
 */
export const findShapesOutsideCanvas = (
  slide: SlideData,
  pres: PresentationData,
): ReadonlyArray<SlideShapeData> => {
  const size = getSlideSize(pres);
  const out: SlideShapeData[] = [];
  for (const shape of slide[SLIDE_SHAPES]) {
    const b = getShapeBounds(shape);
    if (b === null) continue;
    if (size === null) {
      out.push(shape);
      continue;
    }
    if (b.x < 0 || b.y < 0 || b.x + b.w > size.width || b.y + b.h > size.height) {
      out.push(shape);
    }
  }
  return out;
};

/**
 * Every slide whose layout's `cSld@name` matches the given string.
 * Useful for batch operations on slides sharing a layout — for
 * example, restyling every "Title and Content" slide in a deck.
 *
 * Matching is exact (case-sensitive). Slides without a resolved
 * layout are skipped.
 */
export const findSlidesByLayoutName = (
  pres: PresentationData,
  layoutName: string,
): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    const layout = getSlideLayout(slide);
    if (layout !== null && getSlideLayoutName(layout) === layoutName) out.push(slide);
  }
  return out;
};

/**
 * Every slide whose layout `@type` (e.g. `'title'`, `'blank'`,
 * `'obj'`) matches. Sibling of `findSlidesByLayoutName`, but keyed
 * on the OOXML layout-type enum rather than the human-facing name —
 * stable across locales and template providers.
 */
export const findSlidesByLayoutType = (
  pres: PresentationData,
  layoutType: SlideLayoutType | string,
): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    const layout = getSlideLayout(slide);
    if (layout !== null && getSlideLayoutType(layout) === layoutType) out.push(slide);
  }
  return out;
};

/**
 * Returns the package part name (`/ppt/media/imageN.ext`) of
 * whichever image the shape carries — picture (`<p:pic>`) or
 * image-fill (`<a:blipFill>` nested under `<p:spPr>`). Returns
 * `null` when the shape has no embedded image, or the rel points
 * at an external `r:link` rather than an internal target.
 *
 * Useful for addressing the media part directly with
 * `setMediaPartBytes` or `readPackagePart`.
 */
/**
 * Returns the external URL of the picture when its `<a:blip>` carries an
 * `r:link` (external) relationship rather than an `r:embed`. Returns
 * `null` for embedded pictures, non-picture shapes, or when the
 * relationship doesn't resolve.
 *
 * PowerPoint emits `r:link` when the user inserts via "Link to file"
 * instead of "Insert Picture". The bytes live outside the package, so
 * `getShapeImageBytes` can't render them — readers / preview tools
 * should fall back to this URL or a placeholder.
 */
export const getShapeImageLinkUrl = (shape: SlideShapeData): string | null => {
  const slide = shape[SHAPE_SLIDE];
  const rels = slide[INTERNAL_PACKAGE].getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return null;
  // Find the blip element on either picture or shape-with-image-fill.
  let blip: XmlElement | null = null;
  if (shape[SHAPE_SNAPSHOT].kind === 'picture') {
    const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
    if (blipFill) blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  } else {
    const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
    if (spPr) {
      const bf = firstChildElement(spPr, qname('a', 'blipFill', NS.dml));
      if (bf) blip = firstChildElement(bf, qname('a', 'blip', NS.dml));
    }
  }
  if (!blip) return null;
  const rLink = getAttrValue(blip, qname('r', 'link', NS.officeDocRels));
  if (!rLink) return null;
  const rel = rels.items.find((r) => r.id === rLink);
  if (!rel || rel.targetMode !== 'External') return null;
  return rel.target;
};

export const getShapeImagePartName = (shape: SlideShapeData): string | null => {
  const slide = shape[SHAPE_SLIDE];
  const rels = slide[INTERNAL_PACKAGE].getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return null;

  const resolve = (rEmbed: string | null): string | null => {
    if (rEmbed === null) return null;
    const rel = rels.items.find((r) => r.id === rEmbed);
    if (!rel || rel.targetMode === 'External') return null;
    const name = rel.target.startsWith('/')
      ? partName(rel.target)
      : resolveTarget(slide[SLIDE_PART_NAME], rel.target);
    return name;
  };

  // Picture shape: <p:pic><p:blipFill><a:blip r:embed="..."/>.
  if (shape[SHAPE_SNAPSHOT].kind === 'picture') {
    const rEmbed = getPictureEmbedRId(shape[SHAPE_ELEMENT]);
    return resolve(rEmbed);
  }

  // Other shapes with image fill: <p:spPr><a:blipFill><a:blip r:embed="..."/>.
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const blipFill = firstChildElement(spPr, qname('a', 'blipFill', NS.dml));
  if (!blipFill) return null;
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) return null;
  return resolve(getAttrValue(blip, qname('r', 'embed', NS.officeDocRels)));
};

/**
 * Returns the bytes of the image used as this shape's *fill*
 * (`<a:blipFill>` nested under `<p:spPr>`, as written by
 * `setShapeImageFill`). Distinct from `getShapeImageBytes`, which only
 * applies to `<p:pic>` picture shapes.
 *
 * Returns null if the shape has no image fill, the blip has no
 * `r:embed`, or the embed points at an external `r:link`.
 */
export const getShapeImageFillBytes = (shape: SlideShapeData): Uint8Array | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const blipFill = firstChildElement(spPr, qname('a', 'blipFill', NS.dml));
  if (!blipFill) return null;
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) return null;
  const rEmbed = getAttrValue(blip, qname('r', 'embed', NS.officeDocRels));
  if (rEmbed === null) return null;
  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return null;
  const rel = rels.items.find((r) => r.id === rEmbed);
  if (!rel || rel.targetMode === 'External') return null;
  const mediaName = rel.target.startsWith('/')
    ? partName(rel.target)
    : resolveTarget(slide[SLIDE_PART_NAME], rel.target);
  const part = pkg.getPart(mediaName);
  return part?.data ?? null;
};

/**
 * Returns the image format token (`'png'`, `'jpeg'`, …) for whichever
 * image bytes the shape carries — picture (`<p:pic>`) or image-fill
 * (`<a:blipFill>` on `<p:spPr>`). Returns `null` if the shape has no
 * embedded image or the bytes don't match a recognized signature.
 */
export const getShapeImageFormat = (shape: SlideShapeData): ImageFormat | null => {
  const bytes = getShapeImageBytes(shape) ?? getShapeImageFillBytes(shape);
  if (bytes === null) return null;
  return detectImageFormat(bytes);
};

/**
 * Reads the picture's opacity (0–1 fraction). Returns `null` when no
 * `<a:alphaModFix>` is present (PowerPoint treats absence as fully
 * opaque); returns `1` when an explicit alphaModFix sets full opacity.
 */
export const getShapeImageOpacity = (shape: SlideShapeData): number | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') return null;
  const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
  if (!blipFill) return null;
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) return null;
  const alpha = firstChildElement(blip, qname('a', 'alphaModFix', NS.dml));
  if (!alpha) return null;
  const amt = getAttrValue(alpha, qname('', 'amt', ''));
  if (amt === null) return 1;
  const n = Number.parseInt(amt, 10);
  if (!Number.isFinite(n)) return null;
  return n / 100000;
};

/**
 * Reads the picture's crop fractions. Returns `null` when no
 * `<a:srcRect>` is present; otherwise returns a fully-populated object
 * with every side filled in (0 for omitted sides on disk).
 */
export const getShapeImageCrop = (shape: SlideShapeData): ImageCrop | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') return null;
  const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
  if (!blipFill) return null;
  const srcRect = firstChildElement(blipFill, qname('a', 'srcRect', NS.dml));
  if (!srcRect) return null;
  const parseSide = (local: string): number => {
    const v = getAttrValue(srcRect, qname('', local, ''));
    if (v === null) return 0;
    const n = Number.parseInt(v, 10);
    return Number.isFinite(n) ? n / 100000 : 0;
  };
  return {
    left: parseSide('l'),
    top: parseSide('t'),
    right: parseSide('r'),
    bottom: parseSide('b'),
  };
};

/**
 * Adjusts the picture's brightness by writing `<a:lumOff val="…"/>`
 * inside `<a:blip>`. The value is a -1..1 fraction:
 *
 *   - `1`     → +100% brightness
 *   - `0` or `null` → no offset (any prior `<a:lumOff>` is removed)
 *   - `-1`    → -100% brightness
 *
 * Throws for non-picture shapes and on values outside [-1, 1].
 *
 * Note: PowerPoint's "Picture Format › Corrections" UI couples this
 * with `<a:lumMod>` for some presets; this primitive sets only
 * `lumOff` to keep the surface honest. Read it back via
 * `getShapeImageBrightness`.
 */
export const setShapeImageBrightness = (shape: SlideShapeData, value: number | null): void => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') {
    throw new Error(
      `setShapeImageBrightness only works on picture shapes; ${shape[SHAPE_SNAPSHOT].kind} is not one`,
    );
  }
  const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
  if (!blipFill) throw new Error('picture has no <p:blipFill>');
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) throw new Error('picture <p:blipFill> has no <a:blip>');
  blip.children = blip.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'lumOff'),
  );
  if (value !== null && value !== 0) {
    if (!Number.isFinite(value) || value < -1 || value > 1) {
      throw new RangeError(`brightness must be in [-1, 1], got ${value}`);
    }
    blip.children.push(
      elem(qname('a', 'lumOff', NS.dml), {
        attrs: [attr(qname('', 'val', ''), String(Math.round(value * 100000)))],
      }),
    );
  }
  commitAndRefresh(shape);
};

/**
 * Adjusts the picture's contrast by writing `<a:lumMod val="…"/>` on
 * `<a:blip>`. The value is a 0..2 fraction:
 *
 *   - `1` or `null` → no modulation (any prior `<a:lumMod>` is removed)
 *   - `0.5`         → 50% of original luminance variance (washed out)
 *   - `1.5`         → 150% (boosted contrast; PowerPoint clamps to
 *                       what the renderer supports)
 *
 * Throws on non-picture shapes and on values outside `[0, 2]`. The
 * primitive maps directly to `ST_PositiveFixedPercentage` × 100000.
 */
export const setShapeImageContrast = (shape: SlideShapeData, value: number | null): void => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') {
    throw new Error(
      `setShapeImageContrast only works on picture shapes; ${shape[SHAPE_SNAPSHOT].kind} is not one`,
    );
  }
  const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
  if (!blipFill) throw new Error('picture has no <p:blipFill>');
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) throw new Error('picture <p:blipFill> has no <a:blip>');
  blip.children = blip.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'lumMod'),
  );
  if (value !== null && value !== 1) {
    if (!Number.isFinite(value) || value < 0 || value > 2) {
      throw new RangeError(`contrast must be in [0, 2], got ${value}`);
    }
    blip.children.push(
      elem(qname('a', 'lumMod', NS.dml), {
        attrs: [attr(qname('', 'val', ''), String(Math.round(value * 100000)))],
      }),
    );
  }
  commitAndRefresh(shape);
};

/**
 * Reads the picture's contrast modulation (the `<a:lumMod>` fraction
 * in [0, 2]). Returns `null` when no `<a:lumMod>` is present.
 */
export const getShapeImageContrast = (shape: SlideShapeData): number | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') return null;
  const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
  if (!blipFill) return null;
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) return null;
  const lumMod = firstChildElement(blip, qname('a', 'lumMod', NS.dml));
  if (!lumMod) return null;
  const v = getAttrValue(lumMod, qname('', 'val', ''));
  if (v === null) return null;
  const n = Number.parseInt(v, 10);
  return Number.isFinite(n) ? n / 100000 : null;
};

/**
 * Reads the picture's brightness offset (the `<a:lumOff>` fraction
 * in [-1, 1]). Returns `null` when no `<a:lumOff>` is present.
 */
export const getShapeImageBrightness = (shape: SlideShapeData): number | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') return null;
  const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
  if (!blipFill) return null;
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) return null;
  const lumOff = firstChildElement(blip, qname('a', 'lumOff', NS.dml));
  if (!lumOff) return null;
  const v = getAttrValue(lumOff, qname('', 'val', ''));
  if (v === null) return null;
  const n = Number.parseInt(v, 10);
  return Number.isFinite(n) ? n / 100000 : null;
};

export const setShapeImageOpacity = (shape: SlideShapeData, opacity: number | null): void => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') {
    throw new Error(
      `setShapeImageOpacity only works on picture shapes; ${shape[SHAPE_SNAPSHOT].kind} is not one`,
    );
  }
  const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
  if (!blipFill) throw new Error('picture has no <p:blipFill>');
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) throw new Error('picture <p:blipFill> has no <a:blip>');

  blip.children = blip.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        c.name.localName === 'alphaModFix'
      ),
  );

  if (opacity !== null) {
    if (!Number.isFinite(opacity) || opacity < 0 || opacity > 1) {
      throw new RangeError(`opacity must be in [0, 1], got ${opacity}`);
    }
    blip.children.push(
      elem(NAME_ALPHA_MOD_FIX_FN, {
        attrs: [attr(ATTR_AMT_FN, String(Math.round(opacity * 100000)))],
      }),
    );
  }
  commitAndRefresh(shape);
};

// ---------------------------------------------------------------------------
// Picture cropping — `<a:srcRect>` inside the picture's `<p:blipFill>`.
//
// Percentages are 0-1 fractions per side, converted to ECMA-376's
// `ST_Percentage` units (1/1000 of a percent, so 0.25 → "25000"). Pass
// `null` to remove an existing crop.

/** Crop a picture by fraction of each side. Omitted sides default to 0. */
export interface ImageCrop {
  readonly left?: number;
  readonly top?: number;
  readonly right?: number;
  readonly bottom?: number;
}

const NAME_BLIP_FILL_FN = qname('p', 'blipFill', NS.pml);
const NAME_SRC_RECT_FN = qname('a', 'srcRect', NS.dml);
const NAME_BLIP_FN = qname('a', 'blip', NS.dml);
const ATTR_CROP_L = qname('', 'l', '');
const ATTR_CROP_T = qname('', 't', '');
const ATTR_CROP_R = qname('', 'r', '');
const ATTR_CROP_B = qname('', 'b', '');

const fractionToST = (n: number | undefined): string | null => {
  if (n === undefined || n === 0) return null;
  if (!Number.isFinite(n) || n < 0 || n >= 1) {
    throw new RangeError(`crop fraction must be in [0, 1), got ${n}`);
  }
  return String(Math.round(n * 100000));
};

/**
 * Sets (or clears) a `<a:srcRect>` on a picture shape, cropping the
 * embedded image by the given fraction on each side. Pass `null` to
 * remove an existing crop.
 *
 * Fractions are in `[0, 1)` per side. `{ left: 0.25 }` clips 25% off
 * the left edge; the visible image stretches to fill the original
 * frame. The shape's geometry (`<a:xfrm>`) is unchanged.
 */
export const setShapeImageCrop = (shape: SlideShapeData, crop: ImageCrop | null): void => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') {
    throw new Error(
      `setShapeImageCrop only works on picture shapes; ${shape[SHAPE_SNAPSHOT].kind} is not one`,
    );
  }
  const pic = shape[SHAPE_ELEMENT];
  const blipFill = firstChildElement(pic, NAME_BLIP_FILL_FN);
  if (!blipFill) throw new Error('picture has no <p:blipFill>');

  // Remove any existing srcRect first.
  blipFill.children = blipFill.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'srcRect'),
  );

  if (crop === null) {
    commitAndRefresh(shape);
    return;
  }

  const attrs: Array<ReturnType<typeof attr>> = [];
  const l = fractionToST(crop.left);
  const t = fractionToST(crop.top);
  const r = fractionToST(crop.right);
  const b = fractionToST(crop.bottom);
  if (l !== null) attrs.push(attr(ATTR_CROP_L, l));
  if (t !== null) attrs.push(attr(ATTR_CROP_T, t));
  if (r !== null) attrs.push(attr(ATTR_CROP_R, r));
  if (b !== null) attrs.push(attr(ATTR_CROP_B, b));

  // <a:srcRect> sits between <a:blip> and <a:stretch> per the schema.
  const srcRect = elem(NAME_SRC_RECT_FN, { attrs });
  const blipIdx = blipFill.children.findIndex(
    (c) => c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'blip',
  );
  if (blipIdx === -1) {
    // No <a:blip>? Just prepend the srcRect.
    blipFill.children.unshift(srcRect);
  } else {
    blipFill.children.splice(blipIdx + 1, 0, srcRect);
  }
  commitAndRefresh(shape);
};

void NAME_BLIP_FN;

// ---------------------------------------------------------------------------
// Animations (single-effect, click-triggered).
//
// v1 scope: exactly one effect per slide, click-triggered, entrance or
// exit preset family. The plan calls this the curated subset; full
// multi-effect timing-tree authoring is post-1.0.

export type { AnimationEffect, AnimationOptions };

const NAME_TIMING_FN = qname('p', 'timing', NS.pml);

const removeExistingTiming = (slide: SlideData): void => {
  slide[SLIDE_DOCUMENT].root.children = slide[SLIDE_DOCUMENT].root.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'timing'),
  );
};

const insertTimingAtEnd = (slide: SlideData, timing: XmlElement): void => {
  // Schema ordering: `<p:timing>` is one of the last children of `<p:sld>`
  // (after cSld, clrMapOvr, transition). Appending to the end of
  // `<p:sld>` keeps the file valid.
  slide[SLIDE_DOCUMENT].root.children.push(timing);
};

/**
 * Sets a single click-triggered animation effect on the given shape.
 * Replaces any existing `<p:timing>` block on the slide — v1 supports
 * exactly one effect per slide. Calling this on a second shape replaces
 * the first.
 *
 * Supported `effect` tokens:
 *
 *   - `'fadeIn'`   entrance fade
 *   - `'fadeOut'`  exit fade
 *   - `'appear'`   instant entrance
 *   - `'disappear'` instant exit
 *
 * `durationMs` defaults to 500ms (fades only — `appear`/`disappear`
 * are instantaneous).
 */
export const setShapeAnimation = (shape: SlideShapeData, opts: AnimationOptions): void => {
  const slide = shape[SHAPE_SLIDE];
  removeExistingTiming(slide);
  const spid = shape[SHAPE_SNAPSHOT].id;
  const timing = buildSingleEffectTiming(spid, opts);
  insertTimingAtEnd(slide, timing);
  commitSlideData(slide);
  refreshSlideData(slide);
};

/**
 * Returns the animation effect bound to this shape via the slide's
 * `<p:timing>` tree, or `null` if the shape has no animation in the
 * v1 single-effect schema we model. Unknown presets are reported as a
 * raw `null` rather than guessing.
 */
export const getShapeAnimation = (shape: SlideShapeData): AnimationEffect | null => {
  const slide = shape[SHAPE_SLIDE];
  const timing = slide[SLIDE_DOCUMENT].root.children.find(
    (c): c is XmlElement =>
      c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'timing',
  );
  if (!timing) return null;

  // Confirm the shape's spid appears in <p:bldLst><p:bldP spid="..."/>.
  const bldLst = firstChildElement(timing, qname('p', 'bldLst', NS.pml));
  if (!bldLst) return null;
  const spidStr = String(shape[SHAPE_SNAPSHOT].id);
  const matched = allChildElements(bldLst, qname('p', 'bldP', NS.pml)).some(
    (b) => getAttrValue(b, qname('', 'spid', '')) === spidStr,
  );
  if (!matched) return null;

  // Walk the timing tree to find the effect cTn for this shape. Our
  // builder emits `<p:cTn presetID="N" presetClass="entr|exit" ...
  // nodeType="clickEffect">` with a `<p:spTgt spid="..."/>` inside. We
  // accept any cTn carrying that combination.
  let presetID: string | null = null;
  let presetClass: string | null = null;
  const walk = (el: XmlElement): boolean => {
    if (el.name.namespaceURI === NS.pml && el.name.localName === 'cTn') {
      const cls = getAttrValue(el, qname('', 'presetClass', ''));
      const id = getAttrValue(el, qname('', 'presetID', ''));
      if (cls && id) {
        // Confirm this cTn targets our shape via a descendant spTgt.
        const targetsShape = (sub: XmlElement): boolean => {
          if (
            sub.name.namespaceURI === NS.pml &&
            sub.name.localName === 'spTgt' &&
            getAttrValue(sub, qname('', 'spid', '')) === spidStr
          ) {
            return true;
          }
          for (const c of sub.children) {
            if (c.kind === 'element' && targetsShape(c)) return true;
          }
          return false;
        };
        if (targetsShape(el)) {
          presetClass = cls;
          presetID = id;
          return true;
        }
      }
    }
    for (const c of el.children) {
      if (c.kind === 'element' && walk(c)) return true;
    }
    return false;
  };
  walk(timing);
  if (!presetID || !presetClass) return null;

  // Map back to AnimationEffect.
  const id = Number.parseInt(presetID, 10);
  if (presetClass === 'entr' && id === 1) return 'appear';
  if (presetClass === 'entr' && id === 10) return 'fadeIn';
  if (presetClass === 'exit' && id === 1) return 'disappear';
  if (presetClass === 'exit' && id === 10) return 'fadeOut';
  return null;
};

/** Removes the slide's `<p:timing>` element entirely. */
export const clearSlideAnimations = (slide: SlideData): void => {
  removeExistingTiming(slide);
  commitSlideData(slide);
  refreshSlideData(slide);
};

void NAME_TIMING_FN;

// ---------------------------------------------------------------------------
// Slide title convenience.
//
// Most decks bind their title placeholder to `type="title"` or `type="ctrTitle"`
// (the latter is the centered hero title on a "Title Slide" layout).
// These two helpers cover ~90% of the "set the slide title" use case.

/**
 * Returns the slide's title text, or `null` if neither a `title` nor
 * a `ctrTitle` placeholder is present.
 */
export const getSlideTitle = (slide: SlideData): string | null => {
  const titleShape =
    findSlidePlaceholder(slide, 'title') ?? findSlidePlaceholder(slide, 'ctrTitle');
  if (titleShape === null) return null;
  return titleShape[SHAPE_SNAPSHOT].text ?? null;
};

/**
 * Returns the slide's body text, or `null` if no `body` placeholder
 * is present. Mirror of `getSlideTitle`; pairs with `setSlideBody`.
 */
export const getSlideBody = (slide: SlideData): string | null => {
  const bodyShape = findSlidePlaceholder(slide, 'body');
  if (bodyShape === null) return null;
  return bodyShape[SHAPE_SNAPSHOT].text ?? null;
};

/**
 * Sets the slide's title text. Looks for a `title` placeholder first,
 * falling back to `ctrTitle`. Throws if neither exists — the slide's
 * layout has no title slot.
 */
export const setSlideTitle = (slide: SlideData, title: string): void => {
  const titleShape =
    findSlidePlaceholder(slide, 'title') ?? findSlidePlaceholder(slide, 'ctrTitle');
  if (titleShape === null) {
    throw new Error('setSlideTitle: slide has no title / ctrTitle placeholder');
  }
  setShapeText(titleShape, title);
};

/**
 * Bulk-fills slide placeholders by type token. Each entry in
 * `byType` maps a `<p:ph type>` token (e.g. `'title'`, `'body'`,
 * `'ftr'`, `'dt'`) to the text to set. Silently skips entries
 * whose placeholder isn't present on the slide.
 *
 * Useful for template-fill workflows where the caller has all the
 * data in one struct.
 */
export const setSlidePlaceholders = (
  slide: SlideData,
  byType: Readonly<Record<string, string>>,
): void => {
  for (const [type, text] of Object.entries(byType)) {
    const shape =
      type === 'title'
        ? (findSlidePlaceholder(slide, 'title') ?? findSlidePlaceholder(slide, 'ctrTitle'))
        : findSlidePlaceholder(slide, type);
    if (shape !== null) setShapeText(shape, text);
  }
};

/**
 * Writes `text` into the first body placeholder on the slide.
 * Newlines start a new paragraph (each becomes its own bullet on
 * layouts that bullet their body placeholder).
 *
 * Throws when the slide has no body placeholder — pair with
 * `findSlideLayoutByType(pres, 'obj')` / `'tx'` to add the slide
 * onto a layout that has one.
 */
export const setSlideBody = (slide: SlideData, text: string): void => {
  const bodyShape = findSlidePlaceholder(slide, 'body');
  if (bodyShape === null) {
    throw new Error('setSlideBody: slide has no body placeholder');
  }
  setShapeText(bodyShape, text);
};

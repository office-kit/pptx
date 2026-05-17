// Presentation-level operations: load, save, and the package-level
// metadata that lives outside of slides — sections, layouts, theme,
// fonts, core/extended properties, thumbnail.

import { readPosition, readSize } from '../../internal/drawingml/index.ts';
import type { Emu } from '../units.ts';
import {
  type ImageFormat,
  type PartName,
  contentTypeForFormat,
  detectImageFormat,
  emptyRels,
  extensionForFormat,
  nextRelId,
  partName,
} from '../../internal/opc/index.ts';
import { OpcPackage } from '../../internal/parts/index.ts';
import {
  REL_TYPES,
  type SlideLayoutType,
  readPresentationPart,
  readSlideLayoutPart,
} from '../../internal/presentationml/index.ts';
import {
  NS,
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
} from '../../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  LAYOUT_PART,
  LAYOUT_PART_NAME,
  type PresentationData,
  SLIDE_PART_NAME,
  type SlideData,
  type SlideLayoutData,
} from '../_internal-symbols.ts';
import {
  PRES_PART_NAME,
  SLIDE_LAYOUT_CONTENT_TYPE,
  decode,
  encode,
  setOpcDefault,
} from './_helpers.ts';
import type { ShapeBounds } from './shapes.ts';
import { getSlides } from './slides.ts';

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

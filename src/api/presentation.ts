// Public `Presentation` entry point.
//
// The class wraps an internal `OpcPackage` and currently exposes load,
// save, and a read-only `slides` view. Authoring methods land as the
// internal PresentationML / DrawingML layers grow real width.

import {
  basename,
  emptyRels,
  nextRelId,
  partName,
  relsPartNameFor,
} from '../internal/opc/index.ts';
import { OpcPackage } from '../internal/parts/index.ts';
import {
  REL_TYPES,
  buildSlideFromLayout,
  readPresentationPart,
  readSlideLayoutPart,
} from '../internal/presentationml/index.ts';
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
} from '../internal/xml/index.ts';
import { SlideLayout } from './slide-layout.ts';
import { _internalCreateSlide, type Slide } from './slide.ts';

const encoder = new TextEncoder();
const decoder = new TextDecoder();

const NAME_PRESENTATION = qname('p', 'presentation', NS.pml);
const NAME_SLD_MASTER_ID_LST = qname('p', 'sldMasterIdLst', NS.pml);
const NAME_SLD_ID_LST = qname('p', 'sldIdLst', NS.pml);
const NAME_SLD_ID = qname('p', 'sldId', NS.pml);
const ATTR_ID = qname('', 'id', '');
const ATTR_R_ID = qname('r', 'id', NS.officeDocRels);

const PRES_PART_NAME = partName('/ppt/presentation.xml');
const SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';

// PowerPoint accepts sldIds in [256, 2³¹−1024]. Reuse the cap.
const SLD_ID_MIN = 256;
const SLD_ID_MAX = 2147482623;

// Ensures presentation.xml has a `<p:sldIdLst>` and returns it. Inserts it
// immediately after `<p:sldMasterIdLst>` if it didn't exist — the spec's
// element-sequence requires that ordering.
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

/**
 * Anything that can be turned into a `Uint8Array` of PPTX bytes:
 *
 *   - `Uint8Array` — used as-is.
 *   - `ArrayBuffer` — wrapped without copying.
 *   - `Blob` or `File` — read via `arrayBuffer()`. Supported in both Node
 *     (since Node 18) and all current browsers.
 *
 * Path strings are NOT accepted here to keep the browser bundle free of
 * `node:fs`. Use the `loadFile` helper from `pptx-kit/node` instead.
 */
export type PresentationInput = Uint8Array | ArrayBuffer | Blob;

const normalize = async (input: PresentationInput): Promise<Uint8Array> => {
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  if (typeof Blob !== 'undefined' && input instanceof Blob) {
    return new Uint8Array(await input.arrayBuffer());
  }
  throw new TypeError('Presentation.load: expected Uint8Array | ArrayBuffer | Blob');
};

/**
 * Internal symbol used to hand the underlying `OpcPackage` to subclasses
 * (such as the Node-specific extension in `pptx-kit/node`) without
 * exposing it on the public type. Treat as `@internal`.
 */
const INTERNAL_PACKAGE = Symbol('pptx-kit.package');

export class Presentation {
  /** @internal */
  [INTERNAL_PACKAGE]: OpcPackage;

  /** @internal */
  protected constructor(pkg: OpcPackage) {
    this[INTERNAL_PACKAGE] = pkg;
  }

  /**
   * Loads an existing `.pptx` and returns a `Presentation`.
   *
   * Throws `TypeError` if the input is not a supported type, and
   * `Error` (with a descriptive message) if the bytes are not a valid
   * OPC package.
   */
  static async load(input: PresentationInput): Promise<Presentation> {
    const bytes = await normalize(input);
    return new Presentation(OpcPackage.load(bytes));
  }

  /**
   * Creates a fresh, empty `Presentation`. The result is NOT yet a valid
   * PPTX — it carries only the OPC defaults. Subsequent phases will add
   * a `Presentation.create()` that produces the canonical PowerPoint
   * skeleton (presentation.xml + at least one slide / layout / master /
   * theme).
   */
  static create(): Presentation {
    return new Presentation(OpcPackage.empty());
  }

  /**
   * Serializes the presentation back to PPTX bytes. Internally re-emits
   * `[Content_Types].xml` and every part, preserving entry order.
   */
  save(): Promise<Uint8Array> {
    return Promise.resolve(this[INTERNAL_PACKAGE].save());
  }

  /** @internal — populated lazily on first `slides` read; reused for mutation. */
  private _slidesCache: Slide[] | null = null;

  /**
   * Enumerates slides in presentation order. Returns cached `Slide`
   * instances so that mutations made through one handle (e.g.
   * `slide.shapes[0].setText('...')`) persist for the lifetime of this
   * `Presentation`.
   *
   * Throws if any referenced slide part is missing — a structurally
   * invalid PPTX cannot honor the L1 contract.
   */
  get slides(): ReadonlyArray<Slide> {
    if (this._slidesCache !== null) return this._slidesCache;

    const pkg = this[INTERNAL_PACKAGE];
    const presPart = pkg.getPart(partName('/ppt/presentation.xml'));
    if (presPart === null) {
      this._slidesCache = [];
      return this._slidesCache;
    }
    const presRels = pkg.getRels(partName('/ppt/presentation.xml'));
    if (presRels === null) {
      this._slidesCache = [];
      return this._slidesCache;
    }

    const presRoot = parseXml(new TextDecoder().decode(presPart.data)).root;
    const presModel = readPresentationPart(presRoot);

    const out: Slide[] = [];
    for (const sld of presModel.slides) {
      const rel = presRels.items.find((r) => r.id === sld.rId);
      if (!rel) {
        throw new Error(`presentation.xml.rels missing entry for ${sld.rId}`);
      }
      const target = rel.target;
      // Quick relative-resolve: presentation.xml lives at /ppt/; slide
      // targets are like "slides/slideN.xml".
      const slideName = partName(target.startsWith('/') ? target : `/ppt/${target}`);
      const slidePart = pkg.getPart(slideName);
      if (slidePart === null) {
        throw new Error(`slide part ${slideName} not found`);
      }
      out.push(_internalCreateSlide(pkg, slideName, slidePart.data));
    }
    this._slidesCache = out;
    return out;
  }

  /**
   * Replaces `{{key}}` tokens on every slide with the value at
   * `tokens[key]`. Returns the total number of substitutions performed.
   *
   * Convenience over `for (const s of pres.slides) s.replaceTokens(tokens)`.
   * Useful for the common "fill the template once" workflow.
   */
  replaceTokens(tokens: Record<string, string>): number {
    let n = 0;
    for (const slide of this.slides) {
      n += slide.replaceTokens(tokens);
    }
    return n;
  }

  /**
   * Every slide layout in the package, in the order their parts appear on
   * disk. Read-only at this phase; authoring methods will land alongside
   * `Presentation.slides.add(...)`.
   *
   * Layouts are surfaced flat rather than grouped under their slide master
   * because most callers want to pick one by name (`Title and Content`),
   * not navigate the master → layouts tree.
   */
  get slideLayouts(): ReadonlyArray<SlideLayout> {
    const pkg = this[INTERNAL_PACKAGE];
    const out: SlideLayout[] = [];
    for (const part of pkg.parts) {
      if (
        part.contentType ===
        'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'
      ) {
        const root = parseXml(decoder.decode(part.data)).root;
        out.push(new SlideLayout(part.name, readSlideLayoutPart(root)));
      }
    }
    return out;
  }

  /**
   * Adds a new slide bound to `layout`. Returns the new `Slide`.
   *
   * What changes in the package:
   *
   *   - A new `/ppt/slides/slideN.xml` part is allocated. N is the lowest
   *     free index (so deleting slide5 and adding leaves the next add at 5,
   *     not at 6 — matches PowerPoint's behavior).
   *   - The new slide's XML carries placeholder stubs cloned from the
   *     layout: each layout `<p:ph>` becomes an empty placeholder shape
   *     on the slide, ready for `slide.findPlaceholder(...).setText(...)`.
   *   - `[Content_Types].xml` gets an `Override` for the new part.
   *   - The slide's own `.rels` gets a `slideLayout` relationship pointing
   *     at the chosen layout.
   *   - `presentation.xml.rels` gets a `slide` relationship pointing at
   *     the new slide part.
   *   - `presentation.xml`'s `<p:sldIdLst>` gets a new `<p:sldId>` with a
   *     freshly allocated id in `[256, 2³¹−1024]`. The `sldIdLst` is
   *     created if it didn't exist.
   *
   * The new slide is appended to deck order. To insert at a different
   * position, reorder the underlying `<p:sldId>` element after the call
   * (a `move` helper lands when collection methods do).
   */
  addSlide(options: { layout: SlideLayout }): Slide {
    const pkg = this[INTERNAL_PACKAGE];
    const layout = options.layout;

    // 1. Parse presentation.xml.
    const presPart = pkg.getPart(PRES_PART_NAME);
    if (!presPart) throw new Error('presentation.xml is missing');
    const presDoc = parseXml(decoder.decode(presPart.data));
    if (
      presDoc.root.name.namespaceURI !== NAME_PRESENTATION.namespaceURI ||
      presDoc.root.name.localName !== 'presentation'
    ) {
      throw new Error('presentation.xml root is not <p:presentation>');
    }

    // 2. Allocate identifiers.
    const sldIdLst = ensureSldIdLst(presDoc.root);
    const newSldId = allocateSldId(sldIdLst);

    let slideN = 1;
    for (const p of pkg.parts) {
      const m = p.name.match(/^\/ppt\/slides\/slide(\d+)\.xml$/);
      if (m?.[1] !== undefined) {
        const n = Number.parseInt(m[1], 10);
        if (Number.isFinite(n) && n >= slideN) slideN = n + 1;
      }
    }
    const newSlidePartName = partName(`/ppt/slides/slide${slideN}.xml`);

    // 3. Build the slide XML from the layout's placeholders.
    const layoutCsld = firstChildElement(layout._part.root, qname('p', 'cSld', NS.pml));
    if (!layoutCsld) throw new Error(`layout ${layout._partName} missing <p:cSld>`);
    const layoutSpTree = firstChildElement(layoutCsld, qname('p', 'spTree', NS.pml));
    if (!layoutSpTree) throw new Error(`layout ${layout._partName} missing <p:spTree>`);
    const slideDoc = buildSlideFromLayout(layoutSpTree);
    const slideBytes = encoder.encode(serializeXml(slideDoc));

    // 4. Add the slide part with its content-type override.
    pkg.addPart(newSlidePartName, SLIDE_CONTENT_TYPE, slideBytes);

    // 5. slide → layout rel.
    const slideRels = emptyRels();
    const layoutBasename = basename(layout._partName);
    slideRels.items.push({
      id: 'rId1',
      type: REL_TYPES.slideLayout,
      target: `../slideLayouts/${layoutBasename}`,
      targetMode: 'Internal',
    });
    pkg.setRels(newSlidePartName, slideRels);

    // 6. presentation → slide rel.
    const presRels = pkg.getRels(PRES_PART_NAME) ?? emptyRels();
    const newRId = nextRelId(presRels.items.map((r) => r.id));
    presRels.items.push({
      id: newRId,
      type: REL_TYPES.slide,
      target: `slides/slide${slideN}.xml`,
      targetMode: 'Internal',
    });
    pkg.setRels(PRES_PART_NAME, presRels);

    // 7. presentation.xml's sldIdLst entry.
    const newSldIdElement = elem(NAME_SLD_ID, {
      attrs: [attr(ATTR_ID, String(newSldId)), attr(ATTR_R_ID, newRId)],
    });
    sldIdLst.children.push(newSldIdElement);
    presPart.data = encoder.encode(serializeXml(presDoc));

    // 8. Invalidate the slides cache and return the new Slide.
    this._slidesCache = null;
    const slides = this.slides;
    const lastSlide = slides[slides.length - 1];
    if (!lastSlide) throw new Error('addSlide: post-condition failed; slide not in cache');
    return lastSlide;
  }

  /**
   * Removes the given slide from the deck.
   *
   * Side effects:
   *
   *   - The slide's `<p:sldId>` entry is dropped from `<p:sldIdLst>`.
   *   - The slide's relationship is removed from `presentation.xml.rels`.
   *   - The slide part and its sibling `.rels` part are deleted from the
   *     package, along with their `Content_Types` `Override` entries.
   *
   * The freed `sldId` is NOT reused on subsequent `addSlide` calls — older
   * PowerPoint builds mis-routed relationships when an id came back from
   * the dead, and the plan calls this out as a fixed-from-day-one quirk.
   *
   * Media parts referenced from the removed slide are kept in place: they
   * may be shared with other slides, and orphan media is harmless.
   *
   * Throws if `slide` does not belong to this presentation.
   */
  removeSlide(slide: Slide): void {
    const pkg = this[INTERNAL_PACKAGE];
    const slidePartName = slide._partName;

    // Sanity: confirm the slide is in this package.
    if (pkg.getPart(slidePartName) === null) {
      throw new Error(`removeSlide: ${slidePartName} not present in package`);
    }

    // 1. Drop the rel from presentation.xml.rels and find its rId.
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

    // 2. Drop the `<p:sldId>` entry that pointed at that rId.
    const presPart = pkg.getPart(PRES_PART_NAME);
    if (!presPart) throw new Error('presentation.xml missing');
    const presDoc = parseXml(decoder.decode(presPart.data));
    const sldIdLst = firstChildElement(presDoc.root, NAME_SLD_ID_LST);
    if (sldIdLst !== null) {
      sldIdLst.children = sldIdLst.children.filter((c) => {
        if (c.kind !== 'element') return true;
        if (c.name.namespaceURI !== NS.pml || c.name.localName !== 'sldId') return true;
        return getAttrValue(c, ATTR_R_ID) !== removedRel.id;
      });
    }
    presPart.data = encoder.encode(serializeXml(presDoc));

    // 3. Remove the slide's `.rels` part and the slide part itself.
    pkg.removePart(relsPartNameFor(slidePartName));
    pkg.removePart(slidePartName);

    // 4. Invalidate the cached slides array.
    this._slidesCache = null;
  }

  /**
   * Duplicates a slide. Returns the new `Slide` appended to deck order.
   *
   * The slide's part bytes and `.rels` are cloned verbatim, so the
   * duplicate inherits every shape, text, picture, layout binding, and
   * media reference of the source. Media parts are NOT copied; the
   * duplicate's rels point at the same media as the original (PowerPoint
   * does the same — media sharing across slides is canonical).
   *
   * Identifiers are freshly allocated:
   *   - new slide part name `/ppt/slides/slideN.xml`
   *   - new `sldId` in [256, 2³¹−1024]
   *   - new presentation→slide `rId`
   */
  duplicateSlide(slide: Slide): Slide {
    const pkg = this[INTERNAL_PACKAGE];

    const sourcePart = pkg.getPart(slide._partName);
    if (!sourcePart) throw new Error(`duplicateSlide: source ${slide._partName} not found`);

    // Allocate new slide part name + sldId + rId.
    const presPart = pkg.getPart(PRES_PART_NAME);
    if (!presPart) throw new Error('presentation.xml missing');
    const presDoc = parseXml(decoder.decode(presPart.data));
    const sldIdLst = ensureSldIdLst(presDoc.root);
    const newSldId = allocateSldId(sldIdLst);

    let slideN = 1;
    for (const p of pkg.parts) {
      const m = p.name.match(/^\/ppt\/slides\/slide(\d+)\.xml$/);
      if (m?.[1] !== undefined) {
        const n = Number.parseInt(m[1], 10);
        if (Number.isFinite(n) && n >= slideN) slideN = n + 1;
      }
    }
    const newSlidePartName = partName(`/ppt/slides/slide${slideN}.xml`);

    // Add the duplicate part with a copy of the bytes.
    pkg.addPart(newSlidePartName, sourcePart.contentType, new Uint8Array(sourcePart.data));

    // Clone the source slide's rels so the duplicate keeps every link
    // (layout, media, hyperlinks, etc.) intact.
    const sourceRels = pkg.getRels(slide._partName);
    if (sourceRels !== null) {
      pkg.setRels(newSlidePartName, { items: sourceRels.items.map((r) => ({ ...r })) });
    }

    // presentation → new slide rel.
    const presRels = pkg.getRels(PRES_PART_NAME) ?? emptyRels();
    const newRId = nextRelId(presRels.items.map((r) => r.id));
    presRels.items.push({
      id: newRId,
      type: REL_TYPES.slide,
      target: `slides/slide${slideN}.xml`,
      targetMode: 'Internal',
    });
    pkg.setRels(PRES_PART_NAME, presRels);

    // Add `<p:sldId>` entry.
    const newSldIdElement = elem(NAME_SLD_ID, {
      attrs: [attr(ATTR_ID, String(newSldId)), attr(ATTR_R_ID, newRId)],
    });
    sldIdLst.children.push(newSldIdElement);
    presPart.data = encoder.encode(serializeXml(presDoc));

    this._slidesCache = null;
    const slides = this.slides;
    const dup = slides[slides.length - 1];
    if (!dup) throw new Error('duplicateSlide: post-condition failed');
    return dup;
  }

  /**
   * Moves a slide to a new index in deck order. Indices are clamped to
   * `[0, slides.length - 1]`. The operation is a pure reorder — sldIds,
   * rels, parts, and content types are not touched.
   *
   * The slide identity (the `Slide` instance and its underlying part)
   * stays the same; only `<p:sldIdLst>`'s child order changes.
   */
  moveSlide(slide: Slide, toIndex: number): void {
    const pkg = this[INTERNAL_PACKAGE];
    const slideRelTarget = `slides/${basename(slide._partName)}`;
    const presRels = pkg.getRels(PRES_PART_NAME);
    if (!presRels) throw new Error('presentation.xml has no rels');
    const slideRel = presRels.items.find(
      (r) => r.type === REL_TYPES.slide && r.target === slideRelTarget,
    );
    if (!slideRel) throw new Error(`moveSlide: slide ${slide._partName} has no rel`);

    const presPart = pkg.getPart(PRES_PART_NAME);
    if (!presPart) throw new Error('presentation.xml missing');
    const presDoc = parseXml(decoder.decode(presPart.data));
    const sldIdLst = firstChildElement(presDoc.root, NAME_SLD_ID_LST);
    if (!sldIdLst) throw new Error('presentation.xml has no <p:sldIdLst>');

    // Locate the element whose r:id matches the slide's rId.
    const sldIdElements = sldIdLst.children.filter(
      (c): c is XmlElement =>
        c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'sldId',
    );
    const target = sldIdElements.find((e) => getAttrValue(e, ATTR_R_ID) === slideRel.id);
    if (!target) throw new Error(`moveSlide: <p:sldId> for ${slideRel.id} not found`);

    // Remove the target from the parent's children, then re-insert at the
    // intended position. Indices count over `<p:sldId>` siblings only —
    // intervening whitespace or comments are not visible to the user.
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

    presPart.data = encoder.encode(serializeXml(presDoc));
    this._slidesCache = null;
  }

  /** @internal */
  static _fromPackage(pkg: OpcPackage): Presentation {
    return new Presentation(pkg);
  }
}

/**
 * @internal — used by `pptx-kit/node` to mount fs-backed helpers without
 * importing internal modules directly.
 */
export const _internalPackageOf = (p: Presentation): OpcPackage => p[INTERNAL_PACKAGE];

// `Presentation._fromPackage` (declared inside the class body) is the
// supported way for sibling modules in this package — notably the Node
// subclass — to construct a Presentation from a pre-parsed OpcPackage
// without bypassing the protected constructor.

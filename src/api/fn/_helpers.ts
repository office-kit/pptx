// Cross-cutting helpers shared by the split fn modules.
//
// Helpers used by only one module live with that module. Anything
// referenced from two or more split files is centralized here.

import type { OpcPackage } from '../../internal/parts/index.ts';
import {
  REL_TYPES,
  readSlideLayoutPart,
  readSlidePart,
} from '../../internal/presentationml/index.ts';
import {
  NS,
  type XmlElement,
  elem,
  firstChildElement,
  getAttrValue,
  qname,
  serializeXml,
  walkElements,
} from '../../internal/xml/index.ts';
import { partName } from '../../internal/opc/index.ts';
import {
  INTERNAL_PACKAGE,
  LAYOUT_DOCUMENT,
  LAYOUT_PART,
  LAYOUT_PART_NAME,
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  SHAPE_SNAPSHOT,
  SLIDE_DOCUMENT,
  SLIDE_PART,
  SLIDE_PART_NAME,
  SLIDE_SHAPES,
  type SlideData,
  type SlideLayoutData,
  type SlideShapeData,
} from '../_internal-symbols.ts';

const TEXT_DECODER = new TextDecoder();
const TEXT_ENCODER = new TextEncoder();
export const decode = (b: Uint8Array): string => TEXT_DECODER.decode(b);
export const encode = (s: string): Uint8Array => TEXT_ENCODER.encode(s);

export const SLIDE_LAYOUT_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml';
export const SLIDE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
export const PRES_PART_NAME = partName('/ppt/presentation.xml');

export const NAME_PRESENTATION = qname('p', 'presentation', NS.pml);
export const NAME_SLD_MASTER_ID_LST = qname('p', 'sldMasterIdLst', NS.pml);
export const NAME_SLD_ID_LST = qname('p', 'sldIdLst', NS.pml);
export const NAME_SLD_ID = qname('p', 'sldId', NS.pml);
export const NAME_CSLD = qname('p', 'cSld', NS.pml);
export const NAME_SP_TREE = qname('p', 'spTree', NS.pml);
export const ATTR_ID = qname('', 'id', '');
export const ATTR_R_ID = qname('r', 'id', NS.officeDocRels);

// PowerPoint accepts sldIds in [256, 2³¹−1024]. See plan §Risks.
export const SLD_ID_MIN = 256;
export const SLD_ID_MAX = 2147482623;

// @internal — used by mutation functions to write SlideData state back
// into the package and rebuild the typed view. Free functions, no class
// dependency.

export const commitSlideData = (slide: SlideData): void => {
  const xml = serializeXml(slide[SLIDE_DOCUMENT]);
  const part = slide[INTERNAL_PACKAGE].getPart(slide[SLIDE_PART_NAME]);
  if (!part) throw new Error(`slide part missing: ${slide[SLIDE_PART_NAME]}`);
  part.data = encode(xml);
};

export const refreshSlideData = (slide: SlideData): void => {
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

export const commitLayoutData = (layout: SlideLayoutData): void => {
  const part = layout[INTERNAL_PACKAGE].getPart(layout[LAYOUT_PART_NAME]);
  if (!part) throw new Error(`slide layout part missing: ${layout[LAYOUT_PART_NAME]}`);
  part.data = encode(serializeXml(layout[LAYOUT_DOCUMENT]));
  layout[LAYOUT_PART] = readSlideLayoutPart(layout[LAYOUT_DOCUMENT].root);
};

// Rebuild shape handles entirely — used when the shape count changes
// (e.g. removeShape). Existing SlideShapeData identities are dropped;
// SHAPE_SLIDE back-pointers stay consistent because the SlideData
// reference is preserved.
export const rebuildShapesFromDocument = (slide: SlideData): void => {
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

const NAME_TX_BODY = qname('p', 'txBody', NS.pml);

export const requireSpPr = (shape: SlideShapeData): XmlElement => {
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

export const requireTxBody = (shape: SlideShapeData): XmlElement => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'shape') {
    throw new Error(
      `text operations require a shape kind; ${shape[SHAPE_SNAPSHOT].kind} is not text-bearing`,
    );
  }
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (txBody === null) {
    throw new Error(`shape "${shape[SHAPE_SNAPSHOT].name}" has no <p:txBody>`);
  }
  return txBody;
};

const NAME_BODY_PR = qname('a', 'bodyPr', NS.dml);
const NAME_LST_STYLE = qname('a', 'lstStyle', NS.dml);
const NAME_A_P = qname('a', 'p', NS.dml);

export const createTxBody = (): XmlElement =>
  elem(NAME_TX_BODY, {
    children: [elem(NAME_BODY_PR), elem(NAME_LST_STYLE), elem(NAME_A_P)],
  });

/**
 * Returns the shape's `<p:txBody>`, creating an empty one if absent.
 *
 * PowerPoint always gives an autoshape a text body so it can hold text the
 * moment you click in and type. A shape authored without text (e.g.
 * `addSlideShape` with no `text`) has none, so setting text later would
 * otherwise fail — this makes every text-bearing shape editable. Unlike
 * `requireTxBody`, it never throws for a missing body; it still throws for a
 * non-text-bearing shape kind (picture / table / etc.).
 */
export const ensureTxBody = (shape: SlideShapeData): XmlElement => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'shape') {
    throw new Error(
      `text operations require a shape kind; ${shape[SHAPE_SNAPSHOT].kind} is not text-bearing`,
    );
  }
  const existing = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (existing !== null) return existing;
  const txBody = createTxBody();
  // txBody is the last child of <p:sp>, after spPr / style.
  shape[SHAPE_ELEMENT].children.push(txBody);
  return txBody;
};

export const commitAndRefresh = (shape: SlideShapeData): void => {
  commitSlideData(shape[SHAPE_SLIDE]);
  refreshSlideData(shape[SHAPE_SLIDE]);
};

// A link's rel is owned by the `<a:hlinkClick>` that points at it, and nothing else
// in a slide can reach a `hyperlink` or `slide` rel. Left behind after the link is
// removed, a `slide` rel keeps a removed slide alive as a dependency (duplicateSlide
// then fails on it) and a `hyperlink` rel keeps the URL in the file.
const LINK_REL_TYPES: ReadonlySet<string> = new Set([REL_TYPES.hyperlink, REL_TYPES.slide]);

/** Drops the slide's link rels that no element references any more. */
export const releaseUnusedLinkRels = (slide: SlideData): void => {
  const pkg = slide[INTERNAL_PACKAGE];
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
  if (rels === null) return;

  const referenced = new Set<string>();
  walkElements(slide[SLIDE_DOCUMENT].root, (element) => {
    for (const a of element.attrs) {
      if (a.name.namespaceURI === NS.officeDocRels) referenced.add(a.value);
    }
  });

  const kept = rels.items.filter((rel) => !LINK_REL_TYPES.has(rel.type) || referenced.has(rel.id));
  if (kept.length === rels.items.length) return;
  pkg.setRels(slide[SLIDE_PART_NAME], { items: kept });
};

export const requireSpTree = (slide: SlideData): XmlElement => {
  const cSld = firstChildElement(slide[SLIDE_DOCUMENT].root, NAME_CSLD);
  if (!cSld) throw new Error('slide has no <p:cSld>');
  const spTree = firstChildElement(cSld, NAME_SP_TREE);
  if (!spTree) throw new Error('slide has no <p:spTree>');
  return spTree;
};

export const nextShapeId = (slide: SlideData): number => {
  let maxId = 0;
  const walk = (el: XmlElement): void => {
    if (el.name.namespaceURI === NS.pml && el.name.localName === 'cNvPr') {
      const id = Number(getAttrValue(el, qname('', 'id', '')));
      if (id > maxId) maxId = id;
    }
    for (const child of el.children) if (child.kind === 'element') walk(child);
  };
  walk(requireSpTree(slide));
  return Math.max(maxId, 1) + 1;
};

export const appendAndReturnNewShape = (slide: SlideData, child: XmlElement): SlideShapeData => {
  const spTree = requireSpTree(slide);
  spTree.children.push(child);
  commitSlideData(slide);
  const previousLength = slide[SLIDE_SHAPES].length;
  rebuildShapesFromDocument(slide);
  const created = slide[SLIDE_SHAPES][previousLength];
  if (!created) throw new Error('appendShape: post-condition failed');
  return created;
};

export const setOpcDefault = (pkg: OpcPackage, extension: string, contentType: string): void => {
  const has = pkg.contentTypes.defaults.some((d) => d.extension.toLowerCase() === extension);
  if (!has) pkg.contentTypes.defaults.push({ extension, contentType });
};

/** Locate the owning slide/group container without descending into shape content. */
export const findShapeParent = (shape: SlideShapeData): XmlElement | null => {
  const stack = [requireSpTree(shape[SHAPE_SLIDE])];
  const target = shape[SHAPE_ELEMENT];
  while (stack.length) {
    const parent = stack.pop()!;
    for (const child of parent.children) {
      if (child === target) return parent;
      if (
        child.kind === 'element' &&
        child.name.namespaceURI === NS.pml &&
        child.name.localName === 'grpSp'
      )
        stack.push(child);
    }
  }
  return null;
};

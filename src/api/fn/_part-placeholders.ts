// Placeholder views and edits shared by every master-like part: slide layouts,
// slide masters, and the notes and handout masters. These parts all wrap a
// `<p:cSld><p:spTree>` the way a slide does; what differs is who inherits the
// result.

import { readPosition, readSize } from '../../internal/drawingml/index.ts';
import type { SlideShape } from '../../internal/presentationml/index.ts';
import type { OpcPackage, Part } from '../../internal/parts/index.ts';
import type { PartName } from '../../internal/opc/index.ts';
import {
  NS,
  type XmlDocument,
  type XmlElement,
  firstChildElement,
  getAttrValue,
  parseFragment,
  parseXml,
  qname,
  serializeXml,
} from '../../internal/xml/index.ts';
import type { Emu } from '../units.ts';
import { NAME_CSLD, NAME_SP_TREE, decode, encode } from './_helpers.ts';
import type { ShapeBounds } from './shape-read-base.ts';

/**
 * Read-only view of one placeholder on a slide layout or master. Surfaces the
 * fields a slide-author cares about when binding a slide to a layout: which
 * slot is for the title, which is for the body, etc.
 */
export interface SlideLayoutPlaceholder {
  /** `<p:ph type="...">`. `null` when omitted — the spec default is `obj`. */
  readonly type: string | null;
  /** `<p:ph idx="...">`. `null` when omitted — spec default is `0`. */
  readonly idx: number | null;
  /** `<p:cNvPr name="...">` — what PowerPoint shows in the selection pane. */
  readonly name: string;
  /**
   * Position + size in EMU. A layout placeholder with no `<a:xfrm>` of its
   * own inherits them from the master. `null` when the placeholder has no
   * explicit transform.
   */
  readonly bounds: ShapeBounds | null;
}

// Only `p:sp` shapes carry placeholders in real templates; pictures and
// connectors can technically have `<p:ph>` per the schema but PowerPoint never
// authors that. Filter on the `<p:ph>` either way.
export const placeholderShapes = (shapes: ReadonlyArray<SlideShape>): SlideShape[] =>
  shapes.filter((shape) => shape.placeholderType !== null || shape.placeholderIdx !== null);

export const placeholderViews = (
  shapes: ReadonlyArray<SlideShape>,
): ReadonlyArray<SlideLayoutPlaceholder> =>
  placeholderShapes(shapes).map((shape) => {
    const pos = readPosition(shape.element, shape.kind);
    const size = readSize(shape.element, shape.kind);
    const bounds: ShapeBounds | null =
      pos === null || size === null
        ? null
        : { x: pos.x as Emu, y: pos.y as Emu, w: size.w as Emu, h: size.h as Emu };
    return { type: shape.placeholderType, idx: shape.placeholderIdx, name: shape.name, bounds };
  });

/** A parsed package part, written back with `commitPartDoc`. */
export interface PartDoc {
  readonly part: Part;
  readonly doc: XmlDocument;
}

export const readPartDoc = (pkg: OpcPackage, name: PartName): PartDoc => {
  const part = pkg.getPart(name);
  if (part === null) throw new Error(`part ${name} is missing`);
  return { part, doc: parseXml(decode(part.data)) };
};

export const commitPartDoc = ({ part, doc }: PartDoc): void => {
  part.data = encode(serializeXml(doc));
};

export const requirePartSpTree = (root: XmlElement): XmlElement => {
  const cSld = firstChildElement(root, NAME_CSLD);
  const spTree = cSld && firstChildElement(cSld, NAME_SP_TREE);
  if (!spTree) throw new Error(`<p:${root.name.localName}> has no <p:spTree>`);
  return spTree;
};

const NAME_NV_SP_PR = qname('p', 'nvSpPr', NS.pml);
const NAME_NV_PR = qname('p', 'nvPr', NS.pml);
const NAME_PH = qname('p', 'ph', NS.pml);
const ATTR_TYPE = qname('', 'type', '');
const ATTR_IDX = qname('', 'idx', '');

/** The `<p:ph>` of a top-level `<p:sp>`, or `null` for anything else. */
const phOf = (node: XmlElement): XmlElement | null => {
  if (node.name.namespaceURI !== NS.pml || node.name.localName !== 'sp') return null;
  const nvSpPr = firstChildElement(node, NAME_NV_SP_PR);
  const nvPr = nvSpPr && firstChildElement(nvSpPr, NAME_NV_PR);
  return nvPr && firstChildElement(nvPr, NAME_PH);
};

/** `<p:ph type>` of a top-level placeholder; an omitted type is ECMA-376's `obj`. */
export const placeholderTypeOf = (node: XmlElement): string | null => {
  const ph = phOf(node);
  return ph === null ? null : (getAttrValue(ph, ATTR_TYPE) ?? 'obj');
};

/** Every `idx` the part's top-level placeholders use. */
export const placeholderIndices = (spTree: XmlElement): number[] => {
  const out: number[] = [];
  for (const child of spTree.children) {
    if (child.kind !== 'element') continue;
    const ph = phOf(child);
    const idx = ph && getAttrValue(ph, ATTR_IDX);
    if (idx) out.push(Number(idx));
  }
  return out;
};

/** Drops the top-level placeholders whose type is in `types`. */
export const removePlaceholdersOfType = (spTree: XmlElement, types: ReadonlySet<string>): void => {
  spTree.children = spTree.children.filter((child) => {
    if (child.kind !== 'element') return true;
    const type = placeholderTypeOf(child);
    return type === null || !types.has(type);
  });
};

export const hasPlaceholderOfType = (spTree: XmlElement, types: ReadonlySet<string>): boolean =>
  spTree.children.some((child) => {
    if (child.kind !== 'element') return false;
    const type = placeholderTypeOf(child);
    return type !== null && types.has(type);
  });

/**
 * Parses a `<p:sp>` built by `internal/presentationml/default-masters` (which
 * carries its own namespace declarations so it parses on its own) and drops
 * those declarations: the part it joins already binds the prefixes.
 */
export const placeholderElement = (xml: string): XmlElement => {
  const element = parseFragment(xml);
  element.prefixDecls = new Map();
  return element;
};

/**
 * Inserts placeholder `element` of `type` keeping `order` — the part's
 * placeholder types in their canonical sequence — so a restored footer lands
 * after the body and before any later placeholder, where PowerPoint re-adds it.
 */
export const insertPlaceholderInOrder = (
  spTree: XmlElement,
  element: XmlElement,
  type: string,
  order: ReadonlyArray<string>,
): void => {
  const rank = order.indexOf(type);
  const before = spTree.children.findIndex((child) => {
    if (child.kind !== 'element') return false;
    const other = placeholderTypeOf(child);
    return other !== null && order.indexOf(other) > rank;
  });
  if (before < 0) spTree.children.push(element);
  else spTree.children.splice(before, 0, element);
};

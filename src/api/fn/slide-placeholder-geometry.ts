import { resetTextBodyFormatting } from '../../internal/drawingml/text-format.ts';
import { buildPlaceholderStub } from '../../internal/presentationml/slide-builder.ts';
import {
  readFlip,
  readPosition,
  readRotation,
  readSize,
  setFlip,
  setPosition,
  setRotation,
  setSize,
} from '../../internal/drawingml/index.ts';
import { partName, resolveTarget } from '../../internal/opc/index.ts';
import {
  REL_TYPES,
  type PlaceholderType,
  readShapeTreeFromCsldRoot,
  type SlideShape,
} from '../../internal/presentationml/index.ts';
import {
  NS,
  firstChildElement,
  parseXml,
  qname,
  type XmlElement,
  type XmlNode,
} from '../../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  LAYOUT_PART,
  LAYOUT_PART_NAME,
  SHAPE_ELEMENT,
  SHAPE_SNAPSHOT,
  SLIDE_DOCUMENT,
  type SlideData,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import {
  commitSlideData,
  decode,
  refreshSlideData,
  nextShapeId,
  requireSpTree,
  rebuildShapesFromDocument,
} from './_helpers.ts';
import { findSlidePlaceholder, getSlideLayout, getSlideShapes } from './shape-slide-read.ts';

const nvNames = {
  shape: 'nvSpPr',
  picture: 'nvPicPr',
  connector: 'nvCxnSpPr',
  graphicFrame: 'nvGraphicFramePr',
  group: 'nvGrpSpPr',
};
function placeholderElement(shape: SlideShape): XmlElement | null {
  const nv = firstChildElement(shape.element, qname('p', nvNames[shape.kind], NS.pml));
  const properties = nv && firstChildElement(nv, qname('p', 'nvPr', NS.pml));
  return properties ? firstChildElement(properties, qname('p', 'ph', NS.pml)) : null;
}
function topLevelElements(root: XmlElement): Set<XmlNode> {
  const csld = firstChildElement(root, qname('p', 'cSld', NS.pml));
  const tree = csld && firstChildElement(csld, qname('p', 'spTree', NS.pml));
  return new Set(tree?.children ?? []);
}
function masterType(type: string | null): string {
  if (type === 'ctrTitle') return 'title';
  if (type === null || type === 'obj' || type === 'subTitle') return 'body';
  return type;
}

/**
 * Restore top-level placeholders' position, size, rotation and flips from the
 * current layout (or its master when the layout inherits its geometry).
 * Text, formatting, IDs, relationships and other shapes are preserved. Slots
 * without matching layout geometry are left intact; deleted placeholders are
 * not recreated.
 *
 * A placeholder inside a group is left where it is. The layout states a
 * rectangle on the slide, while a grouped shape's own geometry is written in
 * its group's coordinate space — so restoring it would mean either tearing the
 * shape out of the arrangement it was grouped into, or inventing a rectangle
 * the layout never described. Its formatting and appearance are still reset by
 * the two functions below.
 *
 * Returns the number of placeholders restored.
 */
export const resetSlidePlaceholderGeometry = (slide: SlideData): number => {
  const layout = getSlideLayout(slide);
  if (!layout) return 0;
  const layoutElements = topLevelElements(layout[LAYOUT_PART].root);
  const slots = new Map<number, SlideShape>();
  for (const shape of layout[LAYOUT_PART].shapes) {
    if (
      layoutElements.has(shape.element) &&
      placeholderElement(shape) &&
      !slots.has(shape.placeholderIdx ?? 0)
    )
      slots.set(shape.placeholderIdx ?? 0, shape);
  }
  const masters = new Map<string, SlideShape>();
  const pkg = slide[INTERNAL_PACKAGE];
  const layoutName = partName(layout[LAYOUT_PART_NAME]);
  const masterRel = pkg
    .getRels(layoutName)
    ?.items.find((rel) => rel.type === REL_TYPES.slideMaster);
  const masterPart = masterRel && pkg.getPart(resolveTarget(layoutName, masterRel.target));
  if (masterPart) {
    const root = parseXml(decode(masterPart.data)).root;
    const elements = topLevelElements(root);
    for (const shape of readShapeTreeFromCsldRoot(root, 'sldMaster').shapes) {
      const type = masterType(shape.placeholderType);
      if (elements.has(shape.element) && placeholderElement(shape) && !masters.has(type))
        masters.set(type, shape);
    }
  }
  const elements = topLevelElements(slide[SLIDE_DOCUMENT].root);
  let count = 0;
  for (const shape of getSlideShapes(slide)) {
    const snapshot = shape[SHAPE_SNAPSHOT];
    if (!elements.has(shape[SHAPE_ELEMENT]) || !placeholderElement(snapshot)) continue;
    const slot = slots.get(snapshot.placeholderIdx ?? 0);
    if (!slot) continue;
    let source = slot;
    let pos = readPosition(source.element, source.kind);
    let size = readSize(source.element, source.kind);
    if (!pos || !size) {
      const master = masters.get(masterType(slot.placeholderType));
      if (!master) continue;
      source = master;
      pos = readPosition(source.element, source.kind);
      size = readSize(source.element, source.kind);
    }
    if (!pos || !size) continue;
    const target = shape[SHAPE_ELEMENT];
    setPosition(target, snapshot.kind, pos.x, pos.y);
    setSize(target, snapshot.kind, size.w, size.h);
    setRotation(target, snapshot.kind, readRotation(source.element, source.kind));
    setFlip(
      target,
      snapshot.kind,
      readFlip(source.element, source.kind) ?? { horizontal: false, vertical: false },
    );
    count++;
  }
  if (count) {
    commitSlideData(slide);
    refreshSlideData(slide);
  }
  return count;
};

/**
 * Restore missing slide placeholders using the current layout.
 * Existing content, geometry and formatting are untouched, including grouped
 * placeholders. New slots inherit their layout's geometry and text style;
 * layout prompt text and layout-only relationships are not copied.
 * Returns the number of added placeholders. Repeated calls do not add duplicates.
 */
export const addMissingSlidePlaceholders = (slide: SlideData): number => {
  const layout = getSlideLayout(slide);
  if (!layout) return 0;
  const present = new Set<number>();
  for (const shape of getSlideShapes(slide)) {
    const snapshot = shape[SHAPE_SNAPSHOT];
    if (placeholderElement(snapshot)) present.add(snapshot.placeholderIdx ?? 0);
  }
  const elements = topLevelElements(layout[LAYOUT_PART].root);
  const additions: XmlElement[] = [];
  let id = nextShapeId(slide);
  for (const slot of layout[LAYOUT_PART].shapes) {
    if (!elements.has(slot.element)) continue;
    const ph = placeholderElement(slot);
    const index = slot.placeholderIdx ?? 0;
    if (!ph || present.has(index)) continue;
    additions.push(buildPlaceholderStub(id++, ph));
    present.add(index);
  }
  if (!additions.length) return 0;
  insertShapes(slide, additions);
  return additions.length;
};

function insertShapes(slide: SlideData, additions: XmlElement[]): void {
  const tree = requireSpTree(slide);
  // Shape-tree extensions must remain after all shape children.
  const extension = tree.children.findIndex(
    (node) =>
      node.kind === 'element' &&
      node.name.namespaceURI === NS.pml &&
      node.name.localName === 'extLst',
  );
  tree.children.splice(extension < 0 ? tree.children.length : extension, 0, ...additions);
  commitSlideData(slide);
  rebuildShapesFromDocument(slide);
}

/**
 * Adds the single placeholder of `type` that the layout defines and the slide
 * is missing — how an editor inserts a slide number, date or footer into a
 * slide without also restoring every other slot the author deleted.
 *
 * `type` is an ECMA-376 `ST_PlaceholderType` token (`sldNum`, `dt`, `ftr`,
 * `title`, `body`, …), matched as `findSlidePlaceholder` matches it. Returns
 * the slide's placeholder of that type — the one just added or the one already
 * there — or `null` when the layout reserves no such slot. Like the added slots
 * of `addMissingSlidePlaceholders`, it starts empty and inherits its geometry
 * and text style from the layout and master.
 */
export const addSlidePlaceholder = (
  slide: SlideData,
  type: PlaceholderType,
): SlideShapeData | null => {
  const existing = findSlidePlaceholder(slide, type);
  if (existing) return existing;
  const layout = getSlideLayout(slide);
  if (!layout) return null;
  const elements = topLevelElements(layout[LAYOUT_PART].root);
  for (const slot of layout[LAYOUT_PART].shapes) {
    if (!elements.has(slot.element) || slot.placeholderType !== type) continue;
    const ph = placeholderElement(slot);
    if (!ph) continue;
    insertShapes(slide, [buildPlaceholderStub(nextShapeId(slide), ph)]);
    return findSlidePlaceholder(slide, type);
  }
  return null;
};

/**
 * Restore inherited text formatting on placeholders bound to the current
 * layout, including placeholders inside a group. Clears direct run, paragraph
 * and text-body appearance while preserving text, fields, hyperlinks,
 * language, outline levels and unknown extensions. Geometry and shape
 * appearance remain unchanged.
 * Returns the number of matching text placeholders processed.
 */
export const resetSlidePlaceholderTextFormatting = (slide: SlideData): number => {
  const layout = getSlideLayout(slide);
  if (!layout) return 0;
  const layoutElements = topLevelElements(layout[LAYOUT_PART].root);
  const slots = new Set<number>();
  for (const slot of layout[LAYOUT_PART].shapes) {
    if (layoutElements.has(slot.element) && placeholderElement(slot))
      slots.add(slot.placeholderIdx ?? 0);
  }
  let count = 0;
  // Unlike geometry, formatting means the same thing wherever the placeholder
  // sits: a group scales and turns what is inside it, it does not give the
  // text a font. So a grouped placeholder is reset like any other.
  for (const shape of getSlideShapes(slide)) {
    const snapshot = shape[SHAPE_SNAPSHOT];
    if (!placeholderElement(snapshot) || !slots.has(snapshot.placeholderIdx ?? 0)) continue;
    const body = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'txBody', NS.pml));
    if (!body) continue;
    resetTextBodyFormatting(body);
    count++;
  }
  if (count) {
    commitSlideData(slide);
    refreshSlideData(slide);
  }
  return count;
};

const LAYOUT_APPEARANCE_CHILDREN = new Set([
  'prstGeom',
  'custGeom',
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
  'ln',
  'effectLst',
  'effectDag',
  'scene3d',
  'sp3d',
]);

/**
 * Reset the current slide's layout: restore deleted slots, reset top-level
 * placeholder geometry and clear direct shape/text appearance to inherit the
 * layout again. Content, picture crops, relationships, IDs and unrelated shapes
 * remain intact.
 *
 * A placeholder inside a group has its appearance and text formatting reset
 * like any other, but keeps its geometry: where it sits is a statement about
 * the group's arrangement, and the layout has nothing to say about that.
 *
 * Returns the number of layout placeholders processed, including new slots.
 */
export const resetSlideLayout = (slide: SlideData): number => {
  const layout = getSlideLayout(slide);
  if (!layout) return 0;
  addMissingSlidePlaceholders(slide);
  resetSlidePlaceholderGeometry(slide);
  resetSlidePlaceholderTextFormatting(slide);
  const layoutElements = topLevelElements(layout[LAYOUT_PART].root);
  const slots = new Set<number>();
  for (const slot of layout[LAYOUT_PART].shapes) {
    if (layoutElements.has(slot.element) && placeholderElement(slot))
      slots.add(slot.placeholderIdx ?? 0);
  }
  let count = 0;
  for (const shape of getSlideShapes(slide)) {
    const snapshot = shape[SHAPE_SNAPSHOT];
    const element = shape[SHAPE_ELEMENT];
    if (!placeholderElement(snapshot) || !slots.has(snapshot.placeholderIdx ?? 0)) continue;
    const properties = firstChildElement(element, qname('p', 'spPr', NS.pml));
    if (properties) {
      properties.attrs = properties.attrs.filter(
        (a) => a.name.namespaceURI !== '' || a.name.localName !== 'bwMode',
      );
      // The picture's content is p:blipFill, outside spPr; keep it and its crop.
      // Geometry is already reset; leave xfrm and extension metadata in place.
      properties.children = properties.children.filter(
        (c) =>
          c.kind !== 'element' ||
          c.name.namespaceURI !== NS.dml ||
          !LAYOUT_APPEARANCE_CHILDREN.has(c.name.localName),
      );
    }
    // CT_ShapeStyle contains the four direct theme references. Removing this
    // optional node restores the corresponding layout's references.
    element.children = element.children.filter(
      (c) => c.kind !== 'element' || c.name.namespaceURI !== NS.pml || c.name.localName !== 'style',
    );
    count++;
  }
  if (count) {
    commitSlideData(slide);
    refreshSlideData(slide);
  }
  return count;
};

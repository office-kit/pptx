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
} from '../_internal-symbols.ts';
import { commitSlideData, decode, refreshSlideData } from './_helpers.ts';
import { getSlideLayout, getSlideShapes } from './shape-slide-read.ts';

const nvNames = {
  shape: 'nvSpPr',
  picture: 'nvPicPr',
  connector: 'nvCxnSpPr',
  graphicFrame: 'nvGraphicFramePr',
  group: 'nvGrpSpPr',
};
function hasPlaceholder(shape: SlideShape): boolean {
  const nv = firstChildElement(shape.element, qname('p', nvNames[shape.kind], NS.pml));
  const properties = nv && firstChildElement(nv, qname('p', 'nvPr', NS.pml));
  return !!properties && !!firstChildElement(properties, qname('p', 'ph', NS.pml));
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
 * not recreated. Grouped placeholders retain their group-relative geometry.
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
      hasPlaceholder(shape) &&
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
      if (elements.has(shape.element) && hasPlaceholder(shape) && !masters.has(type))
        masters.set(type, shape);
    }
  }
  const elements = topLevelElements(slide[SLIDE_DOCUMENT].root);
  let count = 0;
  for (const shape of getSlideShapes(slide)) {
    const snapshot = shape[SHAPE_SNAPSHOT];
    if (!elements.has(shape[SHAPE_ELEMENT]) || !hasPlaceholder(snapshot)) continue;
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

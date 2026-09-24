import { inkContentPart } from '../../internal/drawingml/ink-content.ts';
// Shape removal and z-order.

import { duplicatePartGraph } from '../../internal/parts/duplicate-graph.ts';
import { emptyRels, nextRelId, partName, resolveTarget } from '../../internal/opc/index.ts';
import { readPictureMediaRef, REL_TYPES } from '../../internal/presentationml/index.ts';
import {
  NS,
  type XmlElement,
  type XmlNode,
  firstChildElement,
  getAttrValue,
  qname,
} from '../../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  SLIDE_DOCUMENT,
  SLIDE_PART_NAME,
  type SlideData,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import {
  appendAndReturnNewShape,
  commitSlideData,
  nextShapeId,
  rebuildShapesFromDocument,
  requireSpTree,
  shapeParent,
} from './_helpers.ts';
import { addMediaTimingNode, removeMediaTimingNodes } from './_media-timing.ts';
// ---------------------------------------------------------------------------
// Shape mutation — removal.

/**
 * Copies a shape into `targetSlide`. The source XML is cloned and
 * appended to the target's `<p:spTree>`. Image rels on the source
 * shape are followed: the linked media part is referenced from the
 * target slide via a freshly allocated rId (no media bytes are
 * copied — both slides share the underlying part). Chart parts and their
 * editable dependencies are cloned so edits to the copy remain independent.
 *
 * v1 requires source and target to live in the same package
 * (`sourceShape`'s slide and `targetSlide` must share the same
 * `OpcPackage`). Cross-package copy is `importSlide` territory.
 *
 * With `sameParent`, the copy stays in its source group and retains local
 * coordinates. This option requires source and target to be the same slide.
 * Returns the new `SlideShapeData` on `targetSlide`.
 */
export const copyShape = (
  targetSlide: SlideData,
  sourceShape: SlideShapeData,
  options: { sameParent?: boolean } = {},
): SlideShapeData => {
  const sourceSlide = sourceShape[SHAPE_SLIDE];
  if (sourceSlide[INTERNAL_PACKAGE] !== targetSlide[INTERNAL_PACKAGE]) {
    throw new Error(
      'copyShape: source and target must be in the same package. Use importSlide for cross-deck copies.',
    );
  }
  const root = requireSpTree(targetSlide);
  const parent = options.sameParent
    ? sourceSlide === targetSlide && shapeParent(root, sourceShape[SHAPE_ELEMENT])
    : root;
  if (!parent)
    throw new Error('copyShape: sameParent requires an attached shape on the target slide.');
  const copied = cloneShapeInto(targetSlide, sourceShape);
  if (parent !== root) {
    root.children.splice(root.children.indexOf(copied[SHAPE_ELEMENT]), 1);
    parent.children.push(copied[SHAPE_ELEMENT]);
    commitSlideData(targetSlide);
    rebuildShapesFromDocument(targetSlide);
  }
  return copied;
};

/**
 * Imports a shape and its referenced package parts into another presentation.
 * Explicit formatting is retained; inherited theme/layout formatting uses the
 * destination slide. The source presentation is never changed.
 */
export const importShape = (targetSlide: SlideData, sourceShape: SlideShapeData): SlideShapeData =>
  cloneShapeInto(targetSlide, sourceShape);

const cloneShapeInto = (targetSlide: SlideData, sourceShape: SlideShapeData): SlideShapeData => {
  const sourceSlide = sourceShape[SHAPE_SLIDE];
  const sourcePkg = sourceSlide[INTERNAL_PACKAGE];
  const pkg = targetSlide[INTERNAL_PACKAGE];
  const sourceEl = sourceShape[SHAPE_ELEMENT];

  // Deep-clone the XML by serializing + re-parsing one element.
  // We wrap in a temporary parent so we can extract the cloned element
  // back out without ambient namespaces leaking from the slide root.
  const cloned = cloneXmlElement(sourceEl);

  // Allocate fresh IDs for the entire copied subtree, including group children.
  const newId = nextShapeId(targetSlide);
  rewriteCNvPrId(cloned, newId);

  // Walk the cloned element for r:embed / r:link references. For each
  // referenced rId in the source slide's rels, copy the rel onto the
  // target slide's rels (allocating a fresh rId) and update the cloned
  // attribute. This covers picture blips + media references.
  const sourceRels = sourcePkg.getRels(sourceSlide[SLIDE_PART_NAME]);
  if (sourceRels) {
    const targetRels = pkg.getRels(targetSlide[SLIDE_PART_NAME]) ?? emptyRels();
    const usedIds = new Set(targetRels.items.map((r) => r.id));
    const copiedDependencies = new Map<string, string>();
    rewriteRIdReferences(cloned, (oldRId) => {
      const sourceRel = sourceRels.items.find((r) => r.id === oldRId);
      if (!sourceRel) return oldRId;
      let target = sourceRel.target;
      if (sourceRel.targetMode !== 'External') {
        const sourcePart = resolveTarget(sourceSlide[SLIDE_PART_NAME], sourceRel.target);
        target = sourcePart;
        if (sourcePkg !== pkg || sourceRel.type === REL_TYPES.chart) {
          const key = sourcePart.toLowerCase();
          let copiedPart = copiedDependencies.get(key);
          if (!copiedPart) {
            const dot = sourcePart.lastIndexOf('.');
            const split = dot > sourcePart.lastIndexOf('/') ? dot : sourcePart.length;
            const stem = sourcePart.slice(0, split);
            const extension = sourcePart.slice(split);
            let n = 1;
            do {
              copiedPart = `${stem}-copy${n++}${extension}`;
            } while (pkg.getPart(partName(copiedPart)));
            duplicatePartGraph(
              pkg,
              sourcePart,
              partName(copiedPart),
              sourcePkg === pkg
                ? new Set([REL_TYPES.image, REL_TYPES.theme, REL_TYPES.slide])
                : new Set(),
              sourcePkg,
            );
            copiedDependencies.set(key, copiedPart);
          }
          target = copiedPart;
        }
      }
      // Look for an existing rel on target with the same type+target;
      // reuse if found to avoid duplicates.
      const existing = targetRels.items.find(
        (r) =>
          r.type === sourceRel.type && r.target === target && r.targetMode === sourceRel.targetMode,
      );
      if (existing) return existing.id;
      const newRId = nextRelId([...usedIds]);
      usedIds.add(newRId);
      targetRels.items.push({ ...sourceRel, id: newRId, target });
      return newRId;
    });
    pkg.setRels(targetSlide[SLIDE_PART_NAME], targetRels);
  }

  // A clip's play controls come from a media time node keyed by shape id, so
  // the copy needs its own node under the id it was just given.
  const media = cloned.name.localName === 'pic' ? readPictureMediaRef(cloned) : null;
  if (media !== null) addMediaTimingNode(targetSlide, media.kind, newId);

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
  const ids = new Map<string, string>();
  const walk = (el: XmlElement, visitor: (element: XmlElement) => void): void => {
    visitor(el);
    for (const child of el.children) if (child.kind === 'element') walk(child, visitor);
  };
  walk(root, (el) => {
    if (
      (el.name.namespaceURI !== NS.pml && el.name.namespaceURI !== NS.p14) ||
      el.name.localName !== 'cNvPr'
    )
      return;
    const attr = el.attrs.find((a) => a.name.namespaceURI === '' && a.name.localName === 'id');
    if (!attr) return;
    const replacement = ids.get(attr.value) ?? String(newId++);
    ids.set(attr.value, replacement);
    el.attrs = el.attrs.map((a) => (a === attr ? { ...a, value: replacement } : a));
  });
  // Connections within a copied group must point to its new children.
  walk(root, (el) => {
    if (el.name.namespaceURI !== NS.dml || !['stCxn', 'endCxn'].includes(el.name.localName)) return;
    const attr = el.attrs.find((a) => a.name.namespaceURI === '' && a.name.localName === 'id');
    const replacement = attr && ids.get(attr.value);
    if (attr && replacement)
      el.attrs = el.attrs.map((a) => (a === attr ? { ...a, value: replacement } : a));
  });
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
// Z-order — move shapes forward / backward inside their immediate parent.
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

const isShapeChild = (node: XmlNode): boolean =>
  node.kind === 'element' &&
  ((node.name.namespaceURI === NS.pml && SHAPE_CHILD_LOCALS.has(node.name.localName)) ||
    inkContentPart(node) !== null);

/** Move `shape` in front of its siblings in the slide or group. */
export const bringShapeToFront = (shape: SlideShapeData): void => {
  const slide = shape[SHAPE_SLIDE];
  const spTree = shapeParent(requireSpTree(slide), shape[SHAPE_ELEMENT]);
  if (!spTree) return;
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
 * Move `shape` behind every other shape in its parent. The
 * `<p:nvGrpSpPr>` / `<p:grpSpPr>` preface — required by the schema —
 * stays at the top.
 */
export const sendShapeToBack = (shape: SlideShapeData): void => {
  const slide = shape[SHAPE_SLIDE];
  const spTree = shapeParent(requireSpTree(slide), shape[SHAPE_ELEMENT]);
  if (!spTree) return;
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
  const spTree = shapeParent(requireSpTree(slide), shape[SHAPE_ELEMENT]);
  if (!spTree) return;
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
 * Returns the shape's z-index among its parent's "real" shape children
 * (`<p:sp>` / `<p:pic>` / `<p:cxnSp>` / `<p:graphicFrame>` / `<p:grpSp>`),
 * skipping the required `<p:nvGrpSpPr>` / `<p:grpSpPr>` preface.
 * Higher numbers render in front.
 */
export const getShapeZIndex = (shape: SlideShapeData): number => {
  const slide = shape[SHAPE_SLIDE];
  const spTree = shapeParent(requireSpTree(slide), shape[SHAPE_ELEMENT]);
  if (!spTree) return -1;
  let i = 0;
  for (const c of spTree.children) {
    if (!isShapeChild(c)) continue;
    if (c === shape[SHAPE_ELEMENT]) return i;
    i++;
  }
  return -1;
};

/**
 * Moves the shape to a specific z-index among its parent's "real"
 * shape children. Index is clamped to the available range. Higher
 * numbers render in front. The required preface elements stay at the
 * top of `<p:spTree>`.
 */
export const setShapeZIndex = (shape: SlideShapeData, toIndex: number): void => {
  const slide = shape[SHAPE_SLIDE];
  const spTree = shapeParent(requireSpTree(slide), shape[SHAPE_ELEMENT]);
  if (!spTree) return;
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
  const spTree = shapeParent(requireSpTree(slide), shape[SHAPE_ELEMENT]);
  if (!spTree) return;
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
  const removedIds = new Set<number>();
  spTree.children = spTree.children.filter((c) => {
    const isShape =
      c.kind === 'element' &&
      c.name.namespaceURI === NS.pml &&
      SHAPE_CHILD_LOCALS.has(c.name.localName);
    if (isShape) collectShapeIds(c, removedIds);
    return !isShape;
  });
  removeMediaTimingNodes(slide, removedIds);
  commitSlideData(slide);
  rebuildShapesFromDocument(slide);
};

export const removeShape = (shape: SlideShapeData): void => {
  const slide = shape[SHAPE_SLIDE];
  const doc = slide[SLIDE_DOCUMENT];
  const cSld = firstChildElement(doc.root, qname('p', 'cSld', NS.pml));
  if (!cSld) return;
  const root = firstChildElement(cSld, qname('p', 'spTree', NS.pml));
  const spTree = root && shapeParent(root, shape[SHAPE_ELEMENT]);
  if (!spTree) return;
  const idx = spTree.children.indexOf(shape[SHAPE_ELEMENT]);
  if (idx < 0) return;
  spTree.children.splice(idx, 1);
  const removedIds = new Set<number>();
  collectShapeIds(shape[SHAPE_ELEMENT], removedIds);
  removeMediaTimingNodes(slide, removedIds);
  commitSlideData(slide);
  rebuildShapesFromDocument(slide);
};

// Ids of `el` and every shape nested in it (a removed group takes its media
// pictures with it), so no media time node is left targeting a shape id that
// no longer exists.
const collectShapeIds = (el: XmlElement, into: Set<number>): void => {
  if (el.name.namespaceURI === NS.pml && el.name.localName === 'cNvPr') {
    const id = Number.parseInt(getAttrValue(el, qname('', 'id', '')) ?? '', 10);
    if (Number.isFinite(id)) into.add(id);
    return;
  }
  for (const c of el.children) {
    if (c.kind === 'element') collectShapeIds(c, into);
  }
};

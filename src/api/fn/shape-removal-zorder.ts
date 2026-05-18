// Shape removal and z-order.

import { emptyRels, nextRelId } from '../../internal/opc/index.ts';
import { NS, type XmlElement, firstChildElement, qname } from '../../internal/xml/index.ts';
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
} from './_helpers.ts';
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

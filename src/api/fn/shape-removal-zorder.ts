// Shape removal and z-order.

import { emptyRels, nextRelId, resolveTarget, type PartName } from '../../internal/opc/index.ts';
import { copyPartGraphs } from '../../internal/parts/duplicate-graph.ts';
import { readPictureMediaRef } from '../../internal/presentationml/index.ts';
import {
  NS,
  type XmlElement,
  firstChildElement,
  getAttrValue,
  qname,
} from '../../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
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
  findShapeParent,
} from './_helpers.ts';
import { addMediaTimingNode, removeMediaTimingNodes } from './_media-timing.ts';
import { planAnimationCopy } from './shape-animation.ts';
import { planShapeAnimationRemoval } from './slide-animation-edit.ts';
// ---------------------------------------------------------------------------
// Shape mutation — removal.

/**
 * Copies a shape into `targetSlide`, including nested shapes and relationships.
 * Within a presentation, media and other referenced parts remain shared.
 * Across presentations, dependencies are cloned with collision-free part names,
 * retaining embedded workbooks, unknown parts and external links.
 * With `preserveGroupTransform`, a nested source retains its ancestor group
 * transforms as wrappers, excluding sibling objects. Returns the outermost
 * copied shape on the target slide.
 */
export const copyShape = (
  targetSlide: SlideData,
  sourceShape: SlideShapeData,
  opts: { readonly preserveGroupTransform?: boolean } = {},
): SlideShapeData => {
  const sourceSlide = sourceShape[SHAPE_SLIDE];
  const sourcePkg = sourceSlide[INTERNAL_PACKAGE];
  const pkg = targetSlide[INTERNAL_PACKAGE];
  const sourceEl = sourceShape[SHAPE_ELEMENT];

  let cloned = cloneXmlElement(sourceEl);
  if (opts.preserveGroupTransform) {
    // Keep only the copied branch inside its ancestor groups. Retaining their
    // transforms also preserves shear from rotated, nonuniformly scaled groups,
    // which cannot be represented by a standalone shape's xfrm.
    const root = requireSpTree(sourceSlide);
    const parents = new Map<XmlElement, XmlElement>();
    const stack = [root];
    while (stack.length) {
      const parent = stack.pop()!;
      for (const child of parent.children) {
        if (child.kind !== 'element' || !isShapeChild(child)) continue;
        parents.set(child, parent);
        if (child.name.localName === 'grpSp') stack.push(child);
      }
    }
    let branch = sourceEl;
    let parent = parents.get(branch);
    while (parent && parent !== root) {
      const wrapper = cloneXmlElement({
        ...parent,
        children: parent.children.filter((child) => !isShapeChild(child)),
      });
      const insertion = parent.children
        .slice(0, parent.children.indexOf(branch))
        .filter((child) => !isShapeChild(child)).length;
      wrapper.children.splice(insertion, 0, cloned);
      cloned = wrapper;
      branch = parent;
      parent = parents.get(branch);
    }
  }
  const newId = nextShapeId(targetSlide);
  const copiedIds = rewriteShapeIds(cloned, newId);
  // A copy animates on its own: the effects the original carries are cloned
  // against the ids the copy has just been given, group children included.
  // Planned here, ahead of the parts and relationships, so a copy this library
  // cannot reproduce faithfully leaves the target slide untouched rather than
  // half-built.
  const copyAnimations = planAnimationCopy(sourceSlide, targetSlide, copiedIds);

  const sourceRels = sourcePkg.getRels(sourceSlide[SLIDE_PART_NAME]);
  const relsById = new Map(sourceRels?.items.map((rel) => [rel.id, rel]));
  const referencedIds = new Set<string>();
  rewriteRIdReferences(cloned, (id) => {
    referencedIds.add(id);
    return id;
  });
  const roots = new Map<PartName, null>();
  for (const id of referencedIds) {
    const rel = relsById.get(id);
    if (!rel) throw new Error(`copyShape: missing relationship ${id}`);
    if (rel.targetMode !== 'External')
      roots.set(resolveTarget(sourceSlide[SLIDE_PART_NAME], rel.target), null);
  }
  const copies =
    sourcePkg === pkg
      ? null
      : copyPartGraphs(
          sourcePkg,
          pkg,
          roots,
          new Set(),
          new Map([[sourceSlide[SLIDE_PART_NAME], targetSlide[SLIDE_PART_NAME]]]),
        );
  const targetRels = pkg.getRels(targetSlide[SLIDE_PART_NAME]) ?? emptyRels();
  const relIds = new Map<string, string>();
  const usedIds = targetRels.items.map((rel) => rel.id);
  const relationshipKey = (type: string, target: string, mode: string): string =>
    JSON.stringify([type, target, mode]);
  const existingRels = new Map(
    targetRels.items.map((rel) => [
      relationshipKey(
        rel.type,
        rel.targetMode === 'External'
          ? rel.target
          : resolveTarget(targetSlide[SLIDE_PART_NAME], rel.target),
        rel.targetMode,
      ),
      rel.id,
    ]),
  );
  for (const id of referencedIds) {
    const rel = relsById.get(id)!;
    const resolved =
      rel.targetMode === 'External'
        ? rel.target
        : resolveTarget(sourceSlide[SLIDE_PART_NAME], rel.target);
    const target = copies?.get(resolved.toLowerCase()) ?? resolved;
    const key = relationshipKey(rel.type, target, rel.targetMode);
    const existing = existingRels.get(key);
    const newRId = existing ?? nextRelId(usedIds);
    if (!existing) {
      usedIds.push(newRId);
      targetRels.items.push({ ...rel, target, id: newRId });
      existingRels.set(key, newRId);
    }
    relIds.set(id, newRId);
  }
  rewriteRIdReferences(cloned, (id) => relIds.get(id)!);
  pkg.setRels(targetSlide[SLIDE_PART_NAME], targetRels);

  copyAnimations();

  // A clip's play controls come from a media time node keyed by shape id, so
  // the copy needs its own node under the id it was just given. It goes into
  // the timing the animations have already landed in.
  const addMediaControls = (el: XmlElement): void => {
    if (el.name.namespaceURI === NS.pml && el.name.localName === 'pic') {
      const media = readPictureMediaRef(el);
      const nv = firstChildElement(el, qname('p', 'nvPicPr', NS.pml));
      const props = nv && firstChildElement(nv, qname('p', 'cNvPr', NS.pml));
      if (media && props)
        addMediaTimingNode(
          targetSlide,
          media.kind,
          Number(getAttrValue(props, qname('', 'id', ''))),
        );
    }
    for (const child of el.children) if (child.kind === 'element') addMediaControls(child);
  };
  addMediaControls(cloned);

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

const rewriteShapeIds = (root: XmlElement, firstId: number): Map<string, string> => {
  const ids = new Map<string, string>();
  let nextId = firstId;
  const walk = (el: XmlElement, rewriteReferences: boolean): void => {
    if (
      el.name.namespaceURI === (rewriteReferences ? NS.dml : NS.pml) &&
      (rewriteReferences
        ? ['stCxn', 'endCxn'].includes(el.name.localName)
        : el.name.localName === 'cNvPr')
    ) {
      el.attrs = el.attrs.map((a) => {
        if (a.name.namespaceURI !== '' || a.name.localName !== 'id') return a;
        if (rewriteReferences) return { ...a, value: ids.get(a.value) ?? a.value };
        const value = String(nextId++);
        ids.set(a.value, value);
        return { ...a, value };
      });
    }
    for (const child of el.children) if (child.kind === 'element') walk(child, rewriteReferences);
  };
  walk(root, false);
  walk(root, true);
  return ids;
};

const rewriteRIdReferences = (root: XmlElement, map: (oldRId: string) => string): void => {
  const walk = (el: XmlElement): void => {
    el.attrs = el.attrs.map((a) => {
      if (
        a.value !== '' &&
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
// Z-order — move shapes forward / backward inside the slide or group container.
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

const lastShapeIndex = (parent: XmlElement): number => {
  for (let i = parent.children.length - 1; i >= 0; i--) {
    if (isShapeChild(parent.children[i]!)) return i;
  }
  return -1;
};

/** Move `shape` in front of all sibling shapes in its parent container. */
export const bringShapeToFront = (shape: SlideShapeData): void => {
  const slide = shape[SHAPE_SLIDE];
  const spTree = findShapeParent(shape);
  if (!spTree) return;
  const target = shape[SHAPE_ELEMENT];
  const idx = spTree.children.indexOf(target);
  if (idx < 0) return;
  const last = lastShapeIndex(spTree);
  if (idx === last) return;
  spTree.children.splice(idx, 1);
  spTree.children.splice(last, 0, target);
  commitSlideData(slide);
  rebuildShapesFromDocument(slide);
};

/**
 * Move `shape` behind every other sibling shape. The
 * `<p:nvGrpSpPr>` / `<p:grpSpPr>` preface — required by the schema —
 * stays at the top.
 */
export const sendShapeToBack = (shape: SlideShapeData): void => {
  const slide = shape[SHAPE_SLIDE];
  const spTree = findShapeParent(shape);
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
  const spTree = findShapeParent(shape);
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
 * Returns the shape's z-index among its parent container's "real" shape children
 * (`<p:sp>` / `<p:pic>` / `<p:cxnSp>` / `<p:graphicFrame>` / `<p:grpSp>`),
 * skipping the required `<p:nvGrpSpPr>` / `<p:grpSpPr>` preface.
 * Higher numbers render in front.
 */
export const getShapeZIndex = (shape: SlideShapeData): number => {
  const spTree = findShapeParent(shape);
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
 * Moves the shape to a specific z-index among its parent container's "real"
 * shape children. Index is clamped to the available range. Higher
 * numbers render in front. The required preface elements stay at the
 * top of `<p:spTree>`.
 */
export const setShapeZIndex = (shape: SlideShapeData, toIndex: number): void => {
  const slide = shape[SHAPE_SLIDE];
  const spTree = findShapeParent(shape);
  if (!spTree) return;
  const target = shape[SHAPE_ELEMENT];
  const originalIndex = spTree.children.indexOf(target);
  const allShapeChildren = spTree.children.filter((c): c is XmlElement => isShapeChild(c));
  const clamped = Math.max(0, Math.min(toIndex, allShapeChildren.length - 1));

  // Remove the target from the tree, then re-insert at the position
  // corresponding to z-index `clamped` among the remaining shapes.
  spTree.children = spTree.children.filter((c) => c !== target);
  const remainingShapes = spTree.children.filter((c): c is XmlElement => isShapeChild(c));
  if (clamped >= remainingShapes.length) {
    const last = lastShapeIndex(spTree);
    spTree.children.splice(last < 0 ? originalIndex : last + 1, 0, target);
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
  const spTree = findShapeParent(shape);
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
  for (const child of spTree.children) {
    if (child.kind === 'element' && isShapeChild(child)) collectShapeIds(child, removedIds);
  }
  // Before the shapes go, so a slide whose timing could not survive losing
  // them keeps both the shapes and their animations.
  const dropAnimations = planShapeAnimationRemoval(slide, removedIds);

  spTree.children = spTree.children.filter((c) => !isShapeChild(c));
  dropAnimations();
  removeMediaTimingNodes(slide, removedIds);
  commitSlideData(slide);
  rebuildShapesFromDocument(slide);
};

export const removeShape = (shape: SlideShapeData): void => {
  const slide = shape[SHAPE_SLIDE];
  const spTree = findShapeParent(shape);
  if (!spTree) return;
  const idx = spTree.children.indexOf(shape[SHAPE_ELEMENT]);
  if (idx < 0) return;
  const removedIds = new Set<number>();
  collectShapeIds(shape[SHAPE_ELEMENT], removedIds);
  // Before the shape goes, so a slide whose timing could not survive losing it
  // keeps both the shape and its animations.
  const dropAnimations = planShapeAnimationRemoval(slide, removedIds);

  spTree.children.splice(idx, 1);
  dropAnimations();
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

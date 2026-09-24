// Group / ungroup shapes (`<p:grpSp>`).
//
// PowerPoint's "Group" command wraps a selection's shapes in a `<p:grpSp>`
// whose transform maps a child coordinate space onto the slide. At
// creation time the child space is set 1:1 with the group's own slide-space
// bounds (`chOff == off`, `chExt == ext`) — see `group-builder.ts`. That
// convention is what makes `ungroupShapes` a pure structural move in the
// common case (no move/resize since grouping): children already carry
// slide-space coordinates, so no rescale is needed. If the group *was*
// subsequently moved or resized (its `off`/`ext` diverged from its
// `chOff`/`chExt`), ungrouping rescales each child's own transform so it
// keeps its on-slide position and size.

import {
  readFlip,
  readRotation,
  readPosition,
  readSize,
  setFlip as writeFlip,
  setRotation as writeRotation,
  setPosition as writePosition,
  setSize as writeSize,
} from '../../internal/drawingml/index.ts';
import { buildGroup, readGroupChildren } from '../../internal/presentationml/index.ts';
import { type XmlElement } from '../../internal/xml/index.ts';
import {
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  SHAPE_SNAPSHOT,
  SLIDE_SHAPES,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import {
  commitSlideData,
  nextShapeId,
  rebuildShapesFromDocument,
  requireSpTree,
  shapeParent,
} from './_helpers.ts';
import { getGroupTransform } from './shape-read-base.ts';

// Session-only membership. Element identity prevents reused shape IDs joining it.
const formerGroups = new WeakMap<XmlElement, readonly XmlElement[]>();

/** Restore session membership after an application's internal serialization round trip. */
export const rememberRegroupShapes = (shapes: ReadonlyArray<SlideShapeData>): void => {
  if (shapes.length < 2) throw new Error('rememberRegroupShapes: at least 2 shapes are required');
  const slide = shapes[0]![SHAPE_SLIDE];
  const parent = shapeParent(requireSpTree(slide), shapes[0]![SHAPE_ELEMENT]);
  const elements = shapes.map((shape) => shape[SHAPE_ELEMENT]);
  if (
    !parent ||
    new Set(elements).size !== elements.length ||
    shapes.some(
      (shape) => shape[SHAPE_SLIDE] !== slide || !parent.children.includes(shape[SHAPE_ELEMENT]),
    )
  )
    throw new Error('rememberRegroupShapes: expected distinct attached siblings');
  for (const element of elements) formerGroups.set(element, elements);
};

/** Available siblings from the shape's most recently ungrouped group. */
export const getRegroupShapes = (shape: SlideShapeData): ReadonlyArray<SlideShapeData> => {
  const element = shape[SHAPE_ELEMENT];
  const members = formerGroups.get(element);
  if (!members) return [];
  const slide = shape[SHAPE_SLIDE];
  const parent = shapeParent(requireSpTree(slide), element);
  if (!parent) return [];
  const available = slide[SLIDE_SHAPES].filter(
    (candidate) =>
      members.includes(candidate[SHAPE_ELEMENT]) &&
      parent.children.includes(candidate[SHAPE_ELEMENT]),
  );
  return available.length >= 2 ? available : [];
};

/** Restore the first eligible former group represented by the selection. */
export const regroupShapes = (shapes: ReadonlyArray<SlideShapeData>): SlideShapeData => {
  for (const shape of shapes) {
    const members = getRegroupShapes(shape);
    if (members.length) return groupShapes(members);
  }
  throw new Error('regroupShapes: no previously ungrouped group is available');
};

/**
 * Groups two or more sibling shapes into a single `<p:grpSp>`,
 * returning the new group as a `SlideShapeData`. The group's
 * slide-space bounds are the union of its members' bounds; the members
 * keep their own relative position/size (nothing is rescaled). The
 * target slide is taken from the first shape — every shape must belong
 * to the same slide.
 *
 * Every shape must:
 *   - belong to the same slide as the others,
 *   - share the same immediate parent (slide shape tree or group),
 *   - have an explicit `<a:xfrm>` (placeholders that inherit position
 *     from the layout have none and can't be grouped),
 *   - appear at most once in `shapes` (grouping the same shape twice
 *     would duplicate its id).
 *
 * The group replaces its members at the position of the earliest one in
 * z-order, so grouping doesn't change how the selection stacks against
 * shapes that weren't part of it.
 */
export const groupShapes = (
  shapes: ReadonlyArray<SlideShapeData>,
  opts: { name?: string } = {},
): SlideShapeData => {
  if (shapes.length < 2) {
    throw new Error('groupShapes: at least 2 shapes are required');
  }
  const slide = shapes[0]![SHAPE_SLIDE];
  const spTree = shapeParent(requireSpTree(slide), shapes[0]![SHAPE_ELEMENT]);
  if (!spTree) throw new Error('groupShapes: shape is not attached to the slide');

  const elements: XmlElement[] = [];
  const seen = new Set<XmlElement>();
  let minX = Number.POSITIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  let maxY = Number.NEGATIVE_INFINITY;
  for (const shape of shapes) {
    if (shape[SHAPE_SLIDE] !== slide) {
      throw new Error('groupShapes: all shapes must belong to the same slide');
    }
    const el = shape[SHAPE_ELEMENT];
    if (seen.has(el)) {
      throw new Error(`groupShapes: shape "${shape[SHAPE_SNAPSHOT].name}" was passed twice`);
    }
    seen.add(el);
    if (!spTree.children.includes(el)) {
      throw new Error(
        `groupShapes: shape "${shape[SHAPE_SNAPSHOT].name}" is not a direct child of the ` +
          'same parent as the other selected shapes',
      );
    }
    const kind = shape[SHAPE_SNAPSHOT].kind;
    const pos = readPosition(el, kind);
    const size = readSize(el, kind);
    if (pos === null || size === null) {
      throw new Error(
        `groupShapes: shape "${shape[SHAPE_SNAPSHOT].name}" has no explicit position/size ` +
          "(placeholders that inherit geometry from the layout can't be grouped)",
      );
    }
    elements.push(el);
    minX = Math.min(minX, pos.x);
    minY = Math.min(minY, pos.y);
    maxX = Math.max(maxX, pos.x + size.w);
    maxY = Math.max(maxY, pos.y + size.h);
  }

  elements.sort((a, b) => spTree.children.indexOf(a) - spTree.children.indexOf(b));

  // Build the group element before touching `spTree.children` — it
  // validates the computed bounds as EMU coordinates and can throw. Doing
  // that after removing the members from the tree would leave the slide
  // missing shapes with no group to replace them.
  const grp = buildGroup({
    id: nextShapeId(slide),
    ...(opts.name !== undefined ? { name: opts.name } : {}),
    x: minX,
    y: minY,
    w: maxX - minX,
    h: maxY - minY,
    children: elements,
  });

  const insertAt = Math.min(...elements.map((el) => spTree.children.indexOf(el)));
  spTree.children = spTree.children.filter((c) => c.kind !== 'element' || !seen.has(c));
  spTree.children.splice(insertAt, 0, grp);
  for (const element of elements) formerGroups.delete(element);

  commitSlideData(slide);
  rebuildShapesFromDocument(slide);
  const created = slide[SLIDE_SHAPES].find((s) => s[SHAPE_ELEMENT] === grp);
  if (!created) throw new Error('groupShapes: post-condition failed');
  return created;
};

/**
 * Reverses `groupShapes`: removes the `<p:grpSp>` and re-inserts its
 * immediate children into its parent at the group's former position,
 * rescaling each child's own transform so it keeps its on-slide position
 * and size (matters when the group was moved/resized after creation, so
 * its `off`/`ext` diverged from its `chOff`/`chExt`). Returns the
 * children as fresh `SlideShapeData` handles, in their original order.
 *
 * Throws if `group` isn't a group shape, or its `<p:grpSpPr>` carries no
 * `<a:xfrm>` (malformed — every authored group has one).
 */
export const ungroupShapes = (group: SlideShapeData): ReadonlyArray<SlideShapeData> => {
  if (group[SHAPE_SNAPSHOT].kind !== 'group') {
    throw new Error('ungroupShapes: shape is not a group');
  }
  const transform = getGroupTransform(group);
  if (!transform) {
    throw new Error('ungroupShapes: group has no <a:xfrm> on <p:grpSpPr>');
  }
  const { outer, inner } = transform;
  const scaleX = inner.w === 0 ? 1 : outer.w / inner.w;
  const scaleY = inner.h === 0 ? 1 : outer.h / inner.h;

  const slide = group[SHAPE_SLIDE];
  const groupEl = group[SHAPE_ELEMENT];
  const spTree = shapeParent(requireSpTree(slide), groupEl);
  if (!spTree) throw new Error('ungroupShapes: group is not attached to the slide');
  const idx = spTree.children.indexOf(groupEl);
  if (idx < 0) throw new Error('ungroupShapes: group is not attached to the slide');

  const rotation = readRotation(groupEl, 'group');
  const flip = readFlip(groupEl, 'group');
  const radians = (rotation * Math.PI) / 180;
  const cos = Math.cos(radians),
    sin = Math.sin(radians);
  const horizontal = flip?.horizontal ?? false,
    vertical = flip?.vertical ?? false;
  const children = readGroupChildren(groupEl);
  const childElements = children.map((child) => {
    const pos = readPosition(child.element, child.kind);
    const size = readSize(child.element, child.kind);
    if (pos !== null && size !== null) {
      const width = size.w * scaleX,
        height = size.h * scaleY;
      const dx = ((pos.x + size.w / 2 - inner.x) * scaleX - outer.w / 2) * (horizontal ? -1 : 1);
      const dy = ((pos.y + size.h / 2 - inner.y) * scaleY - outer.h / 2) * (vertical ? -1 : 1);
      writePosition(
        child.element,
        child.kind,
        Math.round(outer.x + outer.w / 2 + dx * cos - dy * sin - width / 2),
        Math.round(outer.y + outer.h / 2 + dx * sin + dy * cos - height / 2),
      );
      writeSize(child.element, child.kind, Math.round(width), Math.round(height));
      if (rotation || horizontal || vertical) {
        const childRotation = readRotation(child.element, child.kind);
        const childFlip = readFlip(child.element, child.kind);
        writeRotation(
          child.element,
          child.kind,
          rotation + (horizontal !== vertical ? -childRotation : childRotation),
        );
        writeFlip(child.element, child.kind, {
          horizontal: horizontal !== (childFlip?.horizontal ?? false),
          vertical: vertical !== (childFlip?.vertical ?? false),
        });
      }
    }
    return child.element;
  });

  spTree.children.splice(idx, 1, ...childElements);

  for (const element of childElements) formerGroups.set(element, childElements);

  commitSlideData(slide);
  rebuildShapesFromDocument(slide);
  const byElement = new Map(slide[SLIDE_SHAPES].map((s) => [s[SHAPE_ELEMENT], s] as const));
  return childElements.map((el) => {
    const found = byElement.get(el);
    if (!found) throw new Error('ungroupShapes: post-condition failed');
    return found;
  });
};

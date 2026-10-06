import {
  getGroupChildren,
  getShapeBounds,
  getShapeId,
  getShapeKind,
  getShapeRotation,
  getShapeXmlString,
  getSlides,
  getSlideShapes,
  getSlideXmlString,
  type PresentationData,
  type SlideShapeData,
} from '@office-kit/pptx';
import type { PropEdit } from '@office-kit/pptx-dsl/source-edit';

export type DeckChange =
  | { kind: 'geometry'; slide: number; shapeId: number; edits: PropEdit[] }
  | { kind: 'unsupported'; slide: number | null; shapeId: number | null; description: string };

const ANGLE_UNITS = 60000;

interface Located {
  shape: SlideShapeData;
  /** Group members are positioned in their group's child space, not the slide's. */
  grouped: boolean;
}
function shapesById(shapes: readonly SlideShapeData[]): Map<number, Located> {
  const result = new Map<number, Located>();
  const visit = (shape: SlideShapeData, grouped: boolean) => {
    result.set(getShapeId(shape), { shape, grouped });
    for (const child of getGroupChildren(shape)) visit(child, true);
  };
  for (const shape of shapes) visit(shape, false);
  return result;
}

// The shape's own transform carries what geometry edits change; everything
// else in its XML must match for the change to be geometry alone.
function withoutTransform(xml: string): string {
  return xml
    .replace(/(<a:xfrm\b[^>]*?) rot="-?\d+"/, '$1')
    .replace(/<a:off x="-?\d+" y="-?\d+"\/>/, '')
    .replace(/<a:ext cx="\d+" cy="\d+"\/>/, '');
}

// Shapes are compared one by one; the rest of the slide (background, timing,
// transition) must match as a whole.
function withoutShapes(xml: string): string {
  return xml.replace(/<p:spTree>[\s\S]*<\/p:spTree>/, '');
}

function geometryEdits(source: SlideShapeData, edited: SlideShapeData): PropEdit[] {
  const before = getShapeBounds(source);
  const after = getShapeBounds(edited);
  const edits: PropEdit[] = [];
  if (before && after) {
    const pairs = [
      ['x', before.x, after.x],
      ['y', before.y, after.y],
      ['width', before.w, after.w],
      ['height', before.h, after.h],
    ] as const;
    for (const [prop, from, to] of pairs) if (from !== to) edits.push({ prop, from, to });
  }
  const fromAngle = Math.round(getShapeRotation(source) * ANGLE_UNITS);
  const toAngle = Math.round(getShapeRotation(edited) * ANGLE_UNITS);
  if (fromAngle !== toAngle) edits.push({ prop: 'rotation', from: fromAngle, to: toAngle });
  return edits;
}

/**
 * What the editor changed relative to the deck built from source. Shapes are
 * matched by slide index and shape id — the editor edits the deck the server
 * built, so ids agree. A change this diff cannot state as typed props is
 * reported as `unsupported`, never dropped.
 */
export function diffDecks(source: PresentationData, edited: PresentationData): DeckChange[] {
  const changes: DeckChange[] = [];
  const sourceSlides = getSlides(source);
  const editedSlides = getSlides(edited);
  if (sourceSlides.length !== editedSlides.length)
    changes.push({
      kind: 'unsupported',
      slide: null,
      shapeId: null,
      description: `slide count ${sourceSlides.length} → ${editedSlides.length}`,
    });
  const count = Math.min(sourceSlides.length, editedSlides.length);
  for (let slide = 0; slide < count; slide++) {
    const sourceSlide = sourceSlides[slide]!;
    const editedSlide = editedSlides[slide]!;
    if (
      withoutShapes(getSlideXmlString(sourceSlide)) !==
      withoutShapes(getSlideXmlString(editedSlide))
    )
      changes.push({ kind: 'unsupported', slide, shapeId: null, description: 'slide properties' });
    const before = shapesById(getSlideShapes(sourceSlide));
    const after = shapesById(getSlideShapes(editedSlide));
    for (const [shapeId, { shape, grouped }] of after) {
      const original = before.get(shapeId)?.shape;
      if (!original) {
        changes.push({ kind: 'unsupported', slide, shapeId, description: 'shape added' });
        continue;
      }
      // A group's XML holds its members, which are compared on their own.
      if (getShapeKind(shape) === 'group') {
        if (geometryEdits(original, shape).length)
          changes.push({ kind: 'unsupported', slide, shapeId, description: 'group transform' });
        continue;
      }
      if (
        withoutTransform(getShapeXmlString(original)) !== withoutTransform(getShapeXmlString(shape))
      )
        changes.push({
          kind: 'unsupported',
          slide,
          shapeId,
          description: 'properties other than position, size and rotation',
        });
      const edits = geometryEdits(original, shape);
      if (edits.length && grouped)
        changes.push({
          kind: 'unsupported',
          slide,
          shapeId,
          description: 'grouped shape geometry',
        });
      else if (edits.length) changes.push({ kind: 'geometry', slide, shapeId, edits });
    }
    for (const shapeId of before.keys())
      if (!after.has(shapeId))
        changes.push({ kind: 'unsupported', slide, shapeId, description: 'shape removed' });
  }
  return changes;
}

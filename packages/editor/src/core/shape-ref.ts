// References to shapes that outlive the editor's in-memory objects.
//
// Undo, Redo and a failed edit's rollback replace the whole presentation with
// one loaded from bytes, so a `SlideShapeData` held across edits may belong to
// a discarded object graph. A ref names a shape by what the file itself keeps
// stable instead: the slide's part name and the shape's `cNvPr` id.

import {
  findShapeById,
  findSlideByPartName,
  getShapeId,
  getShapeName,
  getShapeSlide,
  getSlidePartName,
  getSlideShapes,
  getSlides,
  type PresentationData,
  type SlideShapeData,
} from '@office-kit/pptx';
import { selectedShapeIds, type Selection } from './selection.ts';

/** A shape on a slide, identified independently of the editor's live objects. */
export interface ShapeRef {
  /** The slide's 0-based position when the ref was taken. Informational: slides can move. */
  readonly slideIndex: number;
  /** The slide's package part name (`/ppt/slides/slide3.xml`), which identifies the slide. */
  readonly slide: string;
  /** The shape's id (`getShapeId`), unique within its slide. */
  readonly shapeId: number;
  /** The shape's name when the ref was taken, as in the Selection Pane. */
  readonly name: string;
}

/** The ref for `shape`, which is on a slide of `presentation`. */
export function shapeRef(presentation: PresentationData, shape: SlideShapeData): ShapeRef {
  const slide = getShapeSlide(shape);
  return {
    slideIndex: getSlides(presentation).indexOf(slide),
    slide: getSlidePartName(slide),
    shapeId: getShapeId(shape),
    name: getShapeName(shape),
  };
}

/** Refs for the shapes `selection` names (the table, for a cell selection). */
export function selectionShapeRefs(
  presentation: PresentationData,
  selection: Selection,
): ShapeRef[] {
  const slide = getSlides(presentation)[selection.slideIndex];
  if (!slide) return [];
  const partName = getSlidePartName(slide);
  return selectedShapeIds(selection).flatMap((shapeId) => {
    // The selection is UI state, and not every edit that removes a shape
    // updates it; a shape that is gone is simply not selected.
    const shape = findShapeById(slide, shapeId);
    return shape
      ? [{ slideIndex: selection.slideIndex, slide: partName, shapeId, name: getShapeName(shape) }]
      : [];
  });
}

/**
 * The shape `ref` names in `presentation`. Throws when its slide or the shape
 * no longer exists, so an agent acting on an old selection fails instead of
 * editing something else.
 */
export function resolveShape(presentation: PresentationData, ref: ShapeRef): SlideShapeData {
  const slide = findSlideByPartName(presentation, ref.slide);
  if (!slide)
    throw new Error(`Shape "${ref.name}" (id ${ref.shapeId}) is gone: its slide was deleted.`);
  const shape = findShapeById(slide, ref.shapeId);
  if (!shape)
    throw new Error(
      `Shape "${ref.name}" (id ${ref.shapeId}) is gone: it was deleted from slide ${ref.slide}.`,
    );
  return shape;
}

/**
 * `selection` without the shapes and slides an edit removed. Selections are
 * kept by index and id, so after an agent's edit they may name things that are
 * gone, while the canvas, panes and ribbon expect them to exist.
 */
export function reconcileSelection(
  presentation: PresentationData,
  selection: Selection,
): Selection {
  const slides = getSlides(presentation);
  const slide = slides[selection.slideIndex];
  if (!slide) return { kind: 'none', slideIndex: Math.max(0, slides.length - 1) };
  if (selection.kind === 'none') return selection;
  if (selection.kind === 'slide') {
    const indices = selection.slideIndices ?? [selection.slideIndex];
    return indices.every((index) => index < slides.length)
      ? selection
      : { kind: 'slide', slideIndex: selection.slideIndex };
  }
  const live = new Set(getSlideShapes(slide).map(getShapeId));
  const none: Selection = { kind: 'none', slideIndex: selection.slideIndex };
  if (selection.kind === 'cell') return live.has(selection.shapeId) ? selection : none;
  const shapeIds = selection.shapeIds.filter((id) => live.has(id));
  if (shapeIds.length === selection.shapeIds.length) return selection;
  return shapeIds.length > 0 ? { ...selection, shapeIds } : none;
}

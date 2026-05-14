// Internal-only symbols shared between the class-based API (in
// `presentation.ts`, `slide.ts`, etc.) and the tree-shakeable free-function
// API (in `fn.ts`). Both APIs read and write the same opaque internal state
// on a `Presentation` / `Slide` / `SlideShape` value, so they need to agree
// on the symbol keys.
//
// Keeping the symbols in a class-free module is critical: a module that
// imports an identifier from this file does NOT transitively pull in any
// class definition. That's what lets the free-function API tree-shake the
// class entries out of consumer bundles.

import type { PartName } from '../internal/opc/index.ts';
import type { OpcPackage } from '../internal/parts/index.ts';
import type {
  CommentAuthor,
  SlideComment,
  SlideLayoutPart,
  SlidePart,
} from '../internal/presentationml/index.ts';
import type { XmlDocument, XmlElement } from '../internal/xml/index.ts';

export const INTERNAL_PACKAGE = Symbol('pptx-kit.package');
export const SLIDE_PART_NAME = Symbol('pptx-kit.slide.partName');
export const SLIDE_DOCUMENT = Symbol('pptx-kit.slide.document');
export const SLIDE_PART = Symbol('pptx-kit.slide.part');
export const SLIDE_SHAPES = Symbol('pptx-kit.slide.shapes');
export const SHAPE_SLIDE = Symbol('pptx-kit.shape.slide');
export const SHAPE_ELEMENT = Symbol('pptx-kit.shape.element');
export const SHAPE_SNAPSHOT = Symbol('pptx-kit.shape.snapshot');
export const LAYOUT_PART_NAME = Symbol('pptx-kit.layout.partName');
export const LAYOUT_PART = Symbol('pptx-kit.layout.part');
export const COMMENT_SLIDE = Symbol('pptx-kit.comment.slide');
export const COMMENT_SNAPSHOT = Symbol('pptx-kit.comment.snapshot');

/**
 * Data shape backing every `Presentation` value. The class implements
 * this interface; the free-function API constructs plain objects that
 * satisfy it. Both representations are interoperable as inputs to any
 * function that accepts `PresentationData`.
 */
export interface PresentationData {
  readonly [INTERNAL_PACKAGE]: OpcPackage;
  /** @internal — populated lazily by `getSlides` / `pres.slides`. */
  _slidesCache: unknown[] | null;
}

/** Data shape backing every `Slide` value. */
export interface SlideData {
  readonly [INTERNAL_PACKAGE]: OpcPackage;
  readonly [SLIDE_PART_NAME]: PartName;
  [SLIDE_DOCUMENT]: XmlDocument;
  [SLIDE_PART]: SlidePart;
  [SLIDE_SHAPES]: SlideShapeData[];
}

/** Data shape backing every `SlideShape` value. */
export interface SlideShapeData {
  readonly [SHAPE_SLIDE]: SlideData;
  [SHAPE_ELEMENT]: XmlElement;
  [SHAPE_SNAPSHOT]: SlidePart['shapes'][number];
}

/** Data shape backing every `SlideLayout` value. */
export interface SlideLayoutData {
  readonly [LAYOUT_PART_NAME]: PartName;
  readonly [LAYOUT_PART]: SlideLayoutPart;
}

/**
 * Opaque handle for one comment on a slide. The `author` is resolved
 * on read against the package-level `commentAuthors.xml`.
 */
export interface SlideCommentData {
  readonly [COMMENT_SLIDE]: SlideData;
  readonly [COMMENT_SNAPSHOT]: SlideComment;
  readonly author: CommentAuthor;
}

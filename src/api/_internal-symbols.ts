// Internal-only symbol keys for the opaque handles the public fn API
// returns (`PresentationData`, `SlideData`, `SlideShapeData`, etc.).
// Defined in a dedicated module so importers never transitively pull
// in heavy authoring code paths — readers stay light.
//
// These use `Symbol.for` (the process-global registry), NOT plain `Symbol`,
// so the keys are identical across separately-bundled copies of this module.
// The public entry (`pptx-kit` → dist/index.js) and the Node entry
// (`pptx-kit/node` → dist/node.js) are distinct bundles, each with its own
// copy of this file; companion packages (e.g. `@pptx-kit/preview`) hold a
// third. With plain `Symbol` a handle built by one bundle is opaque to every
// other — `getSlides(loadPresentationFile(...))` would read `undefined` and
// crash. The `pptx-kit.` namespace keeps the registry keys collision-free.

import type { PartName } from '../internal/opc/index.ts';
import type { OpcPackage } from '../internal/parts/index.ts';
import type {
  CommentAuthor,
  SlideComment,
  SlideLayoutPart,
  SlidePart,
} from '../internal/presentationml/index.ts';
import type { XmlDocument, XmlElement } from '../internal/xml/index.ts';

export const INTERNAL_PACKAGE = Symbol.for('pptx-kit.package');
export const SLIDE_PART_NAME = Symbol.for('pptx-kit.slide.partName');
export const SLIDE_DOCUMENT = Symbol.for('pptx-kit.slide.document');
export const SLIDE_PART = Symbol.for('pptx-kit.slide.part');
export const SLIDE_SHAPES = Symbol.for('pptx-kit.slide.shapes');
export const SHAPE_SLIDE = Symbol.for('pptx-kit.shape.slide');
export const SHAPE_ELEMENT = Symbol.for('pptx-kit.shape.element');
export const SHAPE_SNAPSHOT = Symbol.for('pptx-kit.shape.snapshot');
export const LAYOUT_PART_NAME = Symbol.for('pptx-kit.layout.partName');
export const LAYOUT_PART = Symbol.for('pptx-kit.layout.part');
export const COMMENT_SLIDE = Symbol.for('pptx-kit.comment.slide');
export const COMMENT_SNAPSHOT = Symbol.for('pptx-kit.comment.snapshot');
export const CELL_TABLE = Symbol.for('pptx-kit.cell.table');
export const CELL_ELEMENT = Symbol.for('pptx-kit.cell.element');
export const CELL_ROW = Symbol.for('pptx-kit.cell.row');
export const CELL_COL = Symbol.for('pptx-kit.cell.col');

/**
 * Opaque handle to a loaded / created presentation. Constructed via
 * `loadPresentation` / `createPresentation`; passed positionally to
 * every fn-API helper.
 */
export interface PresentationData {
  readonly [INTERNAL_PACKAGE]: OpcPackage;
  /** @internal — populated lazily by `getSlides`. */
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

/**
 * Opaque handle for one cell of a table graphic-frame shape. Carries
 * its zero-based row/column position plus the parent table shape so
 * mutations can commit back through the standard slide path.
 */
export interface TableCellData {
  readonly [CELL_TABLE]: SlideShapeData;
  readonly [CELL_ELEMENT]: XmlElement;
  readonly [CELL_ROW]: number;
  readonly [CELL_COL]: number;
}

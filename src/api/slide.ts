// Public `Slide` value. Read-only at this phase — authoring lands when the
// drawingml + parts layers grow real width.

import type { ShapeKind, SlidePart } from '../internal/presentationml/index.ts';
import { slideText } from '../internal/presentationml/index.ts';

/** A single shape on a slide as exposed to user code. */
export interface SlideShape {
  /**
   * Discriminator: which kind of shape this is. `'shape'` is the catch-all
   * for `<p:sp>`, which covers text boxes, rectangles, ellipses, and most
   * preset geometries.
   */
  readonly kind: ShapeKind;
  /** OOXML internal numeric id; unique within the slide. */
  readonly id: number;
  /** Human-readable name (PowerPoint shows this in the selection pane). */
  readonly name: string;
  /** Placeholder type if any (`'title'`, `'body'`, ...); `null` otherwise. */
  readonly placeholderType: string | null;
  /** All text on the shape, with `\n` between paragraphs. `''` if none. */
  readonly text: string;
}

export class Slide {
  /** @internal */
  readonly _part: SlidePart;

  /** @internal */
  constructor(part: SlidePart) {
    this._part = part;
  }

  /** Iterates the shapes on this slide in document order (groups flattened). */
  get shapes(): ReadonlyArray<SlideShape> {
    return this._part.shapes.map(
      ({ kind, id, name, placeholderType, text }): SlideShape => ({
        kind,
        id,
        name,
        placeholderType,
        text,
      }),
    );
  }

  /** Concatenated visible text from every shape on the slide. */
  get text(): string {
    return slideText(this._part);
  }
}

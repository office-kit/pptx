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

import type { OpcPackage } from '../internal/parts/index.ts';

export const INTERNAL_PACKAGE = Symbol('pptx-kit.package');

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

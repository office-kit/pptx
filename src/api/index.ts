// Public API surface. The only directory users are intended to import from.
//
// Two parallel APIs are exposed:
//
//   1. The class-based API (`Presentation`, `Slide`, `SlideShape`,
//      `SlideLayout`) — convenient, fluent, and what existing tests use.
//   2. The tree-shakeable free-function API (`loadPresentation`,
//      `savePresentation`, `createPresentation`, ...) — importing only
//      what you use lets modern bundlers drop the class definitions. See
//      `test/tree-shake.test.ts` for the CI guard.
//
// Both APIs operate on the same opaque internal state, so values produced
// by one can flow into the other.

export { type Emu, cm, emu, inches, mm, pt } from './units.ts';

// Class-based legacy API.
export { Presentation, type PresentationInput, _internalPackageOf } from './presentation.ts';
export { type PlaceholderType, Slide, SlideShape } from './slide.ts';
export { SlideLayout } from './slide-layout.ts';

// Tree-shakeable free-function API.
export type {
  PresentationData,
  SlideData,
  SlideLayoutData,
  SlideShapeData,
} from './_internal-symbols.ts';
export {
  addSlide,
  createPresentation,
  duplicateSlide,
  findSlidePlaceholder,
  getShapeFlip,
  getShapeId,
  getShapeKind,
  getShapeName,
  getShapePlaceholderIdx,
  getShapePlaceholderType,
  getShapePosition,
  getShapeRotation,
  getShapeSize,
  getShapeText,
  getSlideLayout,
  getSlideLayouts,
  getSlideShapes,
  getSlideText,
  getSlides,
  loadPresentation,
  moveSlide,
  removeSlide,
  replaceTokensInPresentation,
  replaceTokensInSlide,
  savePresentation,
} from './fn.ts';

export type { BulletStyle, ParagraphAlignment, TextFormat } from '../internal/drawingml/index.ts';
export type {
  PresetShape,
  TransitionEffect,
  TransitionOptions,
} from '../internal/presentationml/index.ts';

// Library version. Replaced at build time by the package version.
export const VERSION = '0.0.0';

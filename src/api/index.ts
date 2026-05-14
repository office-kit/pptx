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
  SlideCommentData,
  SlideData,
  SlideLayoutData,
  SlideShapeData,
} from './_internal-symbols.ts';
export type {
  CommentAuthor,
  CommentPosition,
  SlideComment,
} from '../internal/presentationml/index.ts';
export type { SlideSize } from './fn.ts';
export { SLIDE_SIZE_4_3, SLIDE_SIZE_16_9, SLIDE_SIZE_16_10 } from './fn.ts';
export type { ChartKind, ChartSeries, ChartSpec } from '../internal/chartml/index.ts';
export type { ShapeClickAction } from './fn.ts';
export type { IssueSeverity, ValidationIssue } from './fn.ts';
export type { AnimationEffect, AnimationOptions } from './fn.ts';
export type { ImageCrop } from './fn.ts';
export type {
  GradientFillOptions,
  GradientStop,
  PatternFillOptions,
  PatternPreset,
} from '../internal/drawingml/index.ts';
export type { ShapeFill, SlideBackground, TextAnchor } from './fn.ts';
export {
  addSlide,
  addSlideChart,
  bringShapeForward,
  bringShapeToFront,
  addSlideComment,
  addSlideImage,
  addSlideLine,
  addSlideShape,
  addSlideTable,
  addSlideTextBox,
  clearShapeFill,
  clearShapeStroke,
  clearSlideAnimations,
  clearSlideBackground,
  clearSlideTransition,
  createPresentation,
  duplicateSlide,
  findShapeByName,
  findShapeInPresentation,
  findShapesByKind,
  findShapesByName,
  findSlideLayout,
  findSlidePlaceholder,
  getCommentAuthor,
  getCommentAuthors,
  getCommentDate,
  getCommentPosition,
  getCommentText,
  getShapeFill,
  getShapeFlip,
  getShapeId,
  getShapeKind,
  getShapeName,
  getShapeParagraphCount,
  getShapePlaceholderIdx,
  getShapePlaceholderType,
  getShapePosition,
  getShapeRotation,
  getShapeRunCount,
  getShapeRunText,
  getShapeSize,
  getShapeText,
  getSlideComments,
  getSlideLayout,
  getSlideSize,
  getSlideLayoutName,
  getSlideLayouts,
  getSlideLayoutType,
  getSlideNotes,
  getSlideBackground,
  getSlideShapes,
  getSlideText,
  getSlideTitle,
  getSlideTransition,
  getSlides,
  loadPresentation,
  moveSlide,
  removeShape,
  removeSlide,
  removeSlideComment,
  replaceTokensInPresentation,
  replaceTokensInSlide,
  savePresentation,
  sendShapeBackward,
  sendShapeToBack,
  setShapeAlignment,
  setShapeAnimation,
  setShapeBullets,
  setShapeFill,
  setShapeFlip,
  setShapeGradientFill,
  setShapeHyperlink,
  setShapeImage,
  setShapeImageCrop,
  setShapeImageFill,
  setShapeImageOpacity,
  setShapeClickAction,
  setShapeNoFill,
  setShapeNoStroke,
  setShapePatternFill,
  setParagraphAlignment,
  setParagraphBullet,
  setParagraphLevel,
  setShapePosition,
  setShapeRotation,
  setShapeRunFormat,
  setShapeRunText,
  setShapeSize,
  setShapeStroke,
  setShapeText,
  setShapeTextAnchor,
  setShapeTextFormat,
  setSlideBackground,
  setSlideNotes,
  setSlideSize,
  setSlideTitle,
  setSlideTransition,
  validatePresentation,
} from './fn.ts';

export type { BulletStyle, ParagraphAlignment, TextFormat } from '../internal/drawingml/index.ts';
export type {
  PresetShape,
  TransitionEffect,
  TransitionOptions,
} from '../internal/presentationml/index.ts';

// Library version. Replaced at build time by the package version.
export const VERSION = '0.0.0';

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
  TableCellData,
} from './_internal-symbols.ts';
export type {
  CommentAuthor,
  CommentPosition,
  SlideComment,
} from '../internal/presentationml/index.ts';
export type { SlideSize } from './fn.ts';
export { SLIDE_SIZE_4_3, SLIDE_SIZE_16_9, SLIDE_SIZE_16_10 } from './fn.ts';
export type { ChartKind, ChartSeries, ChartSpec } from '../internal/chartml/index.ts';
export type { SlideChartData } from './fn.ts';
export type { ShapeClickAction } from './fn.ts';
export type { IssueSeverity, ValidationIssue } from './fn.ts';
export type { AnimationEffect, AnimationOptions } from './fn.ts';
export type { ImageCrop } from './fn.ts';
export type {
  ArrowOptions,
  GlowOptions,
  GradientFillOptions,
  GradientStop,
  LineDash,
  LineEndSize,
  LineEndType,
  PatternFillOptions,
  PatternPreset,
  ShadowOptions,
} from '../internal/drawingml/index.ts';
export type { ShapeBounds, ShapeFill, ShapeStroke, SlideBackground, TextAnchor } from './fn.ts';
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
  clearShapeEffects,
  clearShapeFill,
  clearShapeStroke,
  clearSlideAnimations,
  clearSlideBackground,
  clearSlideShapes,
  clearSlideTransition,
  clearTableCellFill,
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
  getShapeAnimation,
  getShapeBounds,
  getShapeFill,
  getShapeFlip,
  getShapeId,
  getShapeImageCrop,
  getShapeImageOpacity,
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
  getShapeStroke,
  getShapeText,
  getSlideComments,
  getSlideLayout,
  getSlideSize,
  getSlideAt,
  getSlideBackground,
  getSlideCharts,
  getSlideIndex,
  getSlideLayoutName,
  getSlideLayouts,
  getSlideLayoutType,
  getSlideNotes,
  getSlideShapes,
  getSlideText,
  getSlideTitle,
  getSlideTransition,
  getSlides,
  getTableCell,
  getTableCellPosition,
  getTableCellText,
  getTableCells,
  isSlideHidden,
  loadPresentation,
  moveSlide,
  removeShape,
  removeSlide,
  removeSlideComment,
  replaceTextInPresentation,
  replaceTextInSlide,
  replaceTokensInPresentation,
  replaceTokensInSlide,
  savePresentation,
  sendShapeBackward,
  sendShapeToBack,
  setChartSpec,
  setShapeAlignment,
  setShapeAnimation,
  setShapeBounds,
  setShapeBullets,
  setShapeFill,
  setShapeFlip,
  setShapeGlow,
  setShapeGradientFill,
  setShapeShadow,
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
  setShapeStrokeArrow,
  setShapeStrokeDash,
  setShapeText,
  setShapeTextAnchor,
  setShapeTextFormat,
  setShapeTextMargins,
  setSlideBackground,
  setSlideHidden,
  setSlideLayout,
  setSlideNotes,
  setSlideSize,
  setSlideTitle,
  setSlideTransition,
  setTableCellAlignment,
  setTableCellFill,
  setTableCellText,
  setTableCellTextFormat,
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

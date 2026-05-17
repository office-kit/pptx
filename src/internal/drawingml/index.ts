// internal/drawingml — a: namespace: shapes, text, geometry, color, effects.
// Allowed imports: internal/xml.

export { paragraphText, paragraphsOf, runsOf, textBodyText } from './text-body.ts';
export {
  applyAlignmentToAllParagraphs,
  applyBulletToAllParagraphs,
  applyBulletToParagraph,
  type BulletStyle,
  type ParagraphAlignment,
  replaceTextInTree,
  replaceTokensInTree,
  setTextBody,
} from './text-body-mutation.ts';
export type { TextFormat } from './text-format.ts';
export { applyFormatToAllRuns, applyRunFormat } from './text-format.ts';
export type { ParsedColor } from './color.ts';
export { buildColorElement, parseColor } from './color.ts';
export type {
  GradientFillOptions,
  GradientStop,
  PatternFillOptions,
  PatternPreset,
} from './fill.ts';
export { clearFill, setGradientFill, setNoFill, setPatternFill, setSolidFill } from './fill.ts';
export type { GlowOptions, ShadowOptions } from './effects.ts';
export { clearEffects, setGlow, setShadow } from './effects.ts';
export type { ArrowOptions, LineDash, LineEndSize, LineEndType, StrokeOptions } from './stroke.ts';
export {
  clearStroke,
  setNoStroke,
  setSolidStroke,
  setStrokeArrow,
  setStrokeDash,
} from './stroke.ts';
export { applyHyperlinkToAllRuns } from './hyperlink.ts';
export { getPictureEmbedRId } from './picture-mutation.ts';
export type { Position, ShapeKindForGeometry, Size } from './geometry.ts';
export { readFlip, readPosition, readRotation, readSize } from './geometry.ts';
export { setFlip, setPosition, setRotation, setSize } from './geometry-mutation.ts';

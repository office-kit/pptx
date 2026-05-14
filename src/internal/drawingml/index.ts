// internal/drawingml — a: namespace: shapes, text, geometry, color, effects.
// Allowed imports: internal/xml.

export { paragraphText, paragraphsOf, runsOf, textBodyText } from './text-body.ts';
export {
  applyBulletToAllParagraphs,
  type BulletStyle,
  replaceTokensInTree,
  setTextBody,
} from './text-body-mutation.ts';
export type { TextFormat } from './text-format.ts';
export { applyFormatToAllRuns, applyRunFormat } from './text-format.ts';
export type { ParsedColor } from './color.ts';
export { buildColorElement, parseColor } from './color.ts';
export { clearFill, setNoFill, setSolidFill } from './fill.ts';
export type { StrokeOptions } from './stroke.ts';
export { clearStroke, setNoStroke, setSolidStroke } from './stroke.ts';
export { getPictureEmbedRId } from './picture-mutation.ts';
export type { Position, ShapeKindForGeometry, Size } from './geometry.ts';
export { readPosition, readSize } from './geometry.ts';
export { setPosition, setSize } from './geometry-mutation.ts';

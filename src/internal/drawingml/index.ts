// internal/drawingml — a: namespace: shapes, text, geometry, color, effects.
// Allowed imports: internal/xml.

export { paragraphText, paragraphsOf, runsOf, textBodyText } from './text-body.ts';
export { replaceTokensInTree, setTextBody } from './text-body-mutation.ts';
export { getPictureEmbedRId } from './picture-mutation.ts';
export type { Position, ShapeKindForGeometry, Size } from './geometry.ts';
export { readPosition, readSize } from './geometry.ts';

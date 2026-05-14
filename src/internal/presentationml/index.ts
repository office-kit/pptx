// internal/presentationml — p: namespace: presentation/slide/layout/master.
// Allowed imports: internal/drawingml, internal/parts, internal/xml.

export { REL_TYPES, type RelType } from './relationship-types.ts';
export type {
  NotesMasterId,
  PresentationPart,
  SlideId,
  SlideMasterId,
  SlideSize,
} from './presentation-part.ts';
export { readPresentationPart } from './presentation-part.ts';
export type { ShapeKind, SlidePart, SlideShape } from './slide-part.ts';
export { readShapeTreeFromCsldRoot, readSlidePart, slideText } from './slide-part.ts';
export type { SlideLayoutPart, SlideLayoutType } from './slide-layout-part.ts';
export { readSlideLayoutPart } from './slide-layout-part.ts';

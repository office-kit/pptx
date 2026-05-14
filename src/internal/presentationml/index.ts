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
export { buildSlideFromLayout } from './slide-builder.ts';
export type { TextBoxOptions } from './text-box-builder.ts';
export { buildTextBox } from './text-box-builder.ts';
export type { PresetShape, ShapeOptions } from './shape-builder.ts';
export { buildShape } from './shape-builder.ts';
export type { ConnectorOptions } from './connector-builder.ts';
export { buildConnector } from './connector-builder.ts';
export type { PictureOptions } from './picture-builder.ts';
export { buildPicture } from './picture-builder.ts';
export type { TableOptions } from './table-builder.ts';
export { buildTable } from './table-builder.ts';
export { buildEmptyNotesSlide } from './notes-slide-builder.ts';

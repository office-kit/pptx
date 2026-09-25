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
export type { PlaceholderType, ShapeKind, SlidePart, SlideShape } from './slide-part.ts';
export {
  readGroupChildren,
  readShapeTreeFromCsldRoot,
  readSlidePart,
  slideText,
} from './slide-part.ts';
export type { SlideLayoutPart, SlideLayoutType } from './slide-layout-part.ts';
export { readSlideLayoutPart } from './slide-layout-part.ts';
export { buildSlideFromLayout } from './slide-builder.ts';
export type { TextBoxOptions } from './text-box-builder.ts';
export { buildTextBox } from './text-box-builder.ts';
export type { PresetShape, ShapeOptions } from './shape-builder.ts';
export { buildShape } from './shape-builder.ts';
export type { ConnectorOptions } from './connector-builder.ts';
export { buildConnector } from './connector-builder.ts';
export type { GroupOptions } from './group-builder.ts';
export { buildGroup } from './group-builder.ts';
export type { PictureOptions } from './picture-builder.ts';
export { buildPicture } from './picture-builder.ts';
export type { MediaFileKind, MediaPictureOptions, PictureMediaRef } from './media-builder.ts';
export {
  buildMediaPicture,
  buildMediaTimingNode,
  buildTimingRoot,
  isMediaTimingNode,
  mediaTimingNodeTarget,
  readPictureMediaRef,
} from './media-builder.ts';
export type { TableOptions } from './table-builder.ts';
export { buildTable, buildTableCell, buildTableRow } from './table-builder.ts';
export { buildEmptyNotesSlide } from './notes-slide-builder.ts';
export type { SlideTransition, TransitionEffect, TransitionOptions } from './transition-builder.ts';
export { buildTransition } from './transition-builder.ts';
export type {
  AnimationDirection,
  AnimationEffect,
  AnimationOptions,
  AnimationStartCondition,
} from './animation-builder.ts';
export {
  ANIMATION_DIRECTIONS,
  ANIMATION_EFFECTS,
  buildSingleEffectTiming,
  FULL_TURN,
  isDirectionalEffect,
  trailingHideDelayMs,
} from './animation-builder.ts';
export type {
  CommentAuthor,
  CommentAuthorList,
  CommentList,
  CommentPosition,
  SlideComment,
} from './comments-part.ts';
export {
  DEFAULT_COMMENT_POSITION,
  buildCommentAuthorListDoc,
  buildCommentListDoc,
  readCommentAuthorList,
  readCommentList,
} from './comments-part.ts';
export type {
  CommentStatus,
  ModernAuthor,
  ModernComment,
  ModernCommentPosition,
  ModernReply,
} from './modern-comments-part.ts';
export {
  COMMENT_STATUSES,
  CREATION_ID_URI,
  MODERN_AUTHORS_CONTENT_TYPE,
  MODERN_COMMENTS_CONTENT_TYPE,
  MODERN_COMMENTS_NS,
  P14_NS,
  PC_NS,
  anchorOf,
  appendModernReply,
  buildModernAuthorElement,
  buildModernAuthorListDoc,
  buildModernCommentElement,
  buildModernCommentListDoc,
  buildModernReplyElement,
  buildSlideAnchor,
  buildUnknownAnchor,
  findModernComment,
  findModernReply,
  modernCommentElements,
  readModernAuthorList,
  readModernCommentList,
  removeModernComment,
  setModernStatus,
  setModernText,
} from './modern-comments-part.ts';

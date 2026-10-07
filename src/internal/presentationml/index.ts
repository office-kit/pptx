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
export {
  DEFAULT_TABLE_STYLE_ID,
  buildTable,
  buildTableCell,
  buildTableRow,
} from './table-builder.ts';
export type { BuiltinTableStyle, BuiltinTableStyleName } from './table-styles.ts';
export {
  BUILTIN_TABLE_STYLES,
  builtinTableStyleIdByName,
  builtinTableStyleXml,
  getBuiltinTableStyle,
} from './table-styles.ts';
export { buildEmptyNotesSlide } from './notes-slide-builder.ts';
export type {
  Box,
  HandoutMasterPlaceholderType,
  LayoutPlaceholderKind,
  MasterPlaceholderType,
  NotesMasterPlaceholderType,
} from './default-masters.ts';
export {
  HANDOUT_MASTER_PLACEHOLDER_TYPES,
  MASTER_PLACEHOLDER_TYPES,
  NOTES_MASTER_PLACEHOLDER_TYPES,
  customSlideLayoutXml,
  defaultHandoutMasterXml,
  defaultNotesMasterXml,
  defaultSlideLayoutsXml,
  defaultSlideMasterXml,
  handoutMasterPlaceholderXml,
  insertedPlaceholderXml,
  layoutFooterPlaceholdersXml,
  layoutTitlePlaceholderXml,
  masterPlaceholderXml,
  notesMasterPlaceholderXml,
} from './default-masters.ts';
export type {
  MorphOption,
  SlideTransition,
  TransitionEffect,
  TransitionOptions,
  TransitionPreset,
} from './transition-builder.ts';
export { buildTransition, transitionEffectNamespace } from './transition-builder.ts';
export type {
  AnimationDirection,
  AnimationEffect,
  AnimationEffectOptions,
  AnimationInOut,
  AnimationOptionDomains,
  AnimationOptionName,
  AnimationOptions,
  AnimationOrientation,
  AnimationShape,
  AnimationStartCondition,
  AnimationTextBuild,
  EffectContext,
} from './animation-builder.ts';
export type {
  AnimationColor,
  AnimationEmphasisOptionName,
  AnimationEmphasisOptions,
  AnimationScaleDirection,
  AnimationSpinDirection,
} from './animation-emphasis-options.ts';
export {
  ANIMATION_EMPHASIS_OPTION_NAMES,
  animationEmphasisOptionNames,
  readAnimationEmphasisOptions,
  resolveAnimationEmphasisOptions,
} from './animation-emphasis-options.ts';
export {
  ANIMATION_DIRECTIONS,
  ANIMATION_EFFECTS,
  ANIMATION_PRESET_ENTRIES,
  animationBuildKind,
  animationOptionDomains,
  buildSingleEffectTiming,
  defaultAnimationDurationMs,
  effectBehaviourNames,
  effectBehaviours,
  effectDurationMs,
  lastBehaviourEndMs,
  setEffectDurationMs,
  isDirectionalEffect,
  resolveAnimationOptions,
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

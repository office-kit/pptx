import { textCaseEdits, type TextCase } from './text-case.ts';
import { keepShowInk, type ShowInkAnnotation } from './show-ink-save.ts';
import { fitTableText } from './table-text-fit.ts';
import { applyTableBorders, type TableBorderMode } from './table-borders.ts';
import { findText, type TextSearchOptions } from './text-search.ts';
import { connectorEndpoints, connectorFrame, shapePoint } from './connector-geometry.ts';
import { shapePresets, lineTool, defaultButtonAction } from './shape-gallery.ts';
import { createHash, randomUUID } from 'node:crypto';
import { readFile, writeFile, rename, rm } from 'node:fs/promises';
import * as pptx from '@office-kit/pptx';

export type EditorLinkAction =
  | { kind: 'url'; url: string }
  | { kind: 'slide'; slide: string }
  | { kind: 'customShow'; id: number; showAndReturn: boolean }
  | {
      kind: 'nextSlide' | 'prevSlide' | 'firstSlide' | 'lastSlide' | 'lastSlideViewed' | 'endShow';
    };
export interface EditorLink {
  action: EditorLinkAction;
  tooltip: string | null;
}
export interface EditorActionSound {
  sound: { name: string; base64: string } | null;
  stopPrevious: boolean;
}
export type EditorActionSounds = Record<'click' | 'hover', EditorActionSound>;
export type EditorTransitionSound =
  | { kind: 'none' | 'stop' }
  | { kind: 'play'; name: string; base64: string; loop: boolean };
function transitionSound(slide: pptx.SlideData): EditorTransitionSound {
  const value = pptx.getSlideTransitionSound(slide);
  return value.kind === 'play'
    ? {
        kind: 'play',
        name: value.name,
        base64: Buffer.from(value.bytes).toString('base64'),
        loop: value.loop,
      }
    : value;
}
function actionSounds(shape: pptx.SlideShapeData): EditorActionSounds {
  const read = (trigger: pptx.ShapeActionTrigger): EditorActionSound => {
    const value = pptx.getShapeActionSound(shape, trigger);
    return {
      stopPrevious: value.stopPrevious,
      sound: value.sound
        ? { name: value.sound.name, base64: Buffer.from(value.sound.bytes).toString('base64') }
        : null,
    };
  };
  return { click: read('click'), hover: read('hover') };
}
function resolveActionSound(value: EditorActionSound): pptx.ShapeActionSound {
  if (!value || typeof value.stopPrevious !== 'boolean' || value.sound === undefined)
    throw new Error('Invalid action sound.');
  if (value.sound === null) return { sound: null, stopPrevious: value.stopPrevious };
  text(value.sound.name);
  const encoded = value.sound.base64;
  if (
    typeof encoded !== 'string' ||
    encoded.length > 28_000_000 ||
    encoded.length % 4 !== 0 ||
    !/^[A-Za-z0-9+/]*={0,2}$/.test(encoded)
  )
    throw new Error('Invalid action sound data.');
  const bytes = Buffer.from(encoded, 'base64');
  if (
    bytes.length < 12 ||
    bytes.length > 20_000_000 ||
    bytes.toString('ascii', 0, 4) !== 'RIFF' ||
    bytes.toString('ascii', 8, 12) !== 'WAVE'
  )
    throw new Error('Choose a WAV sound under 20 MB.');
  return { sound: { name: value.sound.name, bytes }, stopPrevious: value.stopPrevious };
}
function objectLink(shape: pptx.SlideShapeData, hover = false): EditorLink | null {
  const action = hover ? pptx.getShapeHoverAction(shape) : pptx.getShapeClickAction(shape);
  if (!action) return null;
  return {
    action:
      action.kind === 'slide'
        ? { kind: 'slide', slide: pptx.getSlidePartName(action.slide) }
        : action,
    tooltip: hover
      ? pptx.getShapeHoverActionTooltip(shape)
      : pptx.getShapeClickActionTooltip(shape),
  };
}
function resolveActionLink(
  presentation: pptx.PresentationData,
  link: EditorLink | null | undefined,
): pptx.ShapeClickAction | null {
  if (link === undefined) throw new Error('Invalid link selection.');
  let action: pptx.ShapeClickAction | null = null;
  if (link !== null) {
    if (!link || typeof link !== 'object' || !link.action) throw new Error('Invalid link.');
    if (link.tooltip !== null) text(link.tooltip);
    const requested = link.action;
    if (requested.kind === 'url') {
      text(requested.url);
      if (!requested.url.trim()) throw new Error('Enter a link address.');
      action = { kind: 'url', url: requested.url };
    } else if (requested.kind === 'customShow') {
      if (
        !Number.isInteger(requested.id) ||
        typeof requested.showAndReturn !== 'boolean' ||
        !pptx.getCustomShows(presentation).some((show) => show.id === requested.id)
      )
        throw new Error('Custom show destination no longer exists.');
      action = { kind: 'customShow', id: requested.id, showAndReturn: requested.showAndReturn };
    } else if (requested.kind === 'slide') {
      text(requested.slide);
      const destination = pptx
        .getSlides(presentation)
        .find((s) => pptx.getSlidePartName(s) === requested.slide);
      if (!destination) throw new Error('Link destination no longer exists.');
      action = { kind: 'slide', slide: destination };
    } else if (
      ['nextSlide', 'prevSlide', 'firstSlide', 'lastSlide', 'lastSlideViewed', 'endShow'].includes(
        requested.kind,
      )
    ) {
      action = { kind: requested.kind } as pptx.ShapeClickAction;
    } else throw new Error('Invalid link action.');
  }
  return action;
}

function textLinks(
  links: ReturnType<typeof pptx.getShapeTextRangeClickActions>,
): EditorShape['textLinks'] {
  return links.map(({ start, end, action, tooltip }) => ({
    start,
    end,
    link: {
      action:
        action.kind === 'slide'
          ? { kind: 'slide', slide: pptx.getSlidePartName(action.slide) }
          : action,
      tooltip,
    },
  }));
}
export interface EditorShape {
  /** Direct group children, in their parent coordinate system. */
  children?: EditorShape[];
  /** Parent coordinates to slide coordinates: [a, b, c, d, e, f]. */
  parentTransform?: [number, number, number, number, number, number] | null;
  link: EditorLink | null;
  hoverLink: EditorLink | null;
  actionSounds: EditorActionSounds;
  textLinks: Array<{ start: number; end: number; link: EditorLink }>;
  id: number;
  name: string;
  hidden?: boolean;
  regroupIds?: number[];
  kind: string;
  bounds: { x: number; y: number; w: number; h: number } | null;
  rotation: number;
  aspectRatioLocked: boolean;
  preset: string | null;
  imageCrop: pptx.ImageCrop | null;
  imageOpacity: number | null;
  imageBrightness: number | null;
  imageContrast: number | null;
  connections: { start: pptx.ShapeConnection | null; end: pptx.ShapeConnection | null };
  connectionSites: Array<{ x: number; y: number }>;
  flip: { horizontal: boolean; vertical: boolean } | null;
  text: string;
  textable: boolean;
  table?: {
    style: ReturnType<typeof pptx.getTableStyleFlags>;
    columnWidths: number[];
    rowHeights: number[];
    cells: Array<
      Array<{
        text: string;
        textLinks: EditorShape['textLinks'];
        span: ReturnType<typeof pptx.getTableCellSpan>;
        paragraphs: ReturnType<typeof pptx.getTableCellParagraphs>;
        margins: ReturnType<typeof pptx.getTableCellMargins>;
        anchor: ReturnType<typeof pptx.getTableCellAnchor>;
        direction: ReturnType<typeof pptx.getTableCellTextDirection>;
      }>
    >;
  };
  placeholder: string | null;
  format: pptx.TextFormat;
  runs: Array<{ start: number; end: number; format: pptx.TextFormat }>;
  paragraphs: Array<{
    start: number;
    end: number;
    align: string | null;
    bullets: pptx.BulletStyle | null;
    properties: pptx.ParagraphProperties;
  }>;
  fill: string | null;
  paint: {
    fill: pptx.ShapeFill;
    stroke: pptx.ShapeStroke;
    dash: pptx.LineDash | null;
    cap: ReturnType<typeof pptx.getShapeStrokeCap>;
    join: ReturnType<typeof pptx.getShapeStrokeJoin>;
    compound: ReturnType<typeof pptx.getShapeStrokeCompound>;
    headArrow: pptx.ArrowOptions | null;
    tailArrow: pptx.ArrowOptions | null;
    fillOpacity: number | null;
    strokeOpacity: number | null;
  } | null;
  align: string | null;
  bullets: pptx.BulletStyle | null;
  anchor: pptx.TextAnchor | null;
  anchorCenter: boolean;
  textDirection: NonNullable<ReturnType<typeof pptx.getShapeTextDirection>> | 'horz';
  autoFit: pptx.TextAutoFit;
  autoFitParams: ReturnType<typeof pptx.getShapeTextAutoFitParams>;
  textFrame: {
    margins: { left: number; top: number; right: number; bottom: number };
    anchor: pptx.TextAnchor;
    wrap: boolean;
    columns: { count: number; gapEmu?: number } | null;
  } | null;
}
export interface EditorSlide {
  key: string;
  token: string;
  layout: string | null;
  shapes: EditorShape[];
  notes: string;
  hidden: boolean;
  transition: pptx.TransitionOptions | null;
  transitionSound: EditorTransitionSound;
  animations: ReturnType<typeof pptx.getSlideAnimationSequence>;
}
export interface EditorModel {
  sizeType?: string;
  firstSlideNumber: number;
  notesOrientation: 'portrait' | 'landscape';
  width: number;
  height: number;
  gridSpacing: { x: number; y: number } | null;
  snapToGrid: boolean | null;
  guidesVisible: boolean;
  guides: Array<{ id: number; axis: 'x' | 'y'; offset: number; color: string }>;
  sections: Array<{ name: string; slides: string[] }>;
  customShows: Array<{ id: number; name: string; slides: string[] }>;
  showProperties: pptx.SlideShowProperties;
  layouts: Array<{
    key: string;
    name: string;
    type: string | null;
    placeholders: Array<{ bounds: EditorShape['bounds']; type: string | null }>;
  }>;
  slides: EditorSlide[];
}
export interface EditCommand {
  type:
    | 'ink'
    | 'table-borders'
    | 'table-margins'
    | 'table-fill'
    | 'table-style'
    | 'table-cell-size'
    | 'table-distribute'
    | 'table-row-append'
    | 'table-insert'
    | 'table-delete'
    | 'table-merge'
    | 'table-split'
    | 'table-clear'
    | 'table-range-format'
    | 'table-cell-layout'
    | 'table-cell-text'
    | 'table-cell-format'
    | 'table-cell-align'
    | 'table-cell-case'
    | 'table-cell-indent'
    | 'table-cell-paragraph'
    | 'table-cell-line-spacing'
    | 'table-cell-bullets'
    | 'text-replace-all'
    | 'slide-size'
    | 'custom-shows'
    | 'show-properties'
    | 'slide-advance'
    | 'animation-duration'
    | 'animation-delay'
    | 'animation-start'
    | 'animation-remove'
    | 'animation-move'
    | 'animation-reorder'
    | 'animation-add'
    | 'animation-effect'
    | 'animation-settings'
    | 'transition-effect'
    | 'transition-duration'
    | 'transition-sound'
    | 'transition-apply-all'
    | 'grid-spacing'
    | 'grid-snap'
    | 'guides-visible'
    | 'guides'
    | 'update'
    | 'delete'
    | 'duplicate'
    | 'paste'
    | 'insert'
    | 'picture-replace'
    | 'object-link'
    | 'object-actions'
    | 'text-link'
    | 'group'
    | 'ungroup'
    | 'regroup'
    | 'order'
    | 'slide-add'
    | 'slide-layout'
    | 'slide-reset'
    | 'slide-delete'
    | 'slide-duplicate'
    | 'slide-move'
    | 'slide-hidden'
    | 'notes'
    | 'section-add'
    | 'section-rename'
    | 'section-remove'
    | 'section-remove-all'
    | 'section-delete'
    | 'section-move';
  slide: number;
  ink?: ShowInkAnnotation[];
  link?: EditorLink | null;
  hoverLink?: EditorLink | null;
  actionSounds?: EditorActionSounds;
  range?: { start: number; end: number };
  tableBorders?: {
    mode: TableBorderMode;
    color: string;
    weight: number;
    dash: 'solid' | 'dash' | 'dot' | 'dashDot';
  };
  tableAxis?: 'rows' | 'columns';
  tableResizeAnchor?: 'top' | 'bottom';
  tablePlacement?: 'above' | 'below' | 'left' | 'right';
  tableCellSize?: number;
  tableSplit?: { rows: number; columns: number };
  tableRange?: { row: number; column: number; rows: number; columns: number };
  tableMargins?: Partial<Record<'left' | 'right' | 'top' | 'bottom', number>>;
  tableFill?: string | null;
  tableStyle?: Partial<ReturnType<typeof pptx.getTableStyleFlags>>;
  replacement?: { query: string; text: string; options?: TextSearchOptions };
  size?: {
    width: number;
    height: number;
    type?: string;
    firstSlideNumber?: number;
    notesOrientation?: 'portrait' | 'landscape';
  };
  scaleContent?: boolean;
  guides?: EditorModel['guides'];
  gridSpacing?: { x: number; y: number };
  snapToGrid?: boolean;
  guidesVisible?: boolean;
  ids?: number[];
  changes?: {
    name?: string;
    hidden?: boolean;
    connections?: { start?: pptx.ShapeConnection | null; end?: pptx.ShapeConnection | null };
    bounds?: EditorShape['bounds'];
    rotation?: number;
    rotationDelta?: number;
    flipToggle?: 'horizontal' | 'vertical';
    imageCropShape?: string;
    imageFit?: 'fill' | 'fit';
    imageCropAspectRatio?: number;
    imageCrop?: pptx.ImageCrop;
    imageOpacity?: number;
    imageBrightness?: number;
    imageContrast?: number;
    aspectRatioLocked?: boolean;
    flip?: { horizontal: boolean; vertical: boolean };
    text?: string;
    textEdits?: Array<{ start: number; end: number; text: string; format?: pptx.TextFormat }>;
    changeCase?: TextCase;
    format?: pptx.TextFormat;
    range?: { start: number; end: number };
    fill?: string;
    stroke?: { color?: string; widthEmu?: number } | 'none';
    strokeDash?: pptx.LineDash;
    strokeCap?: ReturnType<typeof pptx.getShapeStrokeCap>;
    strokeJoin?: ReturnType<typeof pptx.getShapeStrokeJoin>;
    strokeCompound?: ReturnType<typeof pptx.getShapeStrokeCompound>;
    strokeHeadArrow?: Partial<pptx.ArrowOptions>;
    strokeTailArrow?: Partial<pptx.ArrowOptions>;
    fillOpacity?: number;
    strokeOpacity?: number;
    align?: 'left' | 'center' | 'right' | 'justify' | 'distribute';
    bullets?: 'bullet' | 'number' | 'none';
    anchor?: pptx.TextAnchor;
    anchorCenter?: boolean;
    direction?: 'horz' | 'vert' | 'vert270' | 'wordArtVert';
    columns?: number;
    textColumns?: { count?: number; gapEmu?: number };
    margins?: Partial<NonNullable<EditorShape['textFrame']>['margins']>;
    wrap?: boolean;
    autoFit?: pptx.TextAutoFit;
    autoFitParams?: NonNullable<EditorShape['autoFitParams']>;
    lineSpacing?: number;
    paragraph?: pptx.ParagraphSettings;
    indent?: -1 | 1;
  };
  positions?: Array<{ id: number; bounds: NonNullable<EditorShape['bounds']> }>;
  textFits?: Array<{
    id: number;
    slide?: number;
    bounds?: NonNullable<EditorShape['bounds']>;
    autoFitParams?: NonNullable<EditorShape['autoFitParams']>;
  }>;
  customShows?: EditorModel['customShows'];
  showProperties?: pptx.SlideShowProperties;
  animation?: pptx.AnimationOptions;
  animationStart?: { id: string; start: pptx.AnimationStart };
  animationDelay?: { id: string; delayMs: number };
  animationId?: string;
  animationIds?: string[];
  animationBeforeId?: string | null;
  animationSettings?: pptx.AnimationSettings;
  animationDirection?: 'earlier' | 'later';
  animationTiming?: { id: string; durationMs: number };
  transitionDurationMs?: number;
  transitionSound?: EditorTransitionSound;
  transitionEffect?: Pick<
    pptx.TransitionOptions,
    'effect' | 'direction' | 'orientation' | 'thruBlack' | 'spokes'
  >;
  advanceTiming?: { advanceOnClick: boolean; advanceAfterMs: number | null };
  section?: number;
  layout?: string;
  source?: number;
  snapshot?: string;
  pasteSlideCoordinates?: boolean;
  pasteOffset?: boolean;
  cut?: boolean;
  preset?: string;
  connections?: { start?: pptx.ShapeConnection | null; end?: pptx.ShapeConnection | null };
  from?: { x: number; y: number };
  toPoint?: { x: number; y: number };
  text?: string;
  image?: { base64: string; name: string };
  table?: { rows: number; columns: number };
  cell?: {
    row: number;
    column: number;
    edits: Array<{ start: number; end: number; text: string; format?: pptx.TextFormat }>;
  };
  to?: number;
  hidden?: boolean;
  order?: 'front' | 'back' | 'forward' | 'backward';
  stableOrder?: boolean;
  orderPlacement?: { target: number; side: 'before' | 'after' };
  duplicateInParent?: boolean;
}
interface EditRecord {
  command: EditCommand;
  before: string;
  transitionSoundBefore?: string;
}
export interface EditHistory {
  version: 1;
  cursor: number;
  entries: EditRecord[];
}
export const emptyHistory = (): EditHistory => ({ version: 1, cursor: 0, entries: [] });

/** Carry session state across internal save/load, never across a fresh PPTX import. */
export function restoreEditorSession(presentation: pptx.PresentationData, model: EditorModel) {
  for (const slide of pptx.getSlides(presentation)) {
    const prior = model.slides.find((item) => item.key === pptx.getSlidePartName(slide));
    if (!prior) continue;
    const shapes = new Map(
      pptx.getSlideShapes(slide).map((shape) => [pptx.getShapeId(shape), shape]),
    );
    const restored = new Set<string>();
    const visit = (items: EditorShape[]) => {
      for (const item of items) {
        const ids = item.regroupIds;
        if (ids?.length) {
          const key = ids.join(',');
          if (!restored.has(key)) {
            const members = ids.map((id) => shapes.get(id));
            if (members.every((shape) => shape !== undefined)) pptx.rememberRegroupShapes(members);
            restored.add(key);
          }
        }
        if (item.children) visit(item.children);
      }
    };
    visit(prior.shapes);
  }
}
export const editPath = (entry: string) => entry + '.edits.json';
const hash = (value: string) => createHash('sha256').update(value).digest('hex');
// Keep version-1 journal fingerprints stable when UI-only model fields grow.
function historyShape(shape: EditorShape) {
  const { id, name, kind, bounds, rotation, text, textable, format, fill, align } = shape;
  return { id, name, kind, bounds, rotation, text, textable, format, fill, align };
}

function editorFormat(input: pptx.TextFormat): pptx.TextFormat {
  const format: pptx.TextFormat = {};
  for (const key of ['bold', 'italic'] as const)
    if (input[key] !== undefined) {
      if (typeof input[key] !== 'boolean') throw new Error('Invalid font style.');
      format[key] = input[key];
    }
  if (input.underline !== undefined) {
    if (typeof input.underline !== 'boolean' && !['sng', 'none'].includes(input.underline))
      throw new Error('Invalid underline.');
    format.underline = input.underline;
  }
  if (input.size !== undefined) {
    finite(input.size, 1, 400);
    format.size = input.size;
  }
  if (input.font !== undefined) {
    text(input.font);
    format.font = input.font;
  }
  if (input.color !== undefined) format.color = color(input.color ?? '000000');
  if (input.strike !== undefined) {
    if (typeof input.strike !== 'boolean') throw new Error('Invalid strike.');
    format.strike = input.strike;
  }
  if (input.baseline !== undefined) {
    finite(input.baseline, -1, 1);
    format.baseline = input.baseline;
  }
  if (input.kern !== undefined) {
    finite(input.kern, 0, 400000);
    format.kern = input.kern;
  }
  if (input.spc !== undefined) {
    finite(input.spc, -2000, 2000);
    format.spc = input.spc;
  }
  if (input.highlight !== undefined)
    format.highlight = input.highlight === null ? null : color(input.highlight);
  return format;
}

export function editorModel(presentation: pptx.PresentationData): EditorModel {
  const size = pptx.getSlideSize(presentation);
  const notes = pptx.getNotesSize(presentation);
  return {
    firstSlideNumber: pptx.getFirstSlideNumber(presentation),
    notesOrientation: notes && notes.width > notes.height ? 'landscape' : 'portrait',
    ...(size?.type === undefined ? {} : { sizeType: size.type }),
    width: size?.width ?? pptx.inches(13.333333),
    height: size?.height ?? pptx.inches(7.5),
    gridSpacing: pptx.getGridSpacing(presentation),
    snapToGrid: pptx.getSnapToGrid(presentation),
    guidesVisible: pptx.getDrawingGuidesVisible(presentation),
    guides: pptx.getDrawingGuides(presentation)?.map((guide) => ({
      id: guide.id,
      axis: guide.axis,
      color: guide.color,
      offset:
        guide.position -
        (guide.axis === 'x'
          ? (size?.width ?? pptx.inches(13.333333))
          : (size?.height ?? pptx.inches(7.5))) /
          2,
    })) ?? [
      { id: 1, axis: 'x', offset: 0, color: '#888888' },
      { id: 2, axis: 'y', offset: 0, color: '#888888' },
    ],
    showProperties: pptx.getSlideShowProperties(presentation),
    customShows: pptx.getCustomShows(presentation).map((show) => ({
      id: show.id,
      name: show.name,
      slides: show.slides.map(pptx.getSlidePartName),
    })),
    sections: pptx.getSlideSections(presentation).map((section) => ({
      name: section.name,
      slides: section.slides.map(pptx.getSlidePartName),
    })),
    layouts: pptx.getSlideLayouts(presentation).map((layout) => ({
      key: pptx.getSlideLayoutPartName(layout),
      name: pptx.getSlideLayoutName(layout),
      type: pptx.getSlideLayoutType(layout),
      placeholders: pptx
        .getSlideLayoutPlaceholders(layout)
        .map(({ bounds, type }) => ({ bounds, type })),
    })),
    slides: pptx.getSlides(presentation).map((slide) => {
      const descendants = new Set<number>();
      const collect = (shape: pptx.SlideShapeData) => {
        if (pptx.getShapeKind(shape) !== 'group') return;
        for (const child of pptx.getGroupChildren(shape)) {
          descendants.add(pptx.getShapeId(child));
          collect(child);
        }
      };
      const all = pptx.getSlideShapes(slide);
      all.forEach(collect);
      const contexts = slideGeometryContexts(presentation, slide);
      const describeShape = (shape: pptx.SlideShapeData): EditorShape => {
        let paragraphs = 0;
        if (pptx.getShapeKind(shape) === 'shape') {
          try {
            paragraphs = pptx.getShapeParagraphCount(shape);
          } catch (cause) {
            if (!(cause instanceof Error) || !cause.message.includes('has no <p:txBody>'))
              throw cause;
          }
        }
        const textable = paragraphs > 0;
        const body = textable ? pptx.getShapeBodyPrEffective(presentation, shape) : null;
        const runs: EditorShape['runs'] = [];
        const paragraphStyles: EditorShape['paragraphs'] = [];
        let offset = 0;
        for (let p = 0; p < paragraphs; p++) {
          const start = offset;
          let run = 0;
          for (const element of pptx.getShapeParagraphElements(shape, p)) {
            const length = element.kind === 'br' ? 1 : element.text.length;
            const format =
              element.kind === 'r'
                ? pptx.getShapeRunFormatEffective(presentation, shape, p, run++)
                : element.format;
            runs.push({ start: offset, end: offset + length, format: format ?? {} });
            offset += length;
          }
          const properties = pptx.getParagraphPropertiesEffective(presentation, shape, p);
          properties.align ??=
            !pptx.isShapePlaceholder(shape) && !pptx.isShapeTextBox(shape) ? 'center' : 'left';
          paragraphStyles.push({
            start,
            end: offset,
            align: pptx.getParagraphAlignment(shape, p),
            bullets: pptx.getParagraphBullet(shape, p),
            properties,
          });
          offset++;
        }
        return {
          link: objectLink(shape),
          hoverLink: objectLink(shape, true),
          actionSounds: actionSounds(shape),
          textLinks: textable
            ? pptx.getShapeTextRangeClickActions(shape).map(({ start, end, action, tooltip }) => ({
                start,
                end,
                link: {
                  action:
                    action.kind === 'slide'
                      ? { kind: 'slide', slide: pptx.getSlidePartName(action.slide) }
                      : action,
                  tooltip,
                },
              }))
            : [],
          id: pptx.getShapeId(shape),
          name: pptx.getShapeName(shape),
          hidden: pptx.isShapeHidden(shape),
          ...(pptx.getRegroupShapes(shape).length
            ? { regroupIds: pptx.getRegroupShapes(shape).map(pptx.getShapeId) }
            : {}),
          kind: pptx.getShapeKind(shape),
          ...(pptx.getShapeKind(shape) === 'group'
            ? { children: pptx.getGroupChildren(shape).map(describeShape) }
            : {}),
          ...(descendants.has(pptx.getShapeId(shape))
            ? { parentTransform: parentTransform(shape) }
            : {}),
          bounds: pptx.getShapeBoundsResolved(presentation, shape),
          rotation: pptx.getShapeRotation(shape),
          aspectRatioLocked: pptx.getShapeAspectRatioLocked(shape),
          flip: pptx.getShapeFlip(shape),
          preset: pptx.getShapePreset(shape),
          imageCrop: pptx.getShapeImageCrop(shape),
          imageOpacity: pptx.getShapeImageOpacity(shape),
          imageBrightness: pptx.getShapeImageBrightness(shape),
          imageContrast: pptx.getShapeImageContrast(shape),
          connections: {
            start: pptx.getShapeConnection(shape, 'start'),
            end: pptx.getShapeConnection(shape, 'end'),
          },
          connectionSites: connectionSites(presentation, shape),
          text: pptx.getShapeText(shape),
          textable,
          ...(pptx.isTableShape(shape)
            ? {
                table: {
                  style: pptx.getTableStyleFlags(shape),
                  columnWidths: [...pptx.getTableColumnWidths(shape)],
                  rowHeights: [...pptx.getTableRowHeights(shape)],
                  cells: pptx.getTableCells(shape).map((row) =>
                    row.map((cell) => ({
                      text: pptx.getTableCellText(cell),
                      textLinks: textLinks(pptx.getTableCellTextRangeClickActions(cell)),
                      span: pptx.getTableCellSpan(cell),
                      paragraphs: pptx.getTableCellParagraphs(cell),
                      margins: pptx.getTableCellMargins(cell),
                      anchor: pptx.getTableCellAnchor(cell),
                      direction: pptx.getTableCellTextDirection(cell),
                    })),
                  ),
                },
              }
            : {}),
          placeholder: pptx.isShapePlaceholder(shape)
            ? (pptx.getShapePlaceholderType(shape) ?? 'obj')
            : null,
          runs,
          paragraphs: paragraphStyles,
          format:
            textable && pptx.getShapeRunCount(shape, 0)
              ? pptx.getShapeRunFormatEffective(presentation, shape, 0, 0)
              : {},
          fill: pptx.getShapeKind(shape) === 'shape' ? pptx.getShapeFillColor(shape) : null,
          paint: ['shape', 'connector', 'picture'].includes(pptx.getShapeKind(shape))
            ? {
                fill: pptx.getShapeFillEffective(presentation, shape),
                stroke: pptx.getShapeStrokeEffective(presentation, shape),
                dash: pptx.getShapeStrokeDash(shape),
                cap: pptx.getShapeStrokeCap(shape),
                join: pptx.getShapeStrokeJoin(shape),
                compound: pptx.getShapeStrokeCompound(shape),
                headArrow: pptx.getShapeStrokeArrow(shape, 'head'),
                tailArrow: pptx.getShapeStrokeArrow(shape, 'tail'),
                fillOpacity:
                  pptx.getShapeFill(shape).kind === 'solid'
                    ? (pptx.getShapeFillOpacity(shape) ?? 1)
                    : null,
                strokeOpacity:
                  pptx.getShapeStroke(shape).kind === 'solid'
                    ? (pptx.getShapeStrokeOpacity(shape) ?? 1)
                    : null,
              }
            : null,
          align: textable ? pptx.getParagraphAlignment(shape, 0) : null,
          bullets: textable ? pptx.getParagraphBullet(shape, 0) : null,
          anchor: textable ? pptx.getShapeTextAnchor(shape) : null,
          anchorCenter: body?.anchorCenter ?? false,
          textDirection: body?.vert ?? 'horz',
          autoFit: body?.autoFit ?? 'none',
          autoFitParams: body?.autoFitParams ?? null,
          textFrame: body
            ? {
                margins: {
                  left: body.margins.left ?? 91440,
                  top: body.margins.top ?? 45720,
                  right: body.margins.right ?? 91440,
                  bottom: body.margins.bottom ?? 45720,
                },
                anchor:
                  body.anchor ??
                  (!pptx.isShapePlaceholder(shape) && !pptx.isShapeTextBox(shape)
                    ? 'center'
                    : 'top'),
                wrap: body.wrap !== 'none',
                columns: pptx.getShapeTextColumns(shape),
              }
            : null,
        };
      };
      const parentTransform = (
        shape: pptx.SlideShapeData,
      ): NonNullable<EditorShape['parentTransform']> | null => {
        const context = contexts.get(pptx.getShapeId(shape));
        if (!context) return null;
        const origin = context.project({ x: 0, y: 0 });
        const x = context.project({ x: 1000, y: 0 });
        const y = context.project({ x: 0, y: 1000 });
        return [
          (x.x - origin.x) / 1000,
          (x.y - origin.y) / 1000,
          (y.x - origin.x) / 1000,
          (y.y - origin.y) / 1000,
          origin.x,
          origin.y,
        ];
      };
      const shapes = all
        .filter((shape) => !descendants.has(pptx.getShapeId(shape)))
        .map(describeShape);
      const notes = pptx.getSlideNotes(slide) ?? '';
      const hidden = pptx.isSlideHidden(slide);
      return {
        layout: (() => {
          const layout = pptx.getSlideLayout(slide);
          return layout ? pptx.getSlideLayoutPartName(layout) : null;
        })(),
        key: pptx.getSlidePartName(slide),
        shapes,
        notes,
        hidden,
        transition: pptx.getSlideTransition(slide),
        transitionSound: transitionSound(slide),
        animations: pptx.getSlideAnimationSequence(slide),
        token: hash(
          JSON.stringify({
            shapes: shapes.map(historyShape),
            notes,
            hidden,
            xml: pptx.getSlideXmlString(slide),
          }),
        ),
      };
    }),
  };
}
export async function readHistory(entry: string): Promise<EditHistory> {
  let raw: string;
  try {
    raw = await readFile(editPath(entry), 'utf8');
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') return emptyHistory();
    throw error;
  }
  const value = JSON.parse(raw) as EditHistory;
  if (
    value.version !== 1 ||
    !Array.isArray(value.entries) ||
    !Number.isInteger(value.cursor) ||
    value.cursor < 0 ||
    value.cursor > value.entries.length ||
    value.entries.length > 10000
  )
    throw new Error('Invalid editor history: ' + editPath(entry));
  return value;
}
export async function writeHistory(entry: string, history: EditHistory) {
  const destination = editPath(entry),
    temporary = destination + '.' + randomUUID() + '.tmp';
  try {
    await writeFile(temporary, JSON.stringify(history, null, 2) + '\n');
    await rename(temporary, destination);
  } finally {
    await rm(temporary, { force: true });
  }
}
function beforeToken(presentation: pptx.PresentationData, command: EditCommand) {
  const model = editorModel(presentation);
  const historySlide = (slide: EditorSlide | undefined) =>
    slide
      ? {
          key: slide.key,
          shapes: slide.shapes.map(historyShape),
          notes: slide.notes,
          hidden: slide.hidden,
          token: slide.token,
        }
      : null;
  // Slide insertion/reordering depends on the deck order as well as the target slide.
  return hash(
    JSON.stringify({
      slide: historySlide(model.slides[command.slide]),
      ...(command.type === 'ink'
        ? {
            inkTargets: command.ink?.map((annotation) =>
              historySlide(model.slides.find((slide) => slide.key === annotation.slide)),
            ),
          }
        : {}),
      ...(command.type === 'object-actions' && command.actionSounds !== undefined
        ? {
            sounds: model.slides[command.slide]?.shapes
              .filter((shape) => command.ids?.includes(shape.id))
              .map((shape) => [shape.id, shape.actionSounds]),
          }
        : {}),
      ...(command.type === 'object-actions'
        ? {
            actions: model.slides[command.slide]?.shapes
              .filter((shape) => command.ids?.includes(shape.id))
              .map((shape) => [shape.id, shape.link, shape.hoverLink]),
            linkDestinations: model.slides.map((slide) => slide.key),
          }
        : {}),
      ...(command.type === 'object-link'
        ? {
            links: model.slides[command.slide]?.shapes
              .filter((shape) => command.ids?.includes(shape.id))
              .map((shape) => [shape.id, shape.link]),
            linkDestinations: model.slides.map((slide) => slide.key),
          }
        : {}),
      ...(command.type === 'text-link'
        ? {
            textLinks: model.slides[command.slide]?.shapes
              .filter((shape) => command.ids?.includes(shape.id))
              .map((shape) => [
                shape.id,
                command.cell
                  ? shape.table?.cells.map((row) => row.map((cell) => cell.textLinks))
                  : shape.textLinks,
              ]),
            linkDestinations: model.slides.map((slide) => slide.key),
          }
        : {}),
      ...(command.type === 'picture-replace'
        ? {
            pictureSources: pptx
              .getSlideShapes(pptx.getSlides(presentation)[command.slide]!)
              .filter((shape) => command.ids?.includes(pptx.getShapeId(shape)))
              .map((shape) => [
                pptx.getShapeId(shape),
                createHash('sha256')
                  .update(pptx.getShapeImageBytes(shape) ?? new Uint8Array())
                  .digest('hex'),
              ]),
          }
        : {}),
      ...(command.type === 'text-replace-all'
        ? { searchedSlides: model.slides.map(historySlide) }
        : {}),
      ...(command.type === 'slide-size' ? { size: pptx.getSlideSize(presentation) } : {}),
      ...(command.type === 'slide-size' && command.size?.firstSlideNumber !== undefined
        ? { firstSlideNumber: pptx.getFirstSlideNumber(presentation) }
        : {}),
      ...(command.type === 'slide-size' && command.size?.notesOrientation !== undefined
        ? { notesSize: pptx.getNotesSize(presentation) }
        : {}),
      ...(command.type === 'slide-size' && command.scaleContent
        ? {
            resizedSlides: model.slides.map(historySlide),
            resizedSources: pptx
              .listPackageParts(presentation)
              .filter((part) =>
                /presentationml\.(presentation\.main|slideLayout|slideMaster)\+xml$/.test(
                  part.contentType,
                ),
              )
              .map((part) => [
                part.name,
                createHash('sha256')
                  .update(pptx.readPackagePart(presentation, part.name)!)
                  .digest('hex'),
              ]),
          }
        : {}),
      ...(command.type === 'guides-visible' || command.guidesVisible !== undefined
        ? { guidesVisible: model.guidesVisible }
        : {}),
      ...(command.type === 'grid-spacing' ? { gridSpacing: model.gridSpacing } : {}),
      ...(command.type === 'grid-snap' || command.snapToGrid !== undefined
        ? { snapToGrid: model.snapToGrid }
        : {}),
      ...(command.type === 'guides' ? { guides: model.guides } : {}),
      ...(command.type === 'transition-apply-all'
        ? { slides: model.slides.map(historySlide) }
        : {}),
      ...(command.type === 'show-properties'
        ? {
            showProperties: model.showProperties,
            customShows: model.customShows,
            slideKeys: model.slides.map((slide) => slide.key),
          }
        : {}),
      ...(command.type === 'custom-shows'
        ? {
            customShows: model.customShows,
            slideKeys: model.slides.map((slide) => slide.key),
          }
        : {}),
      ...(command.changes?.autoFit !== undefined
        ? {
            autoFit: model.slides[command.slide]?.shapes
              .filter((shape) => command.ids?.includes(shape.id))
              .map((shape) => [shape.id, shape.autoFit]),
          }
        : {}),
      ...(command.textFits !== undefined
        ? {
            textFits: model.slides[command.slide]?.shapes
              .filter((shape) => command.ids?.includes(shape.id))
              .map((shape) => [
                shape.id,
                shape.autoFit,
                shape.autoFitParams,
                shape.textFrame,
                shape.textDirection,
              ]),
          }
        : {}),
      ...(command.changes?.anchorCenter !== undefined
        ? {
            anchorCenters: model.slides[command.slide]?.shapes
              .filter((shape) => command.ids?.includes(shape.id))
              .map((shape) => [shape.id, shape.anchorCenter]),
          }
        : {}),
      ...(command.changes?.autoFitParams !== undefined
        ? {
            autoFitParams: model.slides[command.slide]?.shapes
              .filter((shape) => command.ids?.includes(shape.id))
              .map((shape) => [shape.id, shape.autoFit, shape.autoFitParams]),
          }
        : {}),
      ...(command.changes?.imageBrightness !== undefined ||
      command.changes?.imageContrast !== undefined
        ? {
            imageCorrections: model.slides[command.slide]?.shapes
              .filter((shape) => command.ids?.includes(shape.id))
              .map((shape) => [shape.id, shape.imageBrightness, shape.imageContrast]),
          }
        : {}),
      ...(command.changes?.imageOpacity !== undefined
        ? {
            imageOpacities: model.slides[command.slide]?.shapes
              .filter((shape) => command.ids?.includes(shape.id))
              .map((shape) => [shape.id, shape.imageOpacity]),
          }
        : {}),
      ...(command.changes?.aspectRatioLocked !== undefined
        ? {
            aspectLocks: model.slides[command.slide]?.shapes
              .filter((shape) => command.ids?.includes(shape.id))
              .map((shape) => [shape.id, shape.aspectRatioLocked]),
          }
        : {}),
      ...(command.changes?.margins !== undefined ||
      command.changes?.wrap !== undefined ||
      command.changes?.textColumns !== undefined
        ? {
            textFrames: model.slides[command.slide]?.shapes
              .filter((shape) => command.ids?.includes(shape.id))
              .map((shape) => [shape.id, shape.textFrame]),
          }
        : {}),
      ...(command.type.startsWith('section-')
        ? { sections: model.sections, order: model.slides.map((s) => s.key) }
        : {}),
      ...(command.type === 'section-delete'
        ? {
            deletedSlides: model.slides
              .filter((slide) => model.sections[command.section ?? -1]?.slides.includes(slide.key))
              .map(historySlide),
          }
        : {}),
      ...(command.type === 'paste' && command.snapshot === undefined
        ? { source: historySlide(model.slides[command.source ?? -1]) }
        : {}),
      ...(command.type.startsWith('slide-') ? { order: model.slides.map((s) => s.key) } : {}),
    }),
  );
}
function transitionSoundToken(presentation: pptx.PresentationData, command: EditCommand) {
  if (!['transition-sound', 'transition-apply-all'].includes(command.type)) return undefined;
  const slides = pptx.getSlides(presentation);
  const targets = command.type === 'transition-apply-all' ? slides : [slides[command.slide]!];
  return hash(JSON.stringify(targets.map(transitionSound)));
}
export function recordEdit(presentation: pptx.PresentationData, command: EditCommand): EditRecord {
  const before = beforeToken(presentation, command);
  const transitionSoundBefore = transitionSoundToken(presentation, command);
  applyEdit(presentation, command);
  return {
    command,
    before,
    ...(transitionSoundBefore !== undefined ? { transitionSoundBefore } : {}),
  };
}
export function replayEdits(presentation: pptx.PresentationData, history: EditHistory) {
  for (const [index, record] of history.entries.slice(0, history.cursor).entries()) {
    if (
      record.before !== beforeToken(presentation, record.command) ||
      (record.transitionSoundBefore !== undefined &&
        record.transitionSoundBefore !== transitionSoundToken(presentation, record.command))
    )
      throw new Error(
        `Visual edit ${index + 1} conflicts with the TSX source. Undo the conflicting edits or reconcile ${'the .edits.json file'} before continuing. Your source and edits are both preserved.`,
      );
    applyEdit(presentation, record.command);
  }
}
function finite(value: unknown, min: number, max: number): asserts value is number {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < min || value > max)
    throw new Error('Invalid numeric edit value.');
}
function bounds(value: NonNullable<EditorShape['bounds']>, allowZero = false): pptx.ShapeBounds {
  finite(value.x, -1e9, 1e9);
  finite(value.y, -1e9, 1e9);
  finite(value.w, allowZero ? 0 : 1, 1e9);
  finite(value.h, allowZero ? 0 : 1, 1e9);
  return { x: pptx.emu(value.x), y: pptx.emu(value.y), w: pptx.emu(value.w), h: pptx.emu(value.h) };
}
function text(value: unknown): asserts value is string {
  if (typeof value !== 'string' || value.length > 100000) throw new Error('Invalid text edit.');
}
function validateTextRange(value: string, start: number, end: number) {
  if (
    !Number.isInteger(start) ||
    !Number.isInteger(end) ||
    start < 0 ||
    end < start ||
    end > value.length
  )
    throw new Error('Invalid text range.');
  for (const offset of [start, end])
    if (
      offset > 0 &&
      offset < value.length &&
      /[\uD800-\uDBFF]/.test(value[offset - 1]!) &&
      /[\uDC00-\uDFFF]/.test(value[offset]!)
    )
      throw new Error('Text range splits a surrogate pair.');
}
function color(value: string) {
  if (!/^#?[0-9a-f]{6}$/i.test(value)) throw new Error('Invalid color.');
  return value.replace(/^#/, '');
}
function shapeFrame(presentation: pptx.PresentationData, shape: pptx.SlideShapeData) {
  return {
    bounds: pptx.getShapeBoundsResolved(presentation, shape),
    rotation: pptx.getShapeRotation(shape),
    flip: pptx.getShapeFlip(shape),
  };
}
function connectionSites(presentation: pptx.PresentationData, shape: pptx.SlideShapeData) {
  const frame = shapeFrame(presentation, shape);
  return frame.bounds
    ? pptx.getShapeConnectionSites(shape, frame.bounds).map((site) => shapePoint(frame, site))
    : [];
}
function setConnections(
  presentation: pptx.PresentationData,
  slide: pptx.SlideData,
  shape: pptx.SlideShapeData,
  connections: NonNullable<EditCommand['connections']>,
) {
  for (const end of ['start', 'end'] as const) {
    const connection = connections[end];
    if (connection === undefined) continue;
    if (connection !== null) {
      const target = pptx
        .getSlideShapes(slide)
        .find((s) => pptx.getShapeId(s) === connection.shapeId);
      if (
        !target ||
        !Number.isInteger(connection.siteIndex) ||
        !connectionSites(presentation, target)[connection.siteIndex]
      )
        throw new Error('Invalid connection site.');
    }
    pptx.setShapeConnection(shape, end, connection);
  }
}
function registerCopy(
  original: pptx.SlideShapeData,
  copied: pptx.SlideShapeData,
  copies: Map<number, pptx.SlideShapeData>,
  originals: Map<number, EditorShape['connections']>,
) {
  const id = pptx.getShapeId(original);
  copies.set(id, copied);
  originals.set(id, {
    start: pptx.getShapeConnection(original, 'start'),
    end: pptx.getShapeConnection(original, 'end'),
  });
  const copiedChildren = pptx.getGroupChildren(copied);
  for (const [index, child] of pptx.getGroupChildren(original).entries()) {
    const copiedChild = copiedChildren[index];
    if (copiedChild) registerCopy(child, copiedChild, copies, originals);
  }
}
function reconnectCopies(
  copies: Map<number, pptx.SlideShapeData>,
  originals: Map<number, EditorShape['connections']>,
) {
  for (const [id, copy] of copies) {
    if (pptx.getShapeKind(copy) !== 'connector') continue;
    for (const end of ['start', 'end'] as const) {
      const attachment = originals.get(id)?.[end];
      const target = attachment && copies.get(attachment.shapeId);
      pptx.setShapeConnection(
        copy,
        end,
        target ? { shapeId: pptx.getShapeId(target), siteIndex: attachment!.siteIndex } : null,
      );
    }
  }
}
type Point = { x: number; y: number };
function slideGeometryContexts(presentation: pptx.PresentationData, slide: pptx.SlideData) {
  const all = pptx.getSlideShapes(slide);
  const children = new Set(
    all
      .filter((s) => pptx.getShapeKind(s) === 'group')
      .flatMap((s) => pptx.getGroupChildren(s).map(pptx.getShapeId)),
  );
  const contexts = new Map<
    number,
    { project: (p: Point) => Point; unproject: (p: Point) => Point }
  >();
  const visit = (shape: pptx.SlideShapeData, project: (p: Point) => Point) => {
    const origin = project({ x: 0, y: 0 }),
      px = project({ x: 1000, y: 0 }),
      py = project({ x: 0, y: 1000 });
    const a = (px.x - origin.x) / 1000,
      b = (px.y - origin.y) / 1000,
      c = (py.x - origin.x) / 1000,
      d = (py.y - origin.y) / 1000;
    const det = a * d - b * c;
    if (Math.abs(det) < 1e-12) return;
    contexts.set(pptx.getShapeId(shape), {
      project,
      unproject: (p) => ({
        x: (d * (p.x - origin.x) - c * (p.y - origin.y)) / det,
        y: (-b * (p.x - origin.x) + a * (p.y - origin.y)) / det,
      }),
    });
    const group = pptx.getGroupTransform(shape);
    if (!group || !group.inner.w || !group.inner.h) return;
    const frame = shapeFrame(presentation, shape);
    for (const child of pptx.getGroupChildren(shape))
      visit(child, (p) =>
        project(
          shapePoint(frame, {
            x: ((p.x - group.inner.x) * group.outer.w) / group.inner.w,
            y: ((p.y - group.inner.y) * group.outer.h) / group.inner.h,
          }),
        ),
      );
  };
  for (const shape of all) if (!children.has(pptx.getShapeId(shape))) visit(shape, (p) => p);
  return contexts;
}
function selectedTableCells(table: pptx.SlideShapeData, range: EditCommand['tableRange']) {
  const dimensions = pptx.getTableDimensions(table);
  if (
    !range ||
    ![range.row, range.column, range.rows, range.columns].every(Number.isInteger) ||
    range.row < 0 ||
    range.column < 0 ||
    range.rows < 1 ||
    range.columns < 1 ||
    range.row + range.rows > dimensions.rows ||
    range.column + range.columns > dimensions.cols
  )
    throw new Error('Invalid table cell range.');
  const targets: Array<{ row: number; column: number }> = [];
  for (let row = 0; row < dimensions.rows; row++) {
    for (let column = 0; column < dimensions.cols; column++) {
      const span = pptx.getTableCellSpan(pptx.getTableCell(table, row, column));
      if (span.hMerge || span.vMerge) continue;
      const bottom = row + span.rowSpan,
        right = column + span.gridSpan;
      if (
        row >= range.row + range.rows ||
        bottom <= range.row ||
        column >= range.column + range.columns ||
        right <= range.column
      )
        continue;
      if (
        row < range.row ||
        column < range.column ||
        bottom > range.row + range.rows ||
        right > range.column + range.columns
      )
        throw new Error('Select the entire merged cell.');
      targets.push({ row, column });
    }
  }
  return targets;
}
export function applyEdit(presentation: pptx.PresentationData, command: EditCommand) {
  applyEditMutation(presentation, command);
  if (
    ![
      'table-split',
      'table-merge',
      'table-cell-text',
      'table-cell-format',
      'table-cell-case',
      'table-cell-indent',
      'table-cell-paragraph',
      'table-cell-line-spacing',
      'table-cell-bullets',
      'table-cell-align',
      'table-cell-layout',
      'table-range-format',
      'table-margins',
      'table-cell-size',
      'table-insert',
      'table-delete',
      'table-row-append',
    ].includes(command.type)
  )
    return;
  const slide = pptx.getSlides(presentation)[command.slide];
  if (!slide) return;
  for (const shape of pptx.getSlideShapes(slide))
    if (command.ids?.includes(pptx.getShapeId(shape)) && pptx.isTableShape(shape))
      fitTableText(presentation, shape);
}

function applyEditMutation(presentation: pptx.PresentationData, command: EditCommand) {
  const priorSlides = pptx.getSlides(presentation);
  const sections = pptx.getSlideSections(presentation).map((section) => ({
    name: section.name,
    slides: [...section.slides],
  }));
  applyEditInternal(presentation, command);
  if (
    sections.length &&
    ['slide-add', 'slide-duplicate', 'slide-delete', 'slide-move'].includes(command.type)
  ) {
    const ordered = pptx.getSlides(presentation);
    const key = pptx.getSlidePartName;
    const owner = new Map(
      sections.flatMap((section, i) => section.slides.map((slide) => [key(slide), i] as const)),
    );
    const source = priorSlides[command.slide];
    if (source && command.type === 'slide-move') {
      const destination = priorSlides[command.to!];
      if (destination) owner.set(key(source), owner.get(key(destination)) ?? 0);
    }
    const sourceOwner = source ? (owner.get(key(source)) ?? 0) : 0;
    for (const slide of ordered) if (!owner.has(key(slide))) owner.set(key(slide), sourceOwner);
    for (const [i, section] of sections.entries())
      section.slides = ordered.filter((slide) => owner.get(key(slide)) === i);
    pptx.setSlideSections(presentation, sections);
  }
  if (
    ![
      'insert',
      'update',
      'text-replace-all',
      'delete',
      'paste',
      'duplicate',
      'group',
      'ungroup',
      'slide-layout',
      'slide-reset',
    ].includes(command.type)
  )
    return;
  // Topology is stored in native OOXML; geometry is refreshed after the entire selection moves.
  for (const slide of pptx.getSlides(presentation)) {
    const shapes = pptx.getSlideShapes(slide);
    const contexts = slideGeometryContexts(presentation, slide);
    for (const shape of shapes) {
      if (pptx.getShapeKind(shape) !== 'connector') continue;
      const context = contexts.get(pptx.getShapeId(shape));
      if (!context) continue;
      const frame = shapeFrame(presentation, shape);
      if (!frame.bounds) continue;
      const points = connectorEndpoints(frame);
      let attached = false;
      for (const end of ['start', 'end'] as const) {
        const connection = pptx.getShapeConnection(shape, end);
        if (!connection) continue;
        const target = shapes.find((s) => pptx.getShapeId(s) === connection.shapeId);
        const site = target && connectionSites(presentation, target)[connection.siteIndex];
        if (!target) {
          pptx.setShapeConnection(shape, end, null);
          continue;
        }
        const targetContext = contexts.get(connection.shapeId);
        if (!site || !targetContext) continue;
        points[end] = context.unproject(targetContext.project(site));
        attached = true;
      }
      if (attached) {
        const updated = connectorFrame(points.start, points.end, frame.rotation);
        pptx.setShapeBounds(shape, bounds(updated.bounds, true));
        pptx.setShapeFlip(shape, updated.flip);
      }
    }
  }
}
function applyEditInternal(presentation: pptx.PresentationData, command: EditCommand) {
  if (!command || typeof command !== 'object') throw new Error('Invalid edit.');
  const slides = pptx.getSlides(presentation);
  if (
    !Number.isInteger(command.slide) ||
    command.slide < 0 ||
    command.slide >= Math.max(1, slides.length)
  )
    throw new Error('Slide no longer exists.');
  const slide = slides[command.slide];
  if (command.type === 'ink') {
    keepShowInk(presentation, command.ink);
    return;
  }
  if (command.type === 'text-replace-all') {
    const replacement = command.replacement;
    if (!replacement) throw new Error('Invalid replacement.');
    text(replacement.query);
    text(replacement.text);
    if (!replacement.query) throw new Error('Find text must not be empty.');
    for (const option of ['matchCase', 'wholeWords'] as const)
      if (
        replacement.options?.[option] !== undefined &&
        typeof replacement.options[option] !== 'boolean'
      )
        throw new Error('Invalid search option.');
    const matches = findText(
      editorModel(presentation).slides,
      replacement.query,
      replacement.options,
    );
    const targets = matches.map((match) => {
      const shape = pptx
        .getSlideShapes(slides[match.slide]!)
        .find((shape) => pptx.getShapeId(shape) === match.shape);
      if (!shape) throw new Error('Shape no longer exists.');
      return { ...match, target: shape };
    });
    const fits = command.textFits ?? [];
    if (!Array.isArray(fits)) throw new Error('Invalid text fit measurements.');
    const seen = new Set<string>();
    for (const fit of fits) {
      const key = fit.slide + ':' + fit.id;
      const match = targets.find((match) => match.slide === fit.slide && match.shape === fit.id);
      if (!match || seen.has(key)) throw new Error('Invalid text fit target.');
      seen.add(key);
      const mode = pptx.getShapeBodyPrEffective(presentation, match.target).autoFit;
      if (fit.bounds) {
        if (mode !== 'shape') throw new Error('Size measurement requires shape autofit.');
        bounds(fit.bounds);
      }
      if (fit.autoFitParams) {
        if (mode !== 'normal') throw new Error('Font measurement requires normal autofit.');
        finite(fit.autoFitParams.fontScale, Number.MIN_VALUE, 1);
        finite(fit.autoFitParams.lnSpcReduction, 0, 1);
      }
    }
    // Work backwards so every range still refers to the original text.
    for (const match of targets.reverse())
      pptx.replaceShapeTextRange(match.target, match.start, match.end, replacement.text);
    for (const fit of fits) {
      const target = targets.find(
        (match) => match.slide === fit.slide && match.shape === fit.id,
      )!.target;
      if (fit.bounds) pptx.setShapeBounds(target, bounds(fit.bounds));
      if (fit.autoFitParams) pptx.setShapeTextAutoFit(target, 'normal', fit.autoFitParams);
    }
    return;
  }
  if (command.type === 'slide-size') {
    if (
      !command.size ||
      typeof command.size !== 'object' ||
      typeof command.size.width !== 'number' ||
      typeof command.size.height !== 'number'
    )
      throw new Error('Invalid slide size.');
    pptx.setSlideSize(
      presentation,
      {
        width: pptx.emu(command.size.width),
        height: pptx.emu(command.size.height),
        ...(command.size.type !== undefined ? { type: command.size.type } : {}),
      },
      {
        scaleContent: command.scaleContent ?? false,
        ...(command.size.notesOrientation === undefined
          ? {}
          : { notesOrientation: command.size.notesOrientation }),
        ...(command.size.firstSlideNumber === undefined
          ? {}
          : { firstSlideNumber: command.size.firstSlideNumber }),
      },
    );
    return;
  }
  if (command.type === 'show-properties') {
    if (!command.showProperties) throw new Error('Missing slideshow settings.');
    pptx.setSlideShowProperties(presentation, command.showProperties);
    return;
  }
  if (command.type === 'custom-shows') {
    if (!Array.isArray(command.customShows)) throw new Error('Invalid custom shows.');
    const slides = new Map(
      pptx.getSlides(presentation).map((slide) => [pptx.getSlidePartName(slide) as string, slide]),
    );
    const shows = command.customShows.map((show) => {
      if (!show || !Array.isArray(show.slides) || typeof show.name !== 'string')
        throw new Error('Invalid custom show.');
      return {
        id: show.id,
        name: show.name,
        slides: show.slides.map((key) => {
          const slide = slides.get(key);
          if (!slide) throw new Error('A custom show slide no longer exists.');
          return slide;
        }),
      };
    });
    pptx.setCustomShows(presentation, shows);
    return;
  }
  if (command.type === 'grid-snap') {
    if (typeof command.snapToGrid !== 'boolean') throw new Error('Invalid grid snapping.');
    pptx.setSnapToGrid(presentation, command.snapToGrid);
    return;
  }
  if (command.type === 'grid-spacing') {
    if (command.snapToGrid !== undefined && typeof command.snapToGrid !== 'boolean')
      throw new Error('Invalid grid snapping.');
    if (!command.gridSpacing) throw new Error('Invalid grid spacing.');
    if (command.guidesVisible !== undefined && typeof command.guidesVisible !== 'boolean')
      throw new Error('Invalid guide visibility.');
    pptx.setGridSpacing(presentation, command.gridSpacing);
    if (command.guidesVisible !== undefined)
      pptx.setDrawingGuidesVisible(presentation, command.guidesVisible);
    if (command.snapToGrid !== undefined) pptx.setSnapToGrid(presentation, command.snapToGrid);
    return;
  }
  if (command.type === 'guides-visible') {
    if (command.snapToGrid !== undefined && typeof command.snapToGrid !== 'boolean')
      throw new Error('Invalid grid snapping.');
    if (typeof command.guidesVisible !== 'boolean') throw new Error('Invalid guide visibility.');
    pptx.setDrawingGuidesVisible(presentation, command.guidesVisible);
    if (command.snapToGrid !== undefined) pptx.setSnapToGrid(presentation, command.snapToGrid);
    return;
  }
  if (command.type === 'guides') {
    if (command.snapToGrid !== undefined && typeof command.snapToGrid !== 'boolean')
      throw new Error('Invalid grid snapping.');
    if (command.guidesVisible !== undefined && typeof command.guidesVisible !== 'boolean')
      throw new Error('Invalid guide visibility.');
    if (!Array.isArray(command.guides)) throw new Error('Invalid drawing guides.');
    const model = editorModel(presentation);
    pptx.setDrawingGuides(
      presentation,
      command.guides.map((guide) => ({
        id: guide?.id,
        axis: guide?.axis,
        color: guide?.color,
        position: guide?.offset + (guide?.axis === 'x' ? model.width : model.height) / 2,
      })),
    );
    if (command.guidesVisible !== undefined)
      pptx.setDrawingGuidesVisible(presentation, command.guidesVisible);
    if (command.snapToGrid !== undefined) pptx.setSnapToGrid(presentation, command.snapToGrid);
    return;
  }
  if (command.type.startsWith('section-')) {
    const sections = pptx
      .getSlideSections(presentation)
      .map((section) => ({ name: section.name, slides: [...section.slides] }));
    const key = pptx.getSlidePartName;
    if (command.type === 'section-remove-all') {
      pptx.setSlideSections(presentation, []);
      return;
    }
    if (command.type === 'section-add') {
      text(command.text ?? 'Untitled Section');
      if (!slide) throw new Error('Slide no longer exists.');
      const firstSection = !sections.length;
      if (!sections.length) sections.push({ name: 'Default Section', slides: [...slides] });
      const owner = sections.findIndex((section) =>
        section.slides.some((s) => key(s) === key(slide)),
      );
      if (owner < 0) throw new Error('Slide is outside the section list.');
      const current = sections[owner]!;
      const position = current.slides.findIndex((s) => key(s) === key(slide));
      if (position === 0 && firstSection) current.name = command.text ?? 'Untitled Section';
      else
        sections.splice(owner + 1, 0, {
          name: command.text ?? 'Untitled Section',
          slides: current.slides.splice(position),
        });
    } else {
      if (!Number.isInteger(command.section) || !sections[command.section!])
        throw new Error('Section no longer exists.');
      const index = command.section!;
      if (command.type === 'section-rename') {
        text(command.text);
        sections[index]!.name = command.text!;
      } else if (command.type === 'section-move') {
        if (!Number.isInteger(command.to) || command.to! < 0 || command.to! >= sections.length)
          throw new Error('Invalid section position.');
        const moved = sections.splice(index, 1)[0]!;
        sections.splice(command.to!, 0, moved);
        const members = new Set(sections.flatMap((section) => section.slides.map(key)));
        const ordered = [
          ...slides.filter((slide) => !members.has(key(slide))),
          ...sections.flatMap((section) => section.slides),
        ];
        ordered.forEach((slide, i) => pptx.moveSlide(presentation, slide, i));
      } else if (command.type === 'section-delete') {
        const removed = sections.splice(index, 1)[0]!;
        for (const slide of removed.slides) pptx.removeSlide(presentation, slide);
      } else if (command.type === 'section-remove') {
        const removed = sections.splice(index, 1)[0]!;
        if (sections.length) {
          const recipient = sections[Math.max(0, index - 1)]!;
          recipient.slides = slides.filter((s) =>
            [...recipient.slides, ...removed.slides].some((member) => key(member) === key(s)),
          );
        }
      } else throw new Error('Unknown section operation.');
    }
    pptx.setSlideSections(presentation, sections);
    return;
  }
  if (command.type === 'slide-add') {
    const layout =
      command.layout === undefined
        ? null
        : pptx.findSlideLayoutByPartName(presentation, command.layout);
    if (command.layout !== undefined && !layout) throw new Error('Slide layout no longer exists.');
    const added = layout
      ? pptx.addSlide(presentation, { layout })
      : pptx.addBlankSlide(presentation);
    pptx.moveSlide(presentation, added, slide ? command.slide + 1 : 0);
    return;
  }
  if (!slide) throw new Error('Slide no longer exists.');
  const target = (id: number) => {
    const shape = pptx.getSlideShapes(slide).find((shape) => pptx.getShapeId(shape) === id);
    if (!shape) throw new Error('Shape no longer exists.');
    return shape;
  };
  const ids = [...new Set(command.ids ?? [])];
  if (ids.length > 1000 || ids.some((id) => !Number.isInteger(id)))
    throw new Error('Invalid selection.');
  const shapes = command.type === 'paste' ? [] : ids.map(target);
  switch (command.type) {
    case 'animation-settings':
      if (
        !Array.isArray(command.animationIds) ||
        !command.animationIds.length ||
        command.animationIds.some((id) => typeof id !== 'string') ||
        !command.animationSettings
      )
        throw new Error('Select animation effects and settings.');
      pptx.setSlideAnimationsSettings(slide, command.animationIds, command.animationSettings);
      break;
    case 'animation-reorder':
      if (
        !Array.isArray(command.animationIds) ||
        !command.animationIds.length ||
        command.animationIds.some((id) => typeof id !== 'string') ||
        (command.animationBeforeId !== null && typeof command.animationBeforeId !== 'string')
      )
        throw new Error('Select animation effects and a destination.');
      pptx.reorderSlideAnimations(slide, command.animationIds, command.animationBeforeId);
      break;
    case 'animation-move':
      if (!command.animationDirection) throw new Error('Select an animation move direction.');
      if (command.animationIds) {
        if (
          !Array.isArray(command.animationIds) ||
          !command.animationIds.length ||
          command.animationIds.some((id) => typeof id !== 'string')
        )
          throw new Error('Select animation effects.');
        pptx.moveSlideAnimations(slide, command.animationIds, command.animationDirection);
      } else {
        if (typeof command.animationId !== 'string') throw new Error('Select an animation effect.');
        pptx.moveSlideAnimation(slide, command.animationId, command.animationDirection);
      }
      break;
    case 'animation-remove':
      if (command.animationIds) {
        if (
          !Array.isArray(command.animationIds) ||
          !command.animationIds.length ||
          command.animationIds.some((id) => typeof id !== 'string')
        )
          throw new Error('Select animation effects.');
        pptx.removeSlideAnimations(slide, command.animationIds);
      } else {
        if (typeof command.animationId !== 'string') throw new Error('Select an animation effect.');
        pptx.removeSlideAnimation(slide, command.animationId);
      }
      break;
    case 'animation-start':
      if (!command.animationStart || typeof command.animationStart.id !== 'string')
        throw new Error('Missing animation start setting.');
      pptx.setSlideAnimationStart(slide, command.animationStart.id, command.animationStart.start);
      break;
    case 'animation-delay':
      if (!command.animationDelay || typeof command.animationDelay.id !== 'string')
        throw new Error('Missing animation timing.');
      pptx.setSlideAnimationDelay(slide, command.animationDelay.id, command.animationDelay.delayMs);
      break;
    case 'animation-duration':
      if (!command.animationTiming || typeof command.animationTiming.id !== 'string')
        throw new Error('Missing animation timing.');
      pptx.setSlideAnimationDuration(
        slide,
        command.animationTiming.id,
        command.animationTiming.durationMs,
      );
      break;
    case 'animation-effect':
      if (typeof command.animationId !== 'string' || !command.animation)
        throw new Error('Select an animation effect.');
      pptx.setSlideAnimationEffect(slide, command.animationId, command.animation.effect);
      break;
    case 'animation-add':
      if (!shapes.length || !command.animation)
        throw new Error('Select an object and animation effect.');
      for (const shape of shapes) pptx.setShapeAnimation(shape, command.animation);
      break;
    case 'table-borders': {
      const data = command.tableBorders;
      if (
        shapes.length !== 1 ||
        !pptx.isTableShape(shapes[0]!) ||
        !data ||
        ![
          'all',
          'none',
          'outside',
          'inside',
          'horizontal',
          'vertical',
          'left',
          'right',
          'top',
          'bottom',
          'tlToBr',
          'blToTr',
        ].includes(data.mode) ||
        !['solid', 'dash', 'dot', 'dashDot'].includes(data.dash)
      )
        throw new Error('Invalid table borders.');
      finite(data.weight, 0.25, 6);
      const pen = {
        color: color(data.color),
        widthEmu: Math.round(data.weight * 12700),
        dash: data.dash,
      };
      if (
        command.cell &&
        (!Number.isInteger(command.cell.row) ||
          !Number.isInteger(command.cell.column) ||
          command.cell.row < 0 ||
          command.cell.column < 0)
      )
        throw new Error('Invalid table cell.');
      applyTableBorders(
        presentation,
        shapes[0]!,
        data.mode,
        pen,
        command.tableRange ?? command.cell,
      );
      return;
    }
    case 'table-margins': {
      if (shapes.length !== 1 || !pptx.isTableShape(shapes[0]!))
        throw new Error('Select one table to change its cell margins.');
      const margins = command.tableMargins;
      if (
        !margins ||
        typeof margins !== 'object' ||
        Array.isArray(margins) ||
        !Object.keys(margins).length ||
        Object.keys(margins).some((key) => !['left', 'right', 'top', 'bottom'].includes(key))
      )
        throw new Error('Invalid table margins.');
      for (const value of Object.values(margins)) finite(value, 0, 51206400);
      const data = command.cell;
      if (
        data &&
        (!Number.isInteger(data.row) ||
          !Number.isInteger(data.column) ||
          data.row < 0 ||
          data.column < 0)
      )
        throw new Error('Invalid table cell.');
      const cells = command.tableRange
        ? selectedTableCells(shapes[0]!, command.tableRange).map(({ row, column }) =>
            pptx.getTableCell(shapes[0]!, row, column),
          )
        : data
          ? [pptx.getTableCell(shapes[0]!, data.row, data.column)]
          : pptx.getTableCells(shapes[0]!).flat();
      if (data) {
        const span = pptx.getTableCellSpan(cells[0]!);
        if (span.hMerge || span.vMerge) throw new Error('Edit the merged cell anchor.');
      }
      for (const cell of cells)
        pptx.setTableCellMargins(cell, { ...pptx.getTableCellMargins(cell), ...margins });
      return;
    }
    case 'table-fill': {
      if (shapes.length !== 1 || !pptx.isTableShape(shapes[0]!))
        throw new Error('Select one table to change its shading.');
      if (command.tableFill !== null && typeof command.tableFill !== 'string')
        throw new Error('Invalid table shading.');
      const fill = command.tableFill === null ? null : color(command.tableFill);
      const data = command.cell;
      if (
        data &&
        (!Number.isInteger(data.row) ||
          !Number.isInteger(data.column) ||
          data.row < 0 ||
          data.column < 0)
      )
        throw new Error('Invalid table cell.');
      const cells = command.tableRange
        ? selectedTableCells(shapes[0]!, command.tableRange).map(({ row, column }) =>
            pptx.getTableCell(shapes[0]!, row, column),
          )
        : data
          ? [pptx.getTableCell(shapes[0]!, data.row, data.column)]
          : pptx.getTableCells(shapes[0]!).flat();
      if (data) {
        const span = pptx.getTableCellSpan(cells[0]!);
        if (span.hMerge || span.vMerge) throw new Error('Edit the merged cell anchor.');
      }
      for (const cell of cells) pptx.setTableCellFill(cell, fill);
      return;
    }
    case 'table-style': {
      if (shapes.length !== 1 || !pptx.isTableShape(shapes[0]!))
        throw new Error('Select one table to change its style.');
      const flags = command.tableStyle;
      const keys = ['firstRow', 'lastRow', 'firstCol', 'lastCol', 'bandRow', 'bandCol'];
      if (
        !flags ||
        typeof flags !== 'object' ||
        Array.isArray(flags) ||
        !Object.keys(flags).length ||
        Object.entries(flags).some(
          ([key, value]) => !keys.includes(key) || typeof value !== 'boolean',
        )
      )
        throw new Error('Invalid table style options.');
      pptx.setTableStyleFlags(shapes[0]!, flags);
      return;
    }
    case 'table-cell-size':
    case 'table-distribute': {
      if (shapes.length !== 1 || !pptx.isTableShape(shapes[0]!))
        throw new Error('Select one table to distribute rows or columns.');
      if (command.tableAxis !== 'rows' && command.tableAxis !== 'columns')
        throw new Error('Invalid table distribution axis.');
      const table = shapes[0]!;
      const rows = command.tableAxis === 'rows';
      const sizes = rows ? pptx.getTableRowHeights(table) : pptx.getTableColumnWidths(table);
      let start = 0,
        count = sizes.length;
      let mergedTarget = !!command.cell;
      if (command.tableRange) {
        const targets = selectedTableCells(table, command.tableRange);
        start = rows ? command.tableRange.row : command.tableRange.column;
        count = rows ? command.tableRange.rows : command.tableRange.columns;
        mergedTarget = targets.length === 1;
      } else if (command.cell) {
        const { row, column } = command.cell;
        if (!Number.isInteger(row) || !Number.isInteger(column) || row < 0 || column < 0)
          throw new Error('Invalid table cell.');
        const cell = pptx.getTableCells(table)[row]?.[column];
        if (!cell) throw new Error('Invalid table cell.');
        const span = pptx.getTableCellSpan(cell);
        if (span.hMerge || span.vMerge) throw new Error('Select the merged cell anchor.');
        start = rows ? row : column;
        count = rows ? span.rowSpan : span.gridSpan;
      }
      if (count < 1 || start + count > sizes.length) throw new Error('Invalid table grid.');
      const total = sizes.slice(start, start + count).reduce((sum, value) => sum + value, 0);
      if (!Number.isSafeInteger(total) || total < count)
        throw new Error('Invalid table dimensions.');
      if (command.type === 'table-cell-size') {
        const requested = command.tableCellSize;
        if (
          typeof requested !== 'number' ||
          !Number.isFinite(requested) ||
          requested < 1 ||
          requested > 51206400
        )
          throw new Error('Invalid cell size.');
        const box = pptx.getShapeBoundsResolved(presentation, table);
        const sum = sizes.reduce((sum, value) => sum + value, 0);
        if (!box || sum <= 0) throw new Error('Table bounds are unavailable.');
        const extent = rows ? box.h : box.w;
        const next = sizes.map((value, index) => {
          if (index < start || index >= start + count) return Math.round((value * extent) / sum);
          return Math.round(mergedTarget ? (requested * value) / total : requested);
        });
        if (next.some((value) => value < 1)) throw new Error('Cell size is too small.');
        const nextExtent = next.reduce((sum, value) => sum + value, 0);
        const growth = nextExtent - extent;
        const angle = (pptx.getShapeRotation(table) * Math.PI) / 180;
        const expanded = bounds({
          x: box.x + ((rows ? -Math.sin(angle) : Math.cos(angle) - 1) * growth) / 2,
          y: box.y + ((rows ? Math.cos(angle) - 1 : Math.sin(angle)) * growth) / 2,
          w: rows ? box.w : nextExtent,
          h: rows ? nextExtent : box.h,
        });
        const set = rows ? pptx.setTableRowHeight : pptx.setTableColumnWidth;
        next.forEach((value, index) => set(table, index, pptx.emu(value)));
        pptx.setShapeBounds(table, expanded);
        return;
      }
      const set = rows ? pptx.setTableRowHeight : pptx.setTableColumnWidth;
      // Rounded boundaries preserve the exact total while differing by at most one EMU.
      for (let i = 0; i < count; i++)
        set(
          table,
          start + i,
          pptx.emu(Math.round((total * (i + 1)) / count) - Math.round((total * i) / count)),
        );
      fitTableText(presentation, table, rows ? { start, count } : undefined);
      return;
    }
    case 'table-clear': {
      if (shapes.length !== 1 || !pptx.isTableShape(shapes[0]!))
        throw new Error('Select cells in one table to clear.');
      const table = shapes[0]!,
        range = command.tableRange;
      const targets = selectedTableCells(table, range);
      for (const { row, column } of targets) {
        const cell = pptx.getTableCell(table, row, column);
        pptx.replaceTableCellTextRange(cell, 0, pptx.getTableCellText(cell).length, '');
      }
      return;
    }
    case 'table-split': {
      if (shapes.length !== 1 || !pptx.isTableShape(shapes[0]!))
        throw new Error('Select one table cell to split.');
      if (!command.cell || !command.tableSplit)
        throw new Error('Select a cell and split dimensions.');
      pptx.splitTableCell(shapes[0]!, command.cell.row, command.cell.column, command.tableSplit);
      return;
    }
    case 'table-merge': {
      if (shapes.length !== 1 || !pptx.isTableShape(shapes[0]!))
        throw new Error('Select cells in one table to merge.');
      const range = command.tableRange;
      if (!range) throw new Error('Select a rectangular range of cells.');
      pptx.mergeTableCells(
        shapes[0]!,
        { row: range.row, col: range.column, rowSpan: range.rows, colSpan: range.columns },
        { coveredText: 'append', allowContainedMerges: true },
      );
      return;
    }
    case 'table-delete': {
      if (shapes.length !== 1 || !pptx.isTableShape(shapes[0]!))
        throw new Error('Select one table to delete rows or columns.');
      if (!['rows', 'columns'].includes(command.tableAxis ?? ''))
        throw new Error('Invalid table deletion axis.');
      const table = shapes[0]!;
      const rows = command.tableAxis === 'rows';
      const sizes = rows ? pptx.getTableRowHeights(table) : pptx.getTableColumnWidths(table);
      if (!sizes.length || sizes.some((size) => !Number.isFinite(size) || size <= 0))
        throw new Error('Invalid table grid.');
      let start = 0,
        count = sizes.length;
      if (command.tableRange) {
        const range = command.tableRange;
        if (
          ![range.row, range.column, range.rows, range.columns].every(Number.isInteger) ||
          range.row < 0 ||
          range.column < 0 ||
          range.rows < 1 ||
          range.columns < 1 ||
          range.row + range.rows > pptx.getTableRowHeights(table).length ||
          range.column + range.columns > pptx.getTableColumnWidths(table).length
        )
          throw new Error('Invalid table cell range.');
        start = rows ? range.row : range.column;
        count = rows ? range.rows : range.columns;
      } else if (command.cell) {
        const { row, column } = command.cell;
        if (!Number.isInteger(row) || !Number.isInteger(column) || row < 0 || column < 0)
          throw new Error('Invalid table cell.');
        const span = pptx.getTableCellSpan(pptx.getTableCell(table, row, column));
        if (span.hMerge || span.vMerge) throw new Error('Edit the merged cell anchor.');
        start = rows ? row : column;
        count = rows ? span.rowSpan : span.gridSpan;
      }
      if (start + count > sizes.length) throw new Error('Invalid table cell span.');
      if (count === sizes.length) {
        pptx.removeShape(table);
        return;
      }
      const box = pptx.getShapeBoundsResolved(presentation, table);
      if (!box) throw new Error('Table bounds are unavailable.');
      const total = sizes.reduce((sum, size) => sum + size, 0);
      const extent = rows ? box.h : box.w;
      let cumulative = 0;
      const displayed = sizes.map((size) => {
        const previous = Math.round((cumulative * extent) / total);
        cumulative += size;
        return Math.round((cumulative * extent) / total) - previous;
      });
      if (displayed.some((size) => size < 1)) throw new Error('Table cells are too small.');
      const growth = -displayed.slice(start, start + count).reduce((sum, size) => sum + size, 0);
      const angle = (pptx.getShapeRotation(table) * Math.PI) / 180;
      const reduced = bounds({
        x: box.x + ((rows ? -Math.sin(angle) : Math.cos(angle) - 1) * growth) / 2,
        y: box.y + ((rows ? Math.cos(angle) - 1 : Math.sin(angle)) * growth) / 2,
        w: box.w + (rows ? 0 : growth),
        h: box.h + (rows ? growth : 0),
      });
      const set = rows ? pptx.setTableRowHeight : pptx.setTableColumnWidth;
      displayed.forEach((size, i) => set(table, i, pptx.emu(size)));
      const remove = rows ? pptx.removeTableRow : pptx.removeTableColumn;
      for (let i = 0; i < count; i++) remove(table, start);
      pptx.setShapeBounds(table, reduced);
      return;
    }
    case 'table-insert': {
      if (shapes.length !== 1 || !pptx.isTableShape(shapes[0]!))
        throw new Error('Select one table to insert rows or columns.');
      const placement = command.tablePlacement;
      if (!placement || !['above', 'below', 'left', 'right'].includes(placement))
        throw new Error('Invalid table insertion position.');
      const table = shapes[0]!;
      const rows = placement === 'above' || placement === 'below';
      const after = placement === 'below' || placement === 'right';
      const sizes = rows ? pptx.getTableRowHeights(table) : pptx.getTableColumnWidths(table);
      if (
        !sizes.length ||
        sizes.length >= 10000 ||
        sizes.some((size) => !Number.isFinite(size) || size <= 0)
      )
        throw new Error('Invalid table grid.');
      let index = after ? sizes.length : 0;
      let count = 1;
      const cell = command.cell;
      if (command.tableRange) {
        const range = command.tableRange;
        const dimensions = pptx.getTableDimensions(table);
        if (
          ![range.row, range.column, range.rows, range.columns].every(Number.isInteger) ||
          range.row < 0 ||
          range.column < 0 ||
          range.rows < 1 ||
          range.columns < 1 ||
          range.row + range.rows > dimensions.rows ||
          range.column + range.columns > dimensions.cols
        )
          throw new Error('Invalid table cell range.');
        count = rows ? range.rows : range.columns;
        index = (rows ? range.row : range.column) + (after ? count : 0);
      } else if (cell) {
        if (
          !Number.isInteger(cell.row) ||
          !Number.isInteger(cell.column) ||
          cell.row < 0 ||
          cell.column < 0
        )
          throw new Error('Invalid table cell.');
        const span = pptx.getTableCellSpan(pptx.getTableCell(table, cell.row, cell.column));
        if (span.hMerge || span.vMerge) throw new Error('Edit the merged cell anchor.');
        index =
          (rows ? cell.row : cell.column) + (after ? (rows ? span.rowSpan : span.gridSpan) : 0);
      }
      if (sizes.length + count > 10000) throw new Error('Table grid is too large.');
      const box = pptx.getShapeBoundsResolved(presentation, table);
      if (!box) throw new Error('Table bounds are unavailable.');
      const total = sizes.reduce((sum, size) => sum + size, 0);
      const extent = rows ? box.h : box.w;
      let cumulative = 0;
      const displayed = sizes.map((size) => {
        const previous = Math.round((cumulative * extent) / total);
        cumulative += size;
        return Math.round((cumulative * extent) / total) - previous;
      });
      const insertedSize = displayed[Math.min(after ? index - 1 : index, displayed.length - 1)]!;
      const growth = insertedSize * count;
      const angle = (pptx.getShapeRotation(table) * Math.PI) / 180;
      const expanded = bounds({
        x: box.x + ((rows ? -Math.sin(angle) : Math.cos(angle) - 1) * growth) / 2,
        y: box.y + ((rows ? Math.cos(angle) - 1 : Math.sin(angle)) * growth) / 2,
        w: box.w + (rows ? 0 : growth),
        h: box.h + (rows ? growth : 0),
      });
      if (displayed.some((size) => size < 1)) throw new Error('Table cells are too small.');
      const set = rows ? pptx.setTableRowHeight : pptx.setTableColumnWidth;
      displayed.forEach((size, i) => set(table, i, pptx.emu(size)));
      for (let offset = 0; offset < count; offset++) {
        if (rows) {
          pptx.insertTableRow(table, index + offset);
          pptx.setTableRowHeight(table, index + offset, pptx.emu(insertedSize));
        } else pptx.insertTableColumn(table, index + offset, pptx.emu(insertedSize));
      }
      pptx.setShapeBounds(table, expanded);
      return;
    }
    case 'table-row-append': {
      if (shapes.length !== 1 || !pptx.isTableShape(shapes[0]!))
        throw new Error('Select one table to add a row.');
      const table = shapes[0]!;
      const heights = pptx.getTableRowHeights(table);
      const dimensions = pptx.getTableDimensions(table);
      if (!heights.length || !dimensions.cols || heights.length >= 10000)
        throw new Error('Cannot add another row to this table.');
      const box = pptx.getShapeBoundsResolved(presentation, table);
      if (!box) throw new Error('Table bounds are unavailable.');
      const height = heights.at(-1)!;
      const total = heights.reduce((sum, value) => sum + value, 0);
      const growth = total > 0 ? (box.h * height) / total : height;
      const angle = (pptx.getShapeRotation(table) * Math.PI) / 180;
      // Keep the rotated table's top edge fixed while extending its bottom edge.
      const expanded = bounds({
        x: box.x - (Math.sin(angle) * growth) / 2,
        y: box.y + ((Math.cos(angle) - 1) * growth) / 2,
        w: box.w,
        h: box.h + growth,
      });
      pptx.insertTableRow(table);
      pptx.setTableRowHeight(table, heights.length, height);
      pptx.setShapeBounds(table, expanded);
      return;
    }

    case 'table-range-format': {
      if (shapes.length !== 1 || !pptx.isTableShape(shapes[0]!))
        throw new Error('Select cells in one table to format.');
      const table = shapes[0]!,
        targets = selectedTableCells(table, command.tableRange);
      const changes = command.changes;
      if (
        !changes ||
        !Object.keys(changes).length ||
        Object.keys(changes).some(
          (key) =>
            ![
              'format',
              'align',
              'bullets',
              'lineSpacing',
              'anchor',
              'direction',
              'paragraph',
              'indent',
              'changeCase',
            ].includes(key),
        )
      )
        throw new Error('Invalid table range format.');
      if (changes.indent !== undefined && ![-1, 1].includes(changes.indent))
        throw new Error('Invalid indentation.');
      const format = changes.format ? editorFormat(changes.format) : null;
      if (
        changes.align !== undefined &&
        !['left', 'center', 'right', 'justify', 'distribute'].includes(changes.align)
      )
        throw new Error('Invalid table cell alignment.');
      if (changes.bullets !== undefined && !['bullet', 'number', 'none'].includes(changes.bullets))
        throw new Error('Invalid table cell bullets.');
      if (changes.lineSpacing !== undefined) finite(changes.lineSpacing, 0.1, 10);
      if (changes.anchor !== undefined && !['top', 'center', 'bottom'].includes(changes.anchor))
        throw new Error('Invalid table cell anchor.');
      if (
        changes.direction !== undefined &&
        !['horz', 'vert', 'vert270', 'wordArtVert'].includes(changes.direction)
      )
        throw new Error('Invalid table cell direction.');
      for (const { row, column } of targets) {
        const cell = pptx.getTableCell(table, row, column),
          end = pptx.getTableCellText(cell).length;
        if (changes.changeCase !== undefined)
          for (const edit of textCaseEdits(pptx.getTableCellText(cell), changes.changeCase))
            pptx.replaceTableCellTextRange(cell, edit.start, edit.end, edit.text);
        if (format) pptx.setTableCellTextFormat(cell, format);
        if (changes.indent !== undefined)
          pptx.shiftTableCellTextRangeLevel(cell, 0, end, changes.indent);
        if (changes.paragraph)
          pptx.setTableCellTextRangeParagraphSettings(cell, 0, end, changes.paragraph);
        if (changes.align !== undefined)
          pptx.setTableCellTextRangeAlignment(cell, 0, end, changes.align);
        if (changes.bullets !== undefined)
          pptx.setTableCellTextRangeBullets(cell, 0, end, changes.bullets);
        if (changes.lineSpacing !== undefined)
          pptx.setTableCellTextRangeLineSpacing(cell, 0, end, {
            kind: 'pct',
            value: changes.lineSpacing,
          });
        if (changes.anchor !== undefined) pptx.setTableCellAnchor(cell, changes.anchor);
        if (changes.direction !== undefined)
          pptx.setTableCellTextDirection(cell, changes.direction);
      }
      return;
    }
    case 'table-cell-layout':
    case 'table-cell-bullets':
    case 'table-cell-case':
    case 'table-cell-indent':
    case 'table-cell-paragraph':
    case 'table-cell-line-spacing':
    case 'table-cell-align':
    case 'table-cell-format':
    case 'table-cell-text': {
      const data = command.cell;
      if (
        shapes.length !== 1 ||
        !pptx.isTableShape(shapes[0]!) ||
        !data ||
        !Number.isInteger(data.row) ||
        !Number.isInteger(data.column) ||
        data.row < 0 ||
        data.column < 0 ||
        !Array.isArray(data.edits) ||
        data.edits.length > 10000
      )
        throw new Error('Invalid table cell edit.');
      const cell = pptx.getTableCell(shapes[0]!, data.row, data.column);
      const span = pptx.getTableCellSpan(cell);
      if (span.hMerge || span.vMerge) throw new Error('Edit the merged cell anchor.');
      if (command.type === 'table-cell-layout') {
        const changes = command.changes;
        if (
          !changes ||
          (changes.anchor === undefined && changes.direction === undefined) ||
          (changes.anchor !== undefined && !['top', 'center', 'bottom'].includes(changes.anchor)) ||
          (changes.direction !== undefined &&
            !['horz', 'vert', 'vert270', 'wordArtVert'].includes(changes.direction))
        )
          throw new Error('Invalid table cell text layout.');
        // Validate both properties before changing either one.
        if (changes.anchor !== undefined) pptx.setTableCellAnchor(cell, changes.anchor);
        if (changes.direction !== undefined)
          pptx.setTableCellTextDirection(cell, changes.direction);
        return;
      }
      if (command.type === 'table-cell-bullets') {
        const changes = command.changes;
        if (
          !changes?.range ||
          !changes.bullets ||
          !['bullet', 'number', 'none'].includes(changes.bullets)
        )
          throw new Error('Invalid table cell bullets.');
        pptx.setTableCellTextRangeBullets(
          cell,
          changes.range.start,
          changes.range.end,
          changes.bullets,
        );
        return;
      }
      if (command.type === 'table-cell-case') {
        const changes = command.changes;
        if (!changes?.range || !changes.changeCase) throw new Error('Invalid text case.');
        for (const edit of textCaseEdits(
          pptx.getTableCellText(cell),
          changes.changeCase,
          changes.range,
        ))
          pptx.replaceTableCellTextRange(cell, edit.start, edit.end, edit.text);
        return;
      }
      if (command.type === 'table-cell-indent') {
        const changes = command.changes;
        if (!changes?.range || changes.indent === undefined)
          throw new Error('Invalid indentation.');
        pptx.shiftTableCellTextRangeLevel(
          cell,
          changes.range.start,
          changes.range.end,
          changes.indent,
        );
        return;
      }
      if (command.type === 'table-cell-paragraph') {
        const changes = command.changes;
        if (!changes?.range || !changes.paragraph) throw new Error('Invalid paragraph settings.');
        pptx.setTableCellTextRangeParagraphSettings(
          cell,
          changes.range.start,
          changes.range.end,
          changes.paragraph,
        );
        return;
      }
      if (command.type === 'table-cell-line-spacing') {
        const changes = command.changes;
        if (!changes?.range || changes.lineSpacing === undefined)
          throw new Error('Invalid table cell line spacing.');
        finite(changes.lineSpacing, 0.1, 10);
        pptx.setTableCellTextRangeLineSpacing(cell, changes.range.start, changes.range.end, {
          kind: 'pct',
          value: changes.lineSpacing,
        });
        return;
      }
      if (command.type === 'table-cell-align') {
        const changes = command.changes;
        if (
          !changes?.range ||
          !changes.align ||
          !['left', 'center', 'right', 'justify', 'distribute'].includes(changes.align)
        )
          throw new Error('Invalid table cell alignment.');
        pptx.setTableCellTextRangeAlignment(
          cell,
          changes.range.start,
          changes.range.end,
          changes.align,
        );
        return;
      }
      if (command.type === 'table-cell-format') {
        if (!command.changes?.format || !command.changes.range)
          throw new Error('Invalid table cell format.');
        const format = editorFormat(command.changes.format);
        const { start, end } = command.changes.range;
        pptx.setTableCellTextRangeFormat(cell, start, end, format);
        return;
      }
      // Validate the full sequence before applying any change.
      let value = pptx.getTableCellText(cell);
      const formats = data.edits.map((edit) => (edit.format ? editorFormat(edit.format) : null));
      for (const edit of data.edits) {
        text(edit.text);
        validateTextRange(value, edit.start, edit.end);
        value =
          value.slice(0, edit.start) + edit.text.replace(/\r\n?/g, '\n') + value.slice(edit.end);
        text(value);
      }
      for (const [index, edit] of data.edits.entries()) {
        pptx.replaceTableCellTextRange(cell, edit.start, edit.end, edit.text);
        const format = formats[index];
        if (format && edit.text.length)
          pptx.setTableCellTextRangeFormat(
            cell,
            edit.start,
            edit.start + edit.text.replace(/\r\n?/g, '\n').length,
            format,
          );
      }
      return;
    }
    case 'slide-layout': {
      const layout = command.layout && pptx.findSlideLayoutByPartName(presentation, command.layout);
      if (!layout) throw new Error('Slide layout no longer exists.');
      pptx.applySlideLayout(presentation, slide, layout);
      return;
    }
    case 'slide-reset':
      pptx.resetSlideLayout(presentation, slide);
      return;
    case 'slide-delete':
      pptx.removeSlide(presentation, slide);
      return;
    case 'slide-duplicate':
      pptx.duplicateSlideAt(presentation, command.slide + 1, slide);
      return;
    case 'slide-move':
      finite(command.to, 0, slides.length - 1);
      if (!Number.isInteger(command.to)) throw new Error('Invalid slide position.');
      pptx.moveSlide(presentation, slide, command.to);
      return;
    case 'transition-apply-all':
      pptx.applySlideTransitionToAll(presentation, slide);
      return;
    case 'transition-sound': {
      const value = command.transitionSound;
      if (!value || !['none', 'stop', 'play'].includes(value.kind))
        throw new Error('Invalid transition sound.');
      if (value.kind === 'play') {
        if (typeof value.loop !== 'boolean')
          throw new Error('Invalid transition sound loop setting.');
        const resolved = resolveActionSound({
          sound: { name: value.name, base64: value.base64 },
          stopPrevious: false,
        });
        pptx.setSlideTransitionSound(slide, { kind: 'play', ...resolved.sound!, loop: value.loop });
      } else pptx.setSlideTransitionSound(slide, value);
      return;
    }
    case 'transition-duration':
      if (command.transitionDurationMs === undefined)
        throw new Error('Missing transition duration.');
      pptx.setSlideTransitionDuration(slide, command.transitionDurationMs);
      return;
    case 'transition-effect':
      if (!command.transitionEffect) throw new Error('Missing transition effect.');
      pptx.setSlideTransitionEffect(slide, command.transitionEffect);
      return;
    case 'slide-advance':
      if (!command.advanceTiming) throw new Error('Missing slide advance timing.');
      pptx.setSlideAdvanceTiming(slide, command.advanceTiming);
      return;
    case 'slide-hidden':
      if (typeof command.hidden !== 'boolean') throw new Error('Invalid hidden state.');
      pptx.setSlideHidden(slide, command.hidden);
      return;
    case 'notes':
      text(command.text);
      pptx.setSlideNotes(slide, command.text);
      return;
    case 'insert': {
      const line = lineTool(command.preset);
      if (line) {
        if (!command.from || !command.toPoint) throw new Error('Line endpoints are required.');
        for (const point of [command.from, command.toPoint]) {
          finite(point.x, -1e9, 1e9);
          finite(point.y, -1e9, 1e9);
        }
        const shape = pptx.addSlideLine(slide, {
          from: { x: pptx.emu(command.from.x), y: pptx.emu(command.from.y) },
          to: { x: pptx.emu(command.toPoint.x), y: pptx.emu(command.toPoint.y) },
          preset: line.geometry,
          name: line.label,
          color: '4472C4',
          widthEmu: pptx.pt(1),
        });
        if (command.connections) setConnections(presentation, slide, shape, command.connections);
        if (line.arrows > 0) pptx.setShapeStrokeArrow(shape, 'tail', { type: 'triangle' });
        if (line.arrows > 1) pptx.setShapeStrokeArrow(shape, 'head', { type: 'triangle' });
        return;
      }
      const box = bounds(
        command.changes?.bounds ?? {
          x: pptx.inches(1),
          y: pptx.inches(1),
          w: pptx.inches(3),
          h: pptx.inches(1),
        },
      );
      let shape: pptx.SlideShapeData;
      if (command.preset === 'table') {
        const table = command.table;
        if (
          !table ||
          !Number.isInteger(table.rows) ||
          !Number.isInteger(table.columns) ||
          table.rows < 1 ||
          table.rows > 75 ||
          table.columns < 1 ||
          table.columns > 75
        )
          throw new Error('Table dimensions must be integers from 1 to 75.');
        pptx.addSlideTable(slide, {
          ...box,
          rows: Array.from({ length: table.rows }, () => Array<string>(table.columns).fill('')),
          firstRow: true,
          bandRow: true,
        });
      } else if (command.preset === 'picture') {
        const image = command.image;
        if (
          !image ||
          typeof image.base64 !== 'string' ||
          image.base64.length > 28_000_000 ||
          image.base64.length % 4 !== 0 ||
          !/^[A-Za-z0-9+/]*={0,2}$/.test(image.base64)
        )
          throw new Error('Invalid picture data.');
        text(image.name);
        const bytes = Buffer.from(image.base64, 'base64');
        if (!bytes.length || bytes.length > 20_000_000)
          throw new Error('Picture must be under 20 MB.');
        pptx.addSlideImage(slide, bytes, { ...box, name: image.name, fit: 'contain' });
      } else if (command.preset === 'text') {
        text(command.text ?? '');
        shape = pptx.addSlideTextBox(slide, { ...box, text: command.text ?? '' });
        pptx.setShapeTextFormat(shape, { size: 18, color: '000000' });
      } else {
        if (!shapePresets.has(command.preset ?? '')) throw new Error('Invalid shape preset.');
        shape = pptx.addSlideShape(slide, { ...box, preset: command.preset!, text: '' });
        pptx.setShapeFill(shape, '4472C4');
        pptx.setShapeStroke(shape, { color: '2F5597', widthEmu: pptx.pt(1) });
        const action = defaultButtonAction(command.preset!);
        if (action) pptx.setShapeClickAction(shape, { kind: action });
      }
      return;
    }
    case 'text-link':
    case 'object-actions':
    case 'object-link': {
      let linkedCell: pptx.TableCellData | undefined;
      if (command.type === 'text-link') {
        if (shapes.length !== 1 || !command.range) throw new Error('Invalid text link selection.');
        if (command.cell) {
          const { row, column } = command.cell;
          if (
            !pptx.isTableShape(shapes[0]!) ||
            !Number.isInteger(row) ||
            !Number.isInteger(column) ||
            row < 0 ||
            column < 0
          )
            throw new Error('Invalid table cell link.');
          linkedCell = pptx.getTableCell(shapes[0]!, row, column);
          const span = pptx.getTableCellSpan(linkedCell);
          if (span.hMerge || span.vMerge) throw new Error('Edit the merged cell anchor.');
        } else if (pptx.getShapeKind(shapes[0]!) !== 'shape')
          throw new Error('Invalid text link selection.');
        validateTextRange(
          linkedCell ? pptx.getTableCellText(linkedCell) : pptx.getShapeText(shapes[0]!),
          command.range.start,
          command.range.end,
        );
        if (command.range.start === command.range.end && (!command.text || !command.link))
          throw new Error('Enter display text and a link to insert.');
        if (command.text !== undefined) {
          text(command.text);
          if (!command.text.length) throw new Error('Enter display text.');
        }
      }
      if (!shapes.length || command.link === undefined) throw new Error('Invalid link selection.');
      const link = command.link;
      const action = resolveActionLink(presentation, link);
      const hoverAction =
        command.type === 'object-actions'
          ? resolveActionLink(presentation, command.hoverLink)
          : null;
      const sounds =
        command.type === 'object-actions' && command.actionSounds !== undefined
          ? {
              click: resolveActionSound(command.actionSounds?.click),
              hover: resolveActionSound(command.actionSounds?.hover),
            }
          : undefined;
      if (command.type === 'text-link') {
        const start = command.range!.start;
        let end = command.range!.end;
        if (command.text !== undefined) {
          const value = linkedCell
            ? pptx.getTableCellText(linkedCell)
            : pptx.getShapeText(shapes[0]!);
          if (value.slice(start, end) !== command.text) {
            if (linkedCell) pptx.replaceTableCellTextRange(linkedCell, start, end, command.text);
            else pptx.replaceShapeTextRange(shapes[0]!, start, end, command.text);
          }
          end = start + command.text.length;
        }
        if (linkedCell)
          pptx.setTableCellTextRangeClickAction(
            linkedCell,
            start,
            end,
            action,
            link?.tooltip ?? undefined,
          );
        else
          pptx.setShapeTextRangeClickAction(
            shapes[0]!,
            start,
            end,
            action,
            link?.tooltip ?? undefined,
          );
      } else
        for (const shape of shapes) {
          const priorSounds =
            command.type === 'object-actions'
              ? {
                  click: pptx.getShapeActionSound(shape, 'click'),
                  hover: pptx.getShapeActionSound(shape, 'hover'),
                }
              : undefined;
          if (
            command.type !== 'object-actions' ||
            JSON.stringify(objectLink(shape)) !== JSON.stringify(link)
          )
            pptx.setShapeClickAction(shape, action, link?.tooltip ?? undefined);
          if (
            command.type === 'object-actions' &&
            JSON.stringify(objectLink(shape, true)) !== JSON.stringify(command.hoverLink)
          )
            pptx.setShapeHoverAction(shape, hoverAction, command.hoverLink?.tooltip ?? undefined);
          if (priorSounds) {
            for (const trigger of ['click', 'hover'] as const) {
              const value = sounds?.[trigger] ?? priorSounds[trigger];
              // Retain sound-only settings when the navigation action is removed.
              const current = pptx.getShapeActionSound(shape, trigger);
              const equal =
                current.stopPrevious === value.stopPrevious &&
                ((current.sound === null && value.sound === null) ||
                  (current.sound !== null &&
                    value.sound !== null &&
                    current.sound.name === value.sound.name &&
                    Buffer.from(current.sound.bytes).equals(value.sound.bytes)));
              if (!equal) pptx.setShapeActionSound(shape, trigger, value);
            }
          }
        }
      return;
    }
    case 'picture-replace': {
      if (!shapes.length || shapes.some((shape) => pptx.getShapeKind(shape) !== 'picture'))
        throw new Error('Select a picture to replace.');
      const image = command.image;
      if (
        !image ||
        typeof image.base64 !== 'string' ||
        image.base64.length > 28_000_000 ||
        image.base64.length % 4 !== 0 ||
        !/^[A-Za-z0-9+/]*={0,2}$/.test(image.base64)
      )
        throw new Error('Invalid picture data.');
      const bytes = Buffer.from(image.base64, 'base64');
      if (!bytes.length || bytes.length > 20_000_000)
        throw new Error('Picture must be under 20 MB.');
      for (const shape of shapes) pptx.setShapeImage(shape, bytes, { isolated: true });
      return;
    }
    case 'delete':
      for (const id of ids) pptx.removeShape(target(id));
      return;
    case 'paste': {
      let copiedPresentation = presentation;
      if (command.snapshot !== undefined) {
        if (
          typeof command.snapshot !== 'string' ||
          command.snapshot.length > 26_000_000 ||
          !/^[A-Za-z0-9+/]+={0,2}$/.test(command.snapshot)
        )
          throw new Error('Invalid clipboard snapshot.');
        if (command.cut) throw new Error('Snapshot paste cannot delete source objects.');
        copiedPresentation = pptx.loadPresentationBytes(Buffer.from(command.snapshot, 'base64'));
      }
      const source = pptx.getSlides(copiedPresentation)[command.source ?? -1];
      if (!source || !Number.isInteger(command.source))
        throw new Error('Copied slide no longer exists.');
      if (command.pasteSlideCoordinates) {
        if (command.snapshot === undefined)
          throw new Error('Slide-coordinate paste requires a clipboard snapshot.');
        const selectedIds = new Set(ids);
        const containsSelected = (shape: pptx.SlideShapeData): boolean =>
          pptx
            .getGroupChildren(shape)
            .some((child) => selectedIds.has(pptx.getShapeId(child)) || containsSelected(child));
        // Only the private clipboard presentation is transformed. Release ancestors
        // from outside in, retaining selected groups and the original source document.
        for (;;) {
          const ancestor = pptx.getSlideShapes(source).find(containsSelected);
          if (!ancestor) break;
          pptx.ungroupShapes(ancestor);
        }
      }
      const copies = new Map<number, pptx.SlideShapeData>();
      const originals = new Map<number, EditorShape['connections']>();
      const sourceShapes = pptx.getSlideShapes(source);
      const selectedSources = ids.map((id) => {
        const original = sourceShapes.find((shape) => pptx.getShapeId(shape) === id);
        if (!original) throw new Error('Copied shape no longer exists.');
        return original;
      });
      // New clipboard captures preserve stacking order regardless of click order.
      // Keep legacy live-source commands deterministic when replaying old journals.
      if (command.snapshot !== undefined)
        selectedSources.sort((a, b) => sourceShapes.indexOf(a) - sourceShapes.indexOf(b));
      for (const original of selectedSources) {
        const copied =
          command.snapshot === undefined
            ? pptx.copyShape(slide, original)
            : pptx.importShape(slide, original);
        registerCopy(original, copied, copies, originals);
        const box = pptx.getShapeBoundsResolved(presentation, copied);
        if (box && (source === slide || command.pasteOffset))
          pptx.setShapePosition(
            copied,
            pptx.emu(box.x + pptx.pt(12)),
            pptx.emu(box.y + pptx.pt(12)),
          );
        if (command.cut) pptx.removeShape(original);
      }
      reconnectCopies(copies, originals);
      return;
    }
    case 'duplicate': {
      const copies = new Map<number, pptx.SlideShapeData>();
      const originals = new Map<number, EditorShape['connections']>();
      const contexts = command.duplicateInParent
        ? slideGeometryContexts(presentation, slide)
        : null;
      const sources = ids.map(target);
      if (contexts) {
        const all = pptx.getSlideShapes(slide);
        sources.sort((a, b) => all.indexOf(a) - all.indexOf(b));
        if (sources.some((s) => !contexts.has(pptx.getShapeId(s))))
          throw new Error('Cannot duplicate within a group with a singular transform.');
      }
      for (const original of sources) {
        const copied = pptx.copyShape(slide, original, { sameParent: !!contexts });
        registerCopy(original, copied, copies, originals);
        const box = pptx.getShapeBoundsResolved(presentation, copied);
        if (box) {
          const context = contexts?.get(pptx.getShapeId(original));
          const origin = context?.project({ x: box.x, y: box.y });
          const position =
            context && origin
              ? context.unproject({ x: origin.x + pptx.pt(12), y: origin.y + pptx.pt(12) })
              : { x: box.x + pptx.pt(12), y: box.y + pptx.pt(12) };
          pptx.setShapePosition(copied, pptx.emu(position.x), pptx.emu(position.y));
        }
      }
      reconnectCopies(copies, originals);
      return;
    }
    case 'group':
      pptx.groupShapes(shapes);
      return;
    case 'regroup':
      pptx.regroupShapes(shapes);
      return;
    case 'ungroup':
      for (const id of ids) pptx.ungroupShapes(target(id));
      return;
    case 'order': {
      if (command.orderPlacement !== undefined) {
        const { target: targetId, side } = command.orderPlacement;
        if (!Number.isInteger(targetId) || !['before', 'after'].includes(side))
          throw new Error('Invalid layer placement.');
        const all = pptx.getSlideShapes(slide);
        const children = all
          .filter((shape) => pptx.getShapeKind(shape) === 'group')
          .map(pptx.getGroupChildren);
        const descendants = new Set(children.flat().map(pptx.getShapeId));
        const containers = [
          all.filter((shape) => !descendants.has(pptx.getShapeId(shape))),
          ...children,
        ];
        const siblings = containers.find((items) =>
          items.some((shape) => pptx.getShapeId(shape) === targetId),
        );
        if (
          !ids.length ||
          ids.includes(targetId) ||
          !siblings ||
          ids.some((id) => !siblings.some((shape) => pptx.getShapeId(shape) === id))
        )
          throw new Error('Layer placement requires siblings and an unselected target.');
        const chosen = new Set(ids);
        const moving = siblings.filter((shape) => chosen.has(pptx.getShapeId(shape)));
        const ordered = siblings.filter((shape) => !chosen.has(pptx.getShapeId(shape)));
        const index = ordered.findIndex((shape) => pptx.getShapeId(shape) === targetId);
        ordered.splice(index + (side === 'after' ? 1 : 0), 0, ...moving);
        ordered.forEach((shape, index) => {
          if (pptx.getShapeZIndex(shape) !== index) pptx.setShapeZIndex(shape, index);
        });
        return;
      }
      const operations = {
        front: pptx.bringShapeToFront,
        back: pptx.sendShapeToBack,
        forward: pptx.bringShapeForward,
        backward: pptx.sendShapeBackward,
      };
      if (!command.order || !Object.hasOwn(operations, command.order))
        throw new Error('Invalid layer order.');
      if (command.stableOrder) {
        const selected = new Set(ids);
        const all = pptx.getSlideShapes(slide);
        const groups = all.filter((s) => pptx.getShapeKind(s) === 'group');
        const children = groups.map(pptx.getGroupChildren);
        const descendants = new Set(children.flat().map(pptx.getShapeId));
        const containers = [all.filter((s) => !descendants.has(pptx.getShapeId(s))), ...children];
        for (const siblings of containers) {
          const chosen = (s: pptx.SlideShapeData) => selected.has(pptx.getShapeId(s));
          if (!siblings.some(chosen)) continue;
          let ordered = [...siblings];
          if (command.order === 'front')
            ordered = [...siblings.filter((s) => !chosen(s)), ...siblings.filter(chosen)];
          else if (command.order === 'back')
            ordered = [...siblings.filter(chosen), ...siblings.filter((s) => !chosen(s))];
          else if (command.order === 'forward') {
            for (let i = ordered.length - 2; i >= 0; i--)
              if (chosen(ordered[i]!) && !chosen(ordered[i + 1]!))
                [ordered[i], ordered[i + 1]] = [ordered[i + 1]!, ordered[i]!];
          } else {
            for (let i = 1; i < ordered.length; i++)
              if (chosen(ordered[i]!) && !chosen(ordered[i - 1]!))
                [ordered[i], ordered[i - 1]] = [ordered[i - 1]!, ordered[i]!];
          }
          ordered.forEach((shape, index) => {
            if (pptx.getShapeZIndex(shape) !== index) pptx.setShapeZIndex(shape, index);
          });
        }
        return;
      }
      for (const id of command.order === 'back' ? [...ids].reverse() : ids)
        operations[command.order](target(id));
      return;
    }
    case 'update': {
      if (
        command.tableResizeAnchor !== undefined &&
        !['top', 'bottom'].includes(command.tableResizeAnchor)
      )
        throw new Error('Invalid table resize anchor.');
      const tableBounds = new Map(
        shapes
          .filter(pptx.isTableShape)
          .map((shape) => [
            pptx.getShapeId(shape),
            pptx.getShapeBoundsResolved(presentation, shape),
          ]),
      );
      const changes = command.changes ?? {};
      if (changes.hidden !== undefined) {
        if (typeof changes.hidden !== 'boolean') throw new Error('Invalid shape visibility.');
        for (const shape of shapes) pptx.setShapeHidden(shape, changes.hidden);
      }
      if (changes.name !== undefined) {
        text(changes.name);
        for (const shape of shapes) pptx.renameShape(shape, changes.name);
      }
      if (command.textFits !== undefined) {
        if (!Array.isArray(command.textFits) || command.textFits.length > ids.length)
          throw new Error('Invalid text fit measurements.');
        const seen = new Set<number>();
        for (const fit of command.textFits) {
          if (!ids.includes(fit.id) || seen.has(fit.id))
            throw new Error('Invalid text fit target.');
          seen.add(fit.id);
        }
      }

      if (changes.rotationDelta !== undefined) {
        finite(changes.rotationDelta, -36000, 36000);
        if (changes.rotation !== undefined)
          throw new Error('Cannot combine rotation and rotation delta.');
      }
      if (command.positions)
        for (const item of command.positions) {
          if (!ids.includes(item.id)) throw new Error('Position is outside selection.');
          const moving = target(item.id);
          if (pptx.getShapeKind(moving) === 'connector') {
            for (const end of ['start', 'end'] as const) {
              const attachment = pptx.getShapeConnection(moving, end);
              if (attachment && !ids.includes(attachment.shapeId))
                pptx.setShapeConnection(moving, end, null);
            }
          }
          pptx.setShapeBounds(
            target(item.id),
            bounds(item.bounds, pptx.getShapeKind(target(item.id)) === 'connector'),
          );
        }
      for (const id of ids) {
        const shape = target(id);
        if (changes.imageBrightness !== undefined) finite(changes.imageBrightness, -1, 1);
        if (changes.imageContrast !== undefined) finite(changes.imageContrast, -1, 1);
        if (changes.imageBrightness !== undefined)
          pptx.setShapeImageBrightness(shape, changes.imageBrightness);
        if (changes.imageContrast !== undefined)
          pptx.setShapeImageContrast(shape, changes.imageContrast);
        if (changes.imageOpacity !== undefined) {
          finite(changes.imageOpacity, 0, 1);
          pptx.setShapeImageOpacity(shape, changes.imageOpacity);
        }
        if (changes.imageCrop !== undefined) pptx.setShapeImageCrop(shape, changes.imageCrop);
        if (changes.imageCropAspectRatio !== undefined)
          pptx.setShapeImageCropAspectRatio(shape, changes.imageCropAspectRatio);
        if (changes.imageFit !== undefined) pptx.setShapeImageFit(shape, changes.imageFit);
        if (changes.imageCropShape !== undefined) {
          if (!shapePresets.has(changes.imageCropShape)) throw new Error('Invalid crop shape.');
          pptx.setShapeImageCropShape(shape, changes.imageCropShape);
        }
        if (changes.aspectRatioLocked !== undefined) {
          if (typeof changes.aspectRatioLocked !== 'boolean')
            throw new Error('Invalid aspect ratio lock.');
          pptx.setShapeAspectRatioLocked(shape, changes.aspectRatioLocked);
        }
        if (changes.connections) setConnections(presentation, slide, shape, changes.connections);
        if (changes.bounds)
          pptx.setShapeBounds(
            shape,
            bounds(changes.bounds, pptx.getShapeKind(shape) === 'connector'),
          );
        if (changes.flip !== undefined) {
          if (
            typeof changes.flip.horizontal !== 'boolean' ||
            typeof changes.flip.vertical !== 'boolean'
          )
            throw new Error('Invalid flip flags.');
          pptx.setShapeFlip(shape, changes.flip);
        }
        if (changes.flipToggle !== undefined) {
          if (!['horizontal', 'vertical'].includes(changes.flipToggle))
            throw new Error('Invalid flip direction.');
          const flip = pptx.getShapeFlip(shape) ?? { horizontal: false, vertical: false };
          pptx.setShapeFlip(shape, { ...flip, [changes.flipToggle]: !flip[changes.flipToggle] });
        }
        if (changes.rotation !== undefined) {
          finite(changes.rotation, -36000, 36000);
          pptx.setShapeRotation(shape, changes.rotation);
        }
        if (changes.rotationDelta !== undefined) {
          const angle = pptx.getShapeRotation(shape) + changes.rotationDelta;
          pptx.setShapeRotation(shape, ((angle % 360) + 360) % 360);
        }
        if (changes.text !== undefined) {
          text(changes.text);
          pptx.setShapeText(shape, changes.text);
        }
        if (changes.textEdits !== undefined) {
          if (!Array.isArray(changes.textEdits) || changes.textEdits.length > 10000)
            throw new Error('Invalid text edits.');
          for (const edit of changes.textEdits) {
            text(edit.text);
            const format = edit.format ? editorFormat(edit.format) : null;
            pptx.replaceShapeTextRange(shape, edit.start, edit.end, edit.text);
            if (format && edit.text.length)
              pptx.setShapeTextRangeFormat(
                shape,
                edit.start,
                edit.start + edit.text.replace(/\r\n?/g, '\n').length,
                format,
              );
          }
        }
        if (changes.changeCase !== undefined) {
          for (const edit of textCaseEdits(
            pptx.getShapeText(shape),
            changes.changeCase,
            changes.range,
          ))
            pptx.replaceShapeTextRange(shape, edit.start, edit.end, edit.text);
        }
        if (changes.fill !== undefined) {
          if (changes.fill === 'none') pptx.setShapeNoFill(shape);
          else {
            const opacity = pptx.getShapeFillOpacity(shape);
            pptx.setShapeFill(shape, color(changes.fill));
            if (opacity !== null) pptx.setShapeFillOpacity(shape, opacity);
          }
        }
        if (changes.stroke !== undefined) {
          if (changes.stroke === 'none') pptx.setShapeNoStroke(shape);
          else {
            const stroke = changes.stroke;
            const opacity = pptx.getShapeStrokeOpacity(shape);
            if (stroke.widthEmu !== undefined) finite(stroke.widthEmu, 0, 20116800);
            pptx.setShapeStroke(shape, {
              ...(stroke.color !== undefined ? { color: color(stroke.color) } : {}),
              ...(stroke.widthEmu !== undefined ? { widthEmu: Math.round(stroke.widthEmu) } : {}),
            });
            if (stroke.color !== undefined && opacity !== null)
              pptx.setShapeStrokeOpacity(shape, opacity);
          }
        }
        if (changes.fillOpacity !== undefined) pptx.setShapeFillOpacity(shape, changes.fillOpacity);
        if (changes.strokeOpacity !== undefined)
          pptx.setShapeStrokeOpacity(shape, changes.strokeOpacity);
        if (changes.strokeCap !== undefined) pptx.setShapeStrokeCap(shape, changes.strokeCap);
        if (changes.strokeJoin !== undefined) pptx.setShapeStrokeJoin(shape, changes.strokeJoin);
        if (changes.strokeCompound !== undefined)
          pptx.setShapeStrokeCompound(shape, changes.strokeCompound);
        if (changes.strokeDash !== undefined) pptx.setShapeStrokeDash(shape, changes.strokeDash);
        for (const [end, options] of [
          ['head', changes.strokeHeadArrow],
          ['tail', changes.strokeTailArrow],
        ] as const) {
          if (options !== undefined)
            pptx.setShapeStrokeArrow(shape, end, {
              type: 'none',
              ...pptx.getShapeStrokeArrow(shape, end),
              ...options,
            });
        }
        if (changes.format) {
          const format = editorFormat(changes.format);
          if (changes.range)
            pptx.setShapeTextRangeFormat(shape, changes.range.start, changes.range.end, format);
          else pptx.setShapeTextFormat(shape, format);
        }
        const paragraphIndices: number[] = [];
        if (changes.range) {
          const { start, end } = changes.range;
          if (
            !Number.isInteger(start) ||
            !Number.isInteger(end) ||
            start < 0 ||
            end < start ||
            end > pptx.getShapeText(shape).length
          )
            throw new Error('Invalid text range.');
        }
        let paragraphStart = 0;
        const paragraphCount =
          changes.align !== undefined ||
          changes.bullets !== undefined ||
          changes.lineSpacing !== undefined ||
          changes.indent !== undefined ||
          changes.paragraph !== undefined
            ? pptx.getShapeParagraphCount(shape)
            : 0;
        for (let p = 0; p < paragraphCount; p++) {
          const length = pptx
            .getShapeParagraphElements(shape, p)
            .reduce((sum, element) => sum + (element.kind === 'br' ? 1 : element.text.length), 0);
          const range = changes.range;
          if (
            !range ||
            (range.start === range.end
              ? range.start >= paragraphStart && range.start <= paragraphStart + length
              : range.start <= paragraphStart + length && range.end > paragraphStart)
          )
            paragraphIndices.push(p);
          paragraphStart += length + 1;
        }
        if (changes.align !== undefined) {
          if (!['left', 'center', 'right', 'justify', 'distribute'].includes(changes.align))
            throw new Error('Invalid alignment.');
          for (const p of paragraphIndices) pptx.setParagraphAlignment(shape, p, changes.align);
        }
        if (changes.paragraph)
          for (const p of paragraphIndices) pptx.setParagraphSettings(shape, p, changes.paragraph);
        if (changes.bullets !== undefined) {
          if (!['bullet', 'number', 'none'].includes(changes.bullets))
            throw new Error('Invalid bullets.');
          for (const p of paragraphIndices) pptx.setParagraphBullet(shape, p, changes.bullets);
        }
        if (changes.margins !== undefined) {
          if (
            !changes.margins ||
            typeof changes.margins !== 'object' ||
            Array.isArray(changes.margins)
          )
            throw new Error('Invalid text margins.');
          for (const [key, value] of Object.entries(changes.margins)) {
            if (!['left', 'right', 'top', 'bottom'].includes(key))
              throw new Error('Invalid text margin.');
            finite(value, 0, 20116800);
          }
          pptx.setShapeTextMargins(shape, changes.margins);
        }
        if (changes.wrap !== undefined) {
          if (typeof changes.wrap !== 'boolean') throw new Error('Invalid text wrap.');
          pptx.setShapeTextWrap(shape, changes.wrap ? 'square' : 'none');
        }
        if (changes.autoFit !== undefined || changes.autoFitParams !== undefined)
          pptx.setShapeTextAutoFit(shape, changes.autoFit ?? 'normal', changes.autoFitParams);
        if (changes.anchorCenter !== undefined)
          pptx.setShapeTextAnchorCenter(shape, changes.anchorCenter);
        if (changes.anchor !== undefined) pptx.setShapeTextAnchor(shape, changes.anchor);
        if (changes.direction !== undefined) pptx.setShapeTextDirection(shape, changes.direction);
        if (changes.columns !== undefined) {
          if (!Number.isInteger(changes.columns)) throw new Error('Invalid columns.');
          finite(changes.columns, 1, 16);
          pptx.setShapeTextColumns(
            shape,
            changes.columns === 1 ? null : { count: changes.columns },
          );
        }
        if (changes.textColumns !== undefined) {
          const value = changes.textColumns;
          if (!value || typeof value !== 'object' || Array.isArray(value))
            throw new Error('Invalid columns.');
          const previous = pptx.getShapeTextColumns(shape);
          const count = value.count ?? previous?.count ?? 1;
          const gapEmu = value.gapEmu ?? previous?.gapEmu ?? 0;
          if (!Number.isInteger(count)) throw new Error('Invalid columns.');
          finite(count, 1, 16);
          finite(gapEmu, 0, 14630400);
          pptx.setShapeTextColumns(shape, { count, gapEmu: Math.round(gapEmu) });
        }
        if (changes.lineSpacing !== undefined || changes.indent !== undefined) {
          if (changes.lineSpacing !== undefined) finite(changes.lineSpacing, 0.1, 10);
          if (changes.indent !== undefined && ![-1, 1].includes(changes.indent))
            throw new Error('Invalid indentation.');
          for (const p of paragraphIndices) {
            if (changes.lineSpacing !== undefined)
              pptx.setParagraphLineSpacing(shape, p, { kind: 'pct', value: changes.lineSpacing });
            if (changes.indent !== undefined)
              pptx.setParagraphLevel(
                shape,
                p,
                Math.max(0, Math.min(8, pptx.getParagraphLevel(shape, p) + changes.indent)),
              );
          }
        }
      }
      for (const [id, previous] of tableBounds) {
        const table = target(id);
        const next = pptx.getShapeBoundsResolved(presentation, table);
        // Moving or rotating a table must not resize its authored rows.
        if (previous && next && (previous.w !== next.w || previous.h !== next.h))
          fitTableText(presentation, table, undefined, command.tableResizeAnchor);
      }
      for (const fit of command.textFits ?? []) {
        const shape = target(fit.id);
        const mode = pptx.getShapeBodyPrEffective(presentation, shape).autoFit;
        if (fit.bounds) {
          if (mode !== 'shape') throw new Error('Size measurement requires shape autofit.');
          pptx.setShapeBounds(shape, bounds(fit.bounds));
        }
        if (fit.autoFitParams) {
          if (mode !== 'normal') throw new Error('Font measurement requires normal autofit.');
          pptx.setShapeTextAutoFit(shape, 'normal', fit.autoFitParams);
        }
      }
      return;
    }
    default:
      throw new Error('Unsupported edit operation.');
  }
}

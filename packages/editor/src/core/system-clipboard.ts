// Slides and objects on the system clipboard.
//
// A copy writes one ClipboardItem. Other presentation apps cannot read a web
// page's objects natively, so they get renditions: styled HTML text and
// tables, an SVG drawing (where the browser writes SVG) and a PNG picture,
// plus plain text. Each app pastes the format it ranks highest. Editor
// instances also read the private payload .pptx (see clipboard-deck.ts), so
// slides and objects stay editable between tabs and browsers.

import {
  getParagraphPropertiesEffective,
  getShapeImageBytes,
  getShapeImageFormat,
  getShapeKind,
  getShapeMedia,
  getShapeBoundsResolved,
  getShapeParagraphElements,
  getShapeRunFormatEffective,
  getShapeText,
  getSlides,
  getSlideSize,
  getTableCellAnchor,
  getTableCellAppearanceEffective,
  getTableCellMargins,
  getTableCellParagraphs,
  getTableCellRunFormatEffective,
  getTableCells,
  getTableCellSpan,
  getTableColumnWidths,
  getTableRowHeights,
  isTableShape,
  toWritableTextFormat,
  type ParagraphProperties,
  type PresentationData,
  type SlideShapeData,
  type TableCellBorder,
  type TextFormat,
} from '@office-kit/pptx';
import {
  ARIAL,
  MONO,
  renderSlideToSvg,
  SANS,
  SERIF,
  TIMES,
  type TextMeasurer,
} from '@office-kit/pptx-preview';
import { shapeScope } from '../canvas/group-space.ts';
import {
  CLIPBOARD_MARKER_ATTRIBUTE,
  clipboardMarker,
  clipboardPlainText,
  clipboardShapeBounds,
  clipboardShapes,
  DECK_CLIPBOARD_TYPE,
  leafShapes,
  parseClipboardMarker,
  type ClipboardDeck,
  type ClipboardKind,
} from './clipboard-deck.ts';
import { textClipboardHtml } from './html-text-clipboard.ts';
import { copyTextRange } from './text-clipboard.ts';
import { t } from '../i18n/i18n.svelte.ts';

const EMU_PER_PX = 9525;
const DEFAULT_SLIDE = { width: 12192000, height: 6858000 };
/** The PNG is drawn at twice the slide's pixel size, so it stays sharp on high-density screens. */
const PNG_SCALE = 2;
/** Larger canvases fail to allocate in some browsers. */
const MAX_PNG_SIDE = 8192;
/** Room around copied objects for outlines and effects drawn outside their geometry. */
const OBJECT_IMAGE_MARGIN_PX = 4;
/** DrawingML's default cell margins: 0.1" left and right, 0.05" top and bottom. */
const DEFAULT_CELL_MARGIN = { left: 91440, right: 91440, top: 45720, bottom: 45720 };

/** Formats every browser with the async Clipboard API writes. */
const BASIC_TYPES = new Set(['text/plain', 'text/html', 'image/png']);
/** Pictures a paste accepts from other apps, in the order it prefers them. */
const IMAGE_TYPES = ['image/png', 'image/svg+xml', 'image/jpeg', 'image/gif', 'image/webp'];
const IMAGE_MIME: Readonly<Record<string, string>> = {
  png: 'image/png',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
  svg: 'image/svg+xml',
};

// The SVG text path names the metric-compatible substitute faces the preview
// measures with. Other apps rarely have them, so the clipboard drawing lists
// the faces they stand in for, then a generic family. The measurer resolves
// the same list, so line breaks match what the receiving app paints.
const FONT_FALLBACKS: Readonly<Record<string, readonly string[]>> = {
  [SANS]: [SANS, 'Calibri', 'Aptos', 'Helvetica Neue', 'Arial', 'sans-serif'],
  [SERIF]: [SERIF, 'Cambria', 'Georgia', 'serif'],
  [ARIAL]: [ARIAL, 'Arial', 'Helvetica', 'sans-serif'],
  [TIMES]: [TIMES, 'Times New Roman', 'Times', 'serif'],
  [MONO]: [MONO, 'Courier New', 'Courier', 'monospace'],
};
const GENERIC_FAMILY = /^(serif|sans-serif|monospace|cursive|fantasy|system-ui)$/;

function fontStack(family: string): string {
  return (FONT_FALLBACKS[family] ?? [family, 'sans-serif'])
    .map((name) => (GENERIC_FAMILY.test(name) ? name : `'${name.replace(/['\\]/g, '')}'`))
    .join(', ');
}

function clipboardTextMeasurer(): TextMeasurer | undefined {
  const context = document.createElement('canvas').getContext('2d');
  if (!context) return undefined;
  return (text, spec) => {
    context.font = `${spec.italic ? 'italic' : 'normal'} ${spec.bold ? 'bold' : 'normal'} ${spec.sizePx}px ${fontStack(spec.family)}`;
    context.fontKerning = spec.kerning === false ? 'none' : 'normal';
    context.letterSpacing = `${spec.letterSpacingPx}px`;
    const metrics = context.measureText(text);
    return {
      widthPx: metrics.width,
      ascentPx: metrics.fontBoundingBoxAscent,
      descentPx: metrics.fontBoundingBoxDescent,
      lineGapPx: 0,
      inkAscentPx: metrics.actualBoundingBoxAscent,
      inkDescentPx: metrics.actualBoundingBoxDescent,
    };
  };
}

const escapeAttribute = (value: string) => value.replace(/&/g, '&amp;').replace(/"/g, '&quot;');
const px = (value: number) => Number(value.toFixed(2));

export interface ClipboardDrawing {
  readonly svg: string;
  /** CSS pixels. */
  readonly width: number;
  readonly height: number;
}

/**
 * The SVG rendition, drawn with SVG text so it needs no HTML renderer: the
 * first copied slide, or the copied objects on a transparent surface cropped
 * to their bounds. Several slides draw the first one; their text is all in
 * the HTML.
 */
export function clipboardDrawing(deck: ClipboardDeck): ClipboardDrawing {
  const slide = getSlides(deck.pres)[0]!;
  const size = getSlideSize(deck.pres) ?? DEFAULT_SLIDE;
  const objects = deck.kind === 'shapes' ? clipboardShapeBounds(deck) : null;
  const box = objects
    ? {
        x: objects.x / EMU_PER_PX - OBJECT_IMAGE_MARGIN_PX,
        y: objects.y / EMU_PER_PX - OBJECT_IMAGE_MARGIN_PX,
        w: objects.w / EMU_PER_PX + 2 * OBJECT_IMAGE_MARGIN_PX,
        h: objects.h / EMU_PER_PX + 2 * OBJECT_IMAGE_MARGIN_PX,
      }
    : { x: 0, y: 0, w: size.width / EMU_PER_PX, h: size.height / EMU_PER_PX };
  const measureText = clipboardTextMeasurer();
  const rendered = renderSlideToSvg(deck.pres, slide, {
    textLayout: 'svg',
    background: deck.kind === 'slides',
    ...(measureText ? { measureText } : {}),
  });
  // Explicit width and height give receiving apps the drawing's real size.
  const svg = rendered
    .replace(
      /viewBox="[^"]*"/,
      `viewBox="${px(box.x)} ${px(box.y)} ${px(box.w)} ${px(box.h)}" width="${px(box.w)}" height="${px(box.h)}"`,
    )
    .replace(
      /font-family="([^"]*)"/g,
      (_, family: string) => `font-family="${escapeAttribute(fontStack(family))}"`,
    );
  return { svg, width: box.w, height: box.h };
}

/** Draws an SVG into a PNG. Pictures inside it are data: URLs, so the canvas is never tainted. */
export async function rasterizeSvg(
  svg: string,
  width: number,
  height: number,
  scale = PNG_SCALE,
): Promise<Blob> {
  const factor = Math.min(scale, MAX_PNG_SIDE / Math.max(width, height, 1));
  const url = URL.createObjectURL(new Blob([svg], { type: 'image/svg+xml' }));
  try {
    const image = new Image();
    image.src = url;
    await image.decode();
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.round(width * factor));
    canvas.height = Math.max(1, Math.round(height * factor));
    canvas.getContext('2d')!.drawImage(image, 0, 0, canvas.width, canvas.height);
    return await new Promise<Blob>((resolve, reject) =>
      canvas.toBlob(
        (blob) => (blob ? resolve(blob) : reject(new Error(t('The picture could not be drawn')))),
        'image/png',
      ),
    );
  } finally {
    URL.revokeObjectURL(url);
  }
}

function dataUrl(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error ?? new Error(t('The picture could not be read')));
    reader.readAsDataURL(blob);
  });
}

const CSS_ALIGN = {
  left: 'left',
  l: 'left',
  center: 'center',
  ctr: 'center',
  right: 'right',
  r: 'right',
  justify: 'justify',
  just: 'justify',
  justLow: 'justify',
  distribute: 'justify',
  dist: 'justify',
  thaiDist: 'justify',
} as const satisfies Record<NonNullable<ParagraphProperties['align']>, string>;

/**
 * A theme font slot the theme leaves empty stays a `+mn-ea`-style token; as
 * a CSS family it would only confuse the receiving app.
 */
function withoutThemeFontTokens(format: TextFormat): TextFormat {
  const out = { ...format };
  for (const key of ['font', 'fontEastAsian', 'fontComplexScript'] as const)
    if (out[key]?.startsWith('+')) delete out[key];
  return out;
}

type TextOwner = { shape: SlideShapeData; cell?: { row: number; col: number } };

/**
 * The text as HTML blocks, one per paragraph, styled as `textClipboardHtml`
 * styles copied text. Bulleted paragraphs become list items, which other
 * apps turn back into bullets.
 */
function paragraphBlocks(pres: PresentationData, { shape, cell }: TextOwner): HTMLElement[] {
  const cellData = cell ? getTableCells(shape)[cell.row]![cell.col]! : undefined;
  const paragraphs = cellData
    ? getTableCellParagraphs(cellData).map((paragraph) => paragraph.elements)
    : getShapeParagraphElements(shape);
  const blocks: HTMLElement[] = [];
  let list: HTMLElement | null = null;
  let offset = 0;
  paragraphs.forEach((elements, index) => {
    const length = elements.reduce(
      (sum, element) => sum + (element.kind === 'br' ? 1 : element.text.length),
      0,
    );
    const copied = copyTextRange(shape, offset, offset + length, cell, (paragraph, run) =>
      withoutThemeFontTokens(
        toWritableTextFormat(
          cellData
            ? getTableCellRunFormatEffective(pres, cellData, paragraph, run)
            : getShapeRunFormatEffective(pres, shape, paragraph, run),
        ),
      ),
    );
    // The next paragraph starts after the separating line feed.
    offset += length + 1;
    const template = document.createElement('template');
    template.innerHTML = textClipboardHtml(copied);
    const block = template.content.querySelector('div')!;
    block.style.margin = '0';
    if (!copied.text) block.append(document.createElement('br'));
    const properties = getParagraphPropertiesEffective(pres, cellData ?? shape, index);
    if (properties.align) block.style.textAlign = CSS_ALIGN[properties.align];
    const bullet = properties.bullet;
    const tag =
      bullet === null || bullet === 'none'
        ? null
        : bullet === 'number' || (typeof bullet === 'object' && 'autoNum' in bullet)
          ? 'ol'
          : 'ul';
    if (!tag) {
      list = null;
      blocks.push(block);
      return;
    }
    if (list?.localName !== tag) {
      list = document.createElement(tag);
      list.style.margin = '0';
      blocks.push(list);
    }
    const item = document.createElement('li');
    item.append(block);
    list.append(item);
  });
  return blocks;
}

const cssColor = (color: string) => (color.startsWith('#') ? color : `#${color}`);

function cssBorder(border: TableCellBorder | null): string {
  if (!border?.color) return 'none';
  const width = Math.max(1, (border.widthEmu ?? EMU_PER_PX) / EMU_PER_PX);
  return `${px(width)}px ${border.dash && border.dash !== 'solid' ? 'dashed' : 'solid'} ${cssColor(border.color)}`;
}

function tableElement(pres: PresentationData, shape: SlideShapeData): HTMLTableElement {
  const table = document.createElement('table');
  table.style.borderCollapse = 'collapse';
  const columns = document.createElement('colgroup');
  for (const width of getTableColumnWidths(shape)) {
    const column = document.createElement('col');
    column.style.width = `${px(width / EMU_PER_PX)}px`;
    columns.append(column);
  }
  table.append(columns);
  const heights = getTableRowHeights(shape);
  getTableCells(shape).forEach((cells, row) => {
    const tr = table.insertRow();
    if (heights[row]) tr.style.height = `${px(heights[row] / EMU_PER_PX)}px`;
    cells.forEach((cell, col) => {
      const span = getTableCellSpan(cell);
      if (span.hMerge || span.vMerge) return;
      const td = tr.insertCell();
      if (span.gridSpan > 1) td.colSpan = span.gridSpan;
      if (span.rowSpan > 1) td.rowSpan = span.rowSpan;
      const { fill, borders } = getTableCellAppearanceEffective(pres, cell);
      if (fill.kind === 'solid') td.style.backgroundColor = cssColor(fill.color);
      td.style.borderLeft = cssBorder(borders.left);
      td.style.borderRight = cssBorder(borders.right);
      td.style.borderTop = cssBorder(borders.top);
      td.style.borderBottom = cssBorder(borders.bottom);
      const margins = getTableCellMargins(cell);
      td.style.padding = (['top', 'right', 'bottom', 'left'] as const)
        .map((side) => `${px((margins[side] ?? DEFAULT_CELL_MARGIN[side]) / EMU_PER_PX)}px`)
        .join(' ');
      const anchor = getTableCellAnchor(cell);
      td.style.verticalAlign = anchor === 'center' ? 'middle' : (anchor ?? 'top');
      td.append(...paragraphBlocks(pres, { shape, cell: { row, col } }));
    });
  });
  return table;
}

async function pictureElement(
  pres: PresentationData,
  shape: SlideShapeData,
): Promise<HTMLImageElement | null> {
  const format = getShapeImageFormat(shape);
  const bytes = getShapeImageBytes(shape);
  const mime = format ? IMAGE_MIME[format] : undefined;
  if (!bytes || !mime || getShapeMedia(shape)) return null;
  const image = document.createElement('img');
  image.src = await dataUrl(new Blob([bytes as BlobPart], { type: mime }));
  const bounds = getShapeBoundsResolved(pres, shape);
  if (bounds) {
    image.width = Math.round(bounds.w / EMU_PER_PX);
    image.height = Math.round(bounds.h / EMU_PER_PX);
  }
  return image;
}

/**
 * The HTML rendition: the copied text as styled paragraphs and lists, tables
 * as tables, and — for copied objects — pictures as data: images. Drawings
 * (shapes, charts) are in the image formats; when nothing else is left the
 * HTML carries `picture` so it still pastes as something.
 */
export async function clipboardHtml(
  deck: ClipboardDeck,
  id: string,
  picture: () => Promise<{ blob: Blob; width: number; height: number }>,
): Promise<string> {
  const root = document.createElement('div');
  root.setAttribute(CLIPBOARD_MARKER_ATTRIBUTE, clipboardMarker(deck.kind, id));
  const shapes =
    deck.kind === 'shapes'
      ? clipboardShapes(deck.pres)
      : getSlides(deck.pres).flatMap((slide) => shapeScope(slide, null).shapes);
  for (const shape of leafShapes(shapes)) {
    if (isTableShape(shape)) root.append(tableElement(deck.pres, shape));
    else if (getShapeKind(shape) === 'picture') {
      const image = deck.kind === 'shapes' ? await pictureElement(deck.pres, shape) : null;
      if (image) root.append(image);
    } else if (getShapeKind(shape) === 'shape' && getShapeText(shape).trim()) {
      root.append(...paragraphBlocks(deck.pres, { shape }));
    }
  }
  if (!root.childElementCount) {
    const { blob, width, height } = await picture();
    const image = document.createElement('img');
    image.src = await dataUrl(blob);
    image.width = Math.round(width);
    image.height = Math.round(height);
    root.append(image);
  }
  return root.outerHTML;
}

function supportsType(type: string): boolean {
  return typeof ClipboardItem.supports === 'function'
    ? ClipboardItem.supports(type)
    : BASIC_TYPES.has(type);
}

export function canWriteSystemClipboard(): boolean {
  return typeof ClipboardItem !== 'undefined' && typeof navigator.clipboard?.write === 'function';
}

export function canReadSystemClipboard(): boolean {
  return typeof navigator !== 'undefined' && typeof navigator.clipboard?.read === 'function';
}

/**
 * Writes the copy to the system clipboard. Call it inside the user gesture:
 * the ClipboardItem is created at once with pending values, which the
 * browser waits for (Safari requires this).
 */
export function writeSystemClipboard(id: string, deck: Promise<ClipboardDeck>): Promise<void> {
  if (!canWriteSystemClipboard())
    return Promise.reject(new Error('The system clipboard is unavailable'));
  const blob = (value: string | Uint8Array, type: string) =>
    new Blob([value as BlobPart], { type });
  const drawing = deck.then(clipboardDrawing);
  let png: Promise<{ blob: Blob; width: number; height: number }> | undefined;
  const picture = () =>
    (png ??= drawing.then(async ({ svg, width, height }) => ({
      blob: await rasterizeSvg(svg, width, height),
      width,
      height,
    })));
  const entries: Record<string, Promise<Blob>> = {
    'text/plain': deck.then((value) => blob(clipboardPlainText(value), 'text/plain')),
    'text/html': deck
      .then((value) => clipboardHtml(value, id, picture))
      .then((html) => blob(html, 'text/html')),
    'image/png': picture().then((value) => value.blob),
  };
  if (supportsType('image/svg+xml'))
    entries['image/svg+xml'] = drawing.then(({ svg }) => blob(svg, 'image/svg+xml'));
  if (supportsType(DECK_CLIPBOARD_TYPE))
    entries[DECK_CLIPBOARD_TYPE] = deck.then((value) => blob(value.bytes, DECK_CLIPBOARD_TYPE));
  return navigator.clipboard.write([new ClipboardItem(entries)]);
}

/**
 * A picture pasted from another app, with its size in CSS pixels. An SVG is
 * drawn into a PNG first: a slide picture needs a bitmap other apps can show.
 */
export async function pastedPicture(
  blob: Blob,
): Promise<{ bytes: Uint8Array; width: number; height: number }> {
  const url = URL.createObjectURL(blob);
  let width: number;
  let height: number;
  try {
    const image = new Image();
    image.src = url;
    await image.decode();
    width = image.naturalWidth;
    height = image.naturalHeight;
  } finally {
    URL.revokeObjectURL(url);
  }
  if (!width || !height) throw new Error(t('The picture has no size'));
  const bitmap =
    blob.type === 'image/svg+xml' ? await rasterizeSvg(await blob.text(), width, height) : blob;
  return { bytes: new Uint8Array(await bitmap.arrayBuffer()), width, height };
}

/** What the system clipboard holds, read lazily: pictures and payloads load only when pasted. */
export interface SystemClipboard {
  /** Set when an editor wrote it. */
  readonly marker: { readonly kind: ClipboardKind; readonly id: string } | null;
  /** The payload .pptx; null when this read cannot see custom formats. */
  readonly deck: (() => Promise<Uint8Array>) | null;
  readonly image: (() => Promise<Blob>) | null;
  readonly html: string;
  readonly text: string;
  readonly empty: boolean;
}

function systemClipboard(parts: Omit<SystemClipboard, 'marker' | 'empty'>): SystemClipboard {
  return {
    ...parts,
    marker: parseClipboardMarker(parts.html),
    empty: !parts.deck && !parts.image && !parts.html && !parts.text,
  };
}

/** Reads through the async Clipboard API; rejects when access is denied. */
export async function readSystemClipboard(): Promise<SystemClipboard> {
  const items = await navigator.clipboard.read();
  const find = (type: string) => items.find((item) => item.types.includes(type));
  const text = async (type: string) => (await find(type)?.getType(type))?.text() ?? '';
  const deckItem = find(DECK_CLIPBOARD_TYPE);
  const imageType = IMAGE_TYPES.find((type) => find(type));
  return systemClipboard({
    deck: deckItem
      ? async () =>
          new Uint8Array(await (await deckItem.getType(DECK_CLIPBOARD_TYPE)).arrayBuffer())
      : null,
    image: imageType ? () => find(imageType)!.getType(imageType) : null,
    html: await text('text/html'),
    text: await text('text/plain'),
  });
}

/** Reads a paste event's data, which never includes custom formats. */
export function systemClipboardFromEvent(data: DataTransfer): SystemClipboard {
  const files = [...data.files];
  const image = IMAGE_TYPES.map((type) => files.find((file) => file.type === type)).find(
    (file) => file !== undefined,
  );
  return systemClipboard({
    deck: null,
    image: image ? () => Promise.resolve(image) : null,
    html: data.getData('text/html'),
    text: data.getData('text/plain'),
  });
}

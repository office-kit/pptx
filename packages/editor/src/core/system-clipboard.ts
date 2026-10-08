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
  getShapeFillEffective,
  getShapeKind,
  getShapeParagraphElements,
  getShapePlaceholderType,
  getShapeRunFormatEffective,
  getShapeStrokeEffective,
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
  isShapeTextBox,
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
/** CSS pixels per inch. */
const CSS_DPI = 96;
const METRES_PER_INCH = 0.0254;
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
 * A drawing with SVG text, so it needs no HTML renderer: one copied slide, or
 * the copied objects on a transparent surface cropped to their bounds. The
 * SVG and PNG entries hold one picture, so a copy of several slides puts the
 * first one there; the HTML has every slide.
 */
export function clipboardDrawing(deck: ClipboardDeck, slideIndex = 0): ClipboardDrawing {
  const slide = getSlides(deck.pres)[slideIndex]!;
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
    const png = await new Promise<Blob>((resolve, reject) =>
      canvas.toBlob(
        (blob) => (blob ? resolve(blob) : reject(new Error(t('The picture could not be drawn')))),
        'image/png',
      ),
    );
    return withPixelDensity(png, (canvas.width / Math.max(width, 1)) * CSS_DPI);
  } finally {
    URL.revokeObjectURL(url);
  }
}

/** Signature plus the IHDR chunk, which every PNG starts with. */
const PNG_HEADER_BYTES = 33;
const PNG_UNIT_METRE = 1;

let crcTable: Uint32Array | undefined;
function crc32(bytes: Uint8Array): number {
  crcTable ??= Uint32Array.from({ length: 256 }, (_, n) => {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c;
  });
  let crc = 0xffffffff;
  for (const byte of bytes) crc = crcTable[(crc ^ byte) & 0xff]! ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

/**
 * Adds a pHYs chunk so other apps size the picture as it appears on the slide.
 * Canvas PNGs carry none, and the reference desktop app on Mac then reads
 * them at 72 DPI: a 2× drawing pastes at nearly three times its size. It
 * ignores width and height on an HTML image, so the density is the only size
 * it takes. (Chromium re-encodes the image/png entry and drops the chunk; the
 * data: image inside the HTML keeps it.)
 */
async function withPixelDensity(png: Blob, dpi: number): Promise<Blob> {
  const bytes = new Uint8Array(await png.arrayBuffer());
  const perMetre = Math.round(dpi / METRES_PER_INCH);
  const chunk = new Uint8Array(21);
  const view = new DataView(chunk.buffer);
  view.setUint32(0, 9);
  chunk.set([0x70, 0x48, 0x59, 0x73], 4); // "pHYs"
  view.setUint32(8, perMetre);
  view.setUint32(12, perMetre);
  chunk[16] = PNG_UNIT_METRE;
  view.setUint32(17, crc32(chunk.subarray(4, 17)));
  return new Blob([bytes.subarray(0, PNG_HEADER_BYTES), chunk, bytes.subarray(PNG_HEADER_BYTES)], {
    type: 'image/png',
  });
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

type Picture = { readonly blob: Blob; readonly width: number; readonly height: number };

/** Text another app can rebuild from HTML: a table, or text with no paint of its own. */
function isPlainText(pres: PresentationData, shape: SlideShapeData): boolean {
  if (isTableShape(shape)) return true;
  if (getShapeKind(shape) !== 'shape' || !getShapeText(shape).trim()) return false;
  // A non-placeholder autoshape that inherits its paint takes it from its theme style.
  if (!isShapeTextBox(shape) && !getShapePlaceholderType(shape)) return false;
  const unpainted = (kind: string) => kind === 'none' || kind === 'inherit';
  return (
    unpainted(getShapeFillEffective(pres, shape).kind) &&
    unpainted(getShapeStrokeEffective(pres, shape).kind)
  );
}

async function imageElement({ blob, width, height }: Picture): Promise<HTMLImageElement> {
  const image = document.createElement('img');
  image.src = await dataUrl(blob);
  image.width = Math.round(width);
  image.height = Math.round(height);
  return image;
}

/**
 * The HTML rendition. Desktop presentation apps rank HTML above pictures and
 * paste it as text boxes, tables and images, so a slide goes in as a picture
 * of the slide, and objects go in as styled text and tables only when they
 * are nothing else; any other object makes the whole copy one picture,
 * rather than pasting the text and dropping the drawing.
 */
export async function clipboardHtml(
  deck: ClipboardDeck,
  id: string,
  picture: (slideIndex: number) => Promise<Picture>,
): Promise<string> {
  const root = document.createElement('div');
  root.setAttribute(CLIPBOARD_MARKER_ATTRIBUTE, clipboardMarker(deck.kind, id));
  const shapes = deck.kind === 'shapes' ? leafShapes(clipboardShapes(deck.pres)) : [];
  if (shapes.length && shapes.every((shape) => isPlainText(deck.pres, shape))) {
    for (const shape of shapes)
      if (isTableShape(shape)) root.append(tableElement(deck.pres, shape));
      else root.append(...paragraphBlocks(deck.pres, { shape }));
  } else {
    const count = deck.kind === 'slides' ? getSlides(deck.pres).length : 1;
    const pictures = await Promise.all(Array.from({ length: count }, (_, index) => picture(index)));
    root.append(...(await Promise.all(pictures.map(imageElement))));
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
  const drawing = deck.then((value) => clipboardDrawing(value));
  const pictures = new Map<number, Promise<Picture>>();
  const picture = (slideIndex: number) => {
    let png = pictures.get(slideIndex);
    if (!png) {
      png = (
        slideIndex === 0 ? drawing : deck.then((value) => clipboardDrawing(value, slideIndex))
      ).then(async ({ svg, width, height }) => ({
        blob: await rasterizeSvg(svg, width, height),
        width,
        height,
      }));
      pictures.set(slideIndex, png);
    }
    return png;
  };
  const entries: Record<string, Promise<Blob>> = {
    'text/plain': deck.then((value) => blob(clipboardPlainText(value), 'text/plain')),
    'text/html': deck
      .then((value) => clipboardHtml(value, id, picture))
      .then((html) => blob(html, 'text/html')),
    'image/png': picture(0).then((value) => value.blob),
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

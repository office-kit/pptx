// Slide-level shape authoring.

import type { Emu } from '../units.ts';
import {
  contentTypeForFormat,
  detectImageFormat,
  emptyRels,
  extensionForFormat,
  type ImageFormat,
  nextRelId,
  partName,
} from '../../internal/opc/index.ts';
import {
  REL_TYPES,
  type PresetShape,
  buildConnector,
  buildPicture,
  buildShape,
  buildTable,
  buildTextBox,
} from '../../internal/presentationml/index.ts';
import {
  INTERNAL_PACKAGE,
  SLIDE_PART_NAME,
  type SlideData,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import { appendAndReturnNewShape, nextShapeId } from './_helpers.ts';
// ---------------------------------------------------------------------------
// Slide-level shape authoring.
//
// Each `addXxx` builds an XML element via an internal builder, appends
// it to the slide's `<p:spTree>`, commits, rebuilds the typed view, and
// returns the new SlideShapeData.

/**
 * Adds a free-form text box to the slide. Returns the new shape.
 *
 * The box is a plain rectangle with no fill or outline carrying one
 * paragraph with one run. The shape id is allocated as one more than
 * the current max id.
 */
export const addSlideTextBox = (
  slide: SlideData,
  opts: { x: Emu; y: Emu; w: Emu; h: Emu; text: string; name?: string },
): SlideShapeData => {
  const sp = buildTextBox({
    id: nextShapeId(slide),
    ...(opts.name !== undefined ? { name: opts.name } : {}),
    x: opts.x,
    y: opts.y,
    w: opts.w,
    h: opts.h,
    text: opts.text,
  });
  return appendAndReturnNewShape(slide, sp);
};

/**
 * Adds a preset shape (rectangle, ellipse, arrow, ...) to the slide.
 * Optional `text` seeds a single run.
 */
export const addSlideShape = (
  slide: SlideData,
  opts: {
    preset: PresetShape | string;
    x: Emu;
    y: Emu;
    w: Emu;
    h: Emu;
    text?: string;
    textAnchor?: 'l' | 'ctr' | 'r' | 't' | 'b';
    name?: string;
  },
): SlideShapeData => {
  const sp = buildShape({
    id: nextShapeId(slide),
    ...(opts.name !== undefined ? { name: opts.name } : {}),
    preset: opts.preset,
    x: opts.x,
    y: opts.y,
    w: opts.w,
    h: opts.h,
    ...(opts.text !== undefined ? { text: opts.text } : {}),
    ...(opts.textAnchor !== undefined ? { textAnchor: opts.textAnchor } : {}),
  });
  return appendAndReturnNewShape(slide, sp);
};

/** Adds a straight-line connector between two points. */
export const addSlideLine = (
  slide: SlideData,
  opts: {
    from: { x: Emu; y: Emu };
    to: { x: Emu; y: Emu };
    color?: string;
    widthEmu?: number;
    name?: string;
  },
): SlideShapeData => {
  const cxn = buildConnector({
    id: nextShapeId(slide),
    ...(opts.name !== undefined ? { name: opts.name } : {}),
    from: opts.from,
    to: opts.to,
    ...(opts.color !== undefined ? { color: opts.color } : {}),
    ...(opts.widthEmu !== undefined ? { widthEmu: opts.widthEmu } : {}),
  });
  return appendAndReturnNewShape(slide, cxn);
};

/**
 * Adds a table to the slide. Cells render as plain text with default
 * theme-aware styling; `firstRow` / `bandRow` flags drive PowerPoint's
 * banded-header look unless options say otherwise.
 */
export const addSlideTable = (
  slide: SlideData,
  opts: {
    x: Emu;
    y: Emu;
    w: Emu;
    h: Emu;
    rows: ReadonlyArray<ReadonlyArray<string>>;
    colWidths?: ReadonlyArray<Emu>;
    rowHeights?: ReadonlyArray<Emu>;
    firstRow?: boolean;
    bandRow?: boolean;
    name?: string;
  },
): SlideShapeData => {
  const frame = buildTable({
    id: nextShapeId(slide),
    ...(opts.name !== undefined ? { name: opts.name } : {}),
    x: opts.x,
    y: opts.y,
    w: opts.w,
    h: opts.h,
    rows: opts.rows,
    ...(opts.colWidths !== undefined ? { colWidths: opts.colWidths } : {}),
    ...(opts.rowHeights !== undefined ? { rowHeights: opts.rowHeights } : {}),
    ...(opts.firstRow !== undefined ? { firstRow: opts.firstRow } : {}),
    ...(opts.bandRow !== undefined ? { bandRow: opts.bandRow } : {}),
  });
  return appendAndReturnNewShape(slide, frame);
};

/**
 * Adds a picture to the slide from raw bytes. Returns the new shape.
 *
 * Allocates a `/ppt/media/imageN.<ext>` part, registers a Content_Types
 * Default if the extension isn't yet covered, allocates a slide→image
 * rel, and appends a `<p:pic>` element to the slide's `<p:spTree>`.
 *
 * Format is detected from magic bytes; pass `opts.format` to override.
 */
export const addSlideImage = (
  slide: SlideData,
  bytes: Uint8Array,
  opts: { x: Emu; y: Emu; w: Emu; h: Emu; format?: ImageFormat; name?: string },
): SlideShapeData => {
  const pkg = slide[INTERNAL_PACKAGE];
  const format = opts.format ?? detectImageFormat(bytes);
  if (format === null) {
    throw new Error(
      'addSlideImage: could not detect image format. Pass options.format explicitly.',
    );
  }
  const contentType = contentTypeForFormat(format);
  const extension = extensionForFormat(format);

  let nextN = 1;
  const mediaPattern = /^\/ppt\/media\/image(\d+)\./;
  for (const p of pkg.parts) {
    const m = p.name.match(mediaPattern);
    if (m?.[1] !== undefined) {
      const n = Number.parseInt(m[1], 10);
      if (Number.isFinite(n) && n >= nextN) nextN = n + 1;
    }
  }
  const newMediaName = partName(`/ppt/media/image${nextN}.${extension}`);

  const hasDefault = pkg.contentTypes.defaults.some((d) => d.extension.toLowerCase() === extension);
  if (!hasDefault) {
    pkg.contentTypes.defaults.push({ extension, contentType });
  }
  pkg.addPart(newMediaName, contentType, bytes);

  const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
  const newRId = nextRelId(rels.items.map((r) => r.id));
  rels.items.push({
    id: newRId,
    type: REL_TYPES.image,
    target: `../media/image${nextN}.${extension}`,
    targetMode: 'Internal',
  });
  pkg.setRels(slide[SLIDE_PART_NAME], rels);

  const pic = buildPicture({
    id: nextShapeId(slide),
    ...(opts.name !== undefined ? { name: opts.name } : {}),
    rEmbed: newRId,
    x: opts.x,
    y: opts.y,
    w: opts.w,
    h: opts.h,
  });
  return appendAndReturnNewShape(slide, pic);
};

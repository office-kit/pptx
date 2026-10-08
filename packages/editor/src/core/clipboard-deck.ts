// The editor's clipboard payload: a small .pptx holding just what was copied.
//
// Copying slides keeps those slides; copying objects keeps their slide with
// only those objects on it, at their positions. Layouts, masters, themes and
// media come along, so the payload pastes back into any editor instance — in
// this tab, another tab or another browser — and opens in other presentation
// apps as an ordinary file. Nothing here touches the DOM.

import {
  compactPackage,
  copyShape,
  findShapeById,
  getGroupChildren,
  getShapeBoundsResolved,
  getShapeId,
  getShapeKind,
  getShapeRotation,
  getShapeText,
  getSlides,
  getSlideShapes,
  getTableCells,
  getTableCellText,
  isTableShape,
  loadPresentation,
  removeShape,
  removeSlide,
  savePresentation,
  type PresentationData,
  type SlideData,
  type SlideShapeData,
} from '@office-kit/pptx';
import { compose, groupMatrix, IDENTITY, shapeScope, type Matrix } from '../canvas/group-space.ts';
import { projectedBounds } from '../canvas/transformed-snapping.ts';
import type { Rect } from '../canvas/snapping.ts';
import { serializeTableClipboard } from './table-clipboard.ts';

/**
 * The private clipboard format carrying the payload .pptx. Browsers that
 * support custom formats (the `web ` prefix) keep it next to the standard
 * text, HTML and image formats other apps read.
 */
export const DECK_CLIPBOARD_TYPE = 'web application/x-pptx-editor-clipboard+zip';

/**
 * Attribute on the clipboard HTML's root element: `<kind>:<id>`. It tells a
 * paste which copy the clipboard holds, including when only HTML is readable
 * (the paste event exposes no custom formats).
 */
export const CLIPBOARD_MARKER_ATTRIBUTE = 'data-pptx-editor-clipboard';

export type ClipboardKind = 'slides' | 'shapes';

export type ClipboardContent =
  | { readonly kind: 'slides'; readonly slideIndices: readonly number[] }
  | { readonly kind: 'shapes'; readonly slideIndex: number; readonly shapeIds: readonly number[] };

export interface ClipboardDeck {
  readonly kind: ClipboardKind;
  /** The payload, read-only from here on: pasting copies out of it. */
  readonly pres: PresentationData;
  readonly bytes: Uint8Array;
}

/**
 * Builds the payload from a saved copy of the deck. `source` is a byte
 * snapshot, so the payload is independent of later edits (and of a cut).
 */
export async function buildClipboardDeck(
  source: Uint8Array,
  content: ClipboardContent,
): Promise<ClipboardDeck> {
  const pres = await loadPresentation(source);
  const slides = getSlides(pres);
  const keep = new Set(content.kind === 'slides' ? content.slideIndices : [content.slideIndex]);
  if (content.kind === 'shapes') isolateShapes(slides[content.slideIndex]!, content.shapeIds);
  for (const [index, slide] of slides.entries()) if (!keep.has(index)) removeSlide(pres, slide);
  // Removing slides leaves their pictures behind; only the copied ones stay.
  compactPackage(pres);
  return { kind: content.kind, pres, bytes: await savePresentation(pres) };
}

/**
 * Leaves exactly the given objects on the slide, at the top level and in
 * stacking order. An object inside a group keeps the group transforms that
 * place it, as `copyShape` writes them.
 */
function isolateShapes(slide: SlideData, shapeIds: readonly number[]): void {
  const original = shapeScope(slide, null).shapes;
  const stacking = new Map(getSlideShapes(slide).map((shape, index) => [getShapeId(shape), index]));
  const selected = shapeIds
    .map((id) => findShapeById(slide, id))
    .filter((shape) => shape !== null)
    .sort((a, b) => stacking.get(getShapeId(a))! - stacking.get(getShapeId(b))!);
  for (const shape of selected) copyShape(slide, shape, { preserveGroupTransform: true });
  for (const shape of original) removeShape(shape);
}

/** The objects a payload pastes: every slide's, or the one slide's top-level objects. */
export function clipboardShapes(pres: PresentationData): SlideShapeData[] {
  const slide = getSlides(pres)[0];
  return slide ? [...shapeScope(slide, null).shapes] : [];
}

/** Objects in reading order, with groups opened up. */
export function leafShapes(shapes: readonly SlideShapeData[]): SlideShapeData[] {
  return shapes.flatMap((shape) =>
    getShapeKind(shape) === 'group' ? leafShapes(getGroupChildren(shape)) : [shape],
  );
}

/** Table cell text, one row per line and tab-separated, as spreadsheets read it. */
function tableText(table: SlideShapeData): string {
  return serializeTableClipboard(getTableCells(table).map((row) => row.map(getTableCellText)));
}

function shapesText(shapes: readonly SlideShapeData[]): string[] {
  return leafShapes(shapes).flatMap((shape) => {
    const text = isTableShape(shape) ? tableText(shape) : getShapeText(shape);
    return text.trim() ? [text] : [];
  });
}

/** The `text/plain` rendition: the copied text in reading order. */
export function clipboardPlainText(deck: ClipboardDeck): string {
  if (deck.kind === 'shapes') return shapesText(clipboardShapes(deck.pres)).join('\n');
  return getSlides(deck.pres)
    .map((slide) => shapesText(shapeScope(slide, null).shapes).join('\n'))
    .filter((text) => text)
    .join('\n\n');
}

/** The slide-space box around the payload's objects, as drawn (rotation and groups included). */
export function clipboardShapeBounds(deck: ClipboardDeck): Rect | null {
  const slide = getSlides(deck.pres)[0];
  if (!slide) return null;
  const rects: Rect[] = [];
  const collect = (shapes: readonly SlideShapeData[], matrix: Matrix): void => {
    for (const shape of shapes) {
      if (getShapeKind(shape) === 'group') {
        collect(getGroupChildren(shape), compose(matrix, groupMatrix(shape)));
        continue;
      }
      const bounds = getShapeBoundsResolved(deck.pres, shape);
      if (bounds)
        rects.push(projectedBounds({ ...bounds, rotation: getShapeRotation(shape) }, matrix));
    }
  };
  collect(shapeScope(slide, null).shapes, IDENTITY);
  if (!rects.length) return null;
  const left = Math.min(...rects.map((rect) => rect.x));
  const top = Math.min(...rects.map((rect) => rect.y));
  const right = Math.max(...rects.map((rect) => rect.x + rect.w));
  const bottom = Math.max(...rects.map((rect) => rect.y + rect.h));
  return { x: left, y: top, w: right - left, h: bottom - top };
}

export function clipboardMarker(kind: ClipboardKind, id: string): string {
  return `${kind}:${id}`;
}

/** Reads the marker from clipboard HTML without parsing the markup. */
export function parseClipboardMarker(html: string): { kind: ClipboardKind; id: string } | null {
  const match = new RegExp(`${CLIPBOARD_MARKER_ATTRIBUTE}="(slides|shapes):([\\w-]{1,64})"`).exec(
    html,
  );
  if (!match) return null;
  return { kind: match[1] === 'slides' ? 'slides' : 'shapes', id: match[2]! };
}

/**
 * Where a paste comes from. This editor's own copy wins while its system
 * clipboard write is still running or could not be made — the system
 * clipboard then holds something older. Otherwise whatever the system
 * clipboard holds is newer, unless it is that same copy.
 */
export function pasteSource(
  own: { readonly id: string; readonly written: 'pending' | 'done' | 'failed' } | null,
  system: { readonly id: string | null; readonly empty: boolean } | null,
): 'own' | 'system' | 'none' {
  if (own && (own.written !== 'done' || !system || system.empty || system.id === own.id))
    return 'own';
  return system && !system.empty ? 'system' : 'none';
}

/** `deck-slide-3.pptx`, `deck-slides-2-4.pptx`: names the slides a download holds. */
export function selectedSlidesFileName(fileName: string, slideIndices: readonly number[]): string {
  const stem = fileName.replace(/\.pptx$/i, '') || 'Presentation';
  const numbers = slideIndices.map((index) => index + 1).join('-');
  return `${stem}-${slideIndices.length === 1 ? 'slide' : 'slides'}-${numbers}.pptx`;
}

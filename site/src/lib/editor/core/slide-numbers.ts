// Deck-wide slide numbers — the way Google Slides presents them: one switch
// that puts a live number on every slide, not a field the author inserts slide
// by slide and keeps in sync.
//
// The number itself is `setShapeTextField(shape, 'slidenum')`; everything here
// decides *which* shape carries it. A template reserves a `sldNum` placeholder
// for exactly this, so the switch fills that slot — position, font and colour
// then come from the layout and master, and the number looks like part of the
// design. Only a deck whose layout has no such slot gets a plain corner box.

import {
  addSlidePlaceholder,
  addSlideTextBox,
  emu,
  getShapeKind,
  getShapeParagraphCount,
  getShapeParagraphElements,
  getSlideShapes,
  getSlideSize,
  inches,
  removeShape,
  setParagraphAlignment,
  setShapeTextField,
  setShapeTextFormat,
  type PresentationData,
  type ShapeBounds,
  type SlideData,
  type SlideShapeData,
} from '@office-kit/pptx';

/** Type size of a generated slide number, matching PowerPoint's own default. */
const SLIDE_NUMBER_PT = 12;

/**
 * The shape that holds nothing but a slide-number field — what this module
 * created, or what a template authored the same way. A box that merely mentions
 * the number among other text is left alone: turning the switch off must not
 * delete an author's content.
 */
export const slideNumberShape = (slide: SlideData): SlideShapeData | null => {
  for (const shape of getSlideShapes(slide)) {
    // Only a `p:sp` can hold a text body of its own; asking a picture, a table
    // or a group for its paragraphs throws rather than reporting none.
    if (getShapeKind(shape) !== 'shape' || getShapeParagraphCount(shape) !== 1) continue;
    const elements = getShapeParagraphElements(shape, 0);
    if (elements.length === 1 && elements[0]!.kind === 'fld' && elements[0]!.type === 'slidenum')
      return shape;
  }
  return null;
};

/**
 * Puts a live slide number on the slide, and leaves it alone when it already
 * has one. Returns the shape carrying the number, or null when the layout
 * reserves no slot and the slide size — which the fallback box is placed
 * against — is unavailable.
 */
export const showSlideNumber = (
  pres: PresentationData,
  slide: SlideData,
): SlideShapeData | null => {
  const existing = slideNumberShape(slide);
  if (existing) return existing;

  // Restores the slot if the author deleted it, and finds it if they did not.
  const placeholder = addSlidePlaceholder(slide, 'sldNum');
  if (placeholder) {
    setShapeTextField(placeholder, 'slidenum');
    return placeholder;
  }

  const bounds = cornerBounds(pres);
  if (!bounds) return null;
  const box = addSlideTextBox(slide, { ...bounds, text: '' });
  // `setShapeTextFormat` writes into runs, never into a field, and
  // `setShapeTextField` carries the run it replaces onto the field — so the
  // size has to be set while the box still holds a run.
  setShapeTextFormat(box, { size: SLIDE_NUMBER_PT });
  setShapeTextField(box, 'slidenum');
  // The field replaces every paragraph, and the old `<a:pPr>` goes with them.
  setParagraphAlignment(box, 0, 'r');
  return box;
};

/** Removes the slide's number, if the slide has one this module recognises. */
export const hideSlideNumber = (slide: SlideData): boolean => {
  const shape = slideNumberShape(slide);
  if (!shape) return false;
  // Removing the `sldNum` placeholder rather than emptying it is what makes the
  // switch symmetrical: turning it back on restores the slot from the layout.
  removeShape(shape);
  return true;
};

/** Bottom-right of the slide, for a deck whose layout reserves no slot. */
const cornerBounds = (pres: PresentationData): ShapeBounds | null => {
  const size = getSlideSize(pres);
  if (!size) return null;
  const w = inches(1.5);
  const h = inches(0.4);
  const margin = inches(0.25);
  return {
    x: emu(Math.max(0, size.width - w - margin)),
    y: emu(Math.max(0, size.height - h - margin)),
    w,
    h,
  };
};

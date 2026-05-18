// Slide title convenience.

import { SHAPE_SNAPSHOT, type SlideData } from '../_internal-symbols.ts';
import { findSlidePlaceholder, setShapeText } from './shapes.ts';

// ---------------------------------------------------------------------------
// Slide title convenience.
//
// Most decks bind their title placeholder to `type="title"` or `type="ctrTitle"`
// (the latter is the centered hero title on a "Title Slide" layout).
// These two helpers cover ~90% of the "set the slide title" use case.

/**
 * Returns the slide's title text, or `null` if neither a `title` nor
 * a `ctrTitle` placeholder is present.
 */
export const getSlideTitle = (slide: SlideData): string | null => {
  const titleShape =
    findSlidePlaceholder(slide, 'title') ?? findSlidePlaceholder(slide, 'ctrTitle');
  if (titleShape === null) return null;
  return titleShape[SHAPE_SNAPSHOT].text ?? null;
};

/**
 * Returns the slide's body text, or `null` if no `body` placeholder
 * is present. Mirror of `getSlideTitle`; pairs with `setSlideBody`.
 */
export const getSlideBody = (slide: SlideData): string | null => {
  const bodyShape = findSlidePlaceholder(slide, 'body');
  if (bodyShape === null) return null;
  return bodyShape[SHAPE_SNAPSHOT].text ?? null;
};

/**
 * Sets the slide's title text. Looks for a `title` placeholder first,
 * falling back to `ctrTitle`. Throws if neither exists — the slide's
 * layout has no title slot.
 */
export const setSlideTitle = (slide: SlideData, title: string): void => {
  const titleShape =
    findSlidePlaceholder(slide, 'title') ?? findSlidePlaceholder(slide, 'ctrTitle');
  if (titleShape === null) {
    throw new Error('setSlideTitle: slide has no title / ctrTitle placeholder');
  }
  setShapeText(titleShape, title);
};

/**
 * Bulk-fills slide placeholders by type token. Each entry in
 * `byType` maps a `<p:ph type>` token (e.g. `'title'`, `'body'`,
 * `'ftr'`, `'dt'`) to the text to set. Silently skips entries
 * whose placeholder isn't present on the slide.
 *
 * Useful for template-fill workflows where the caller has all the
 * data in one struct.
 */
export const setSlidePlaceholders = (
  slide: SlideData,
  byType: Readonly<Record<string, string>>,
): void => {
  for (const [type, text] of Object.entries(byType)) {
    const shape =
      type === 'title'
        ? (findSlidePlaceholder(slide, 'title') ?? findSlidePlaceholder(slide, 'ctrTitle'))
        : findSlidePlaceholder(slide, type);
    if (shape !== null) setShapeText(shape, text);
  }
};

/**
 * Writes `text` into the first body placeholder on the slide.
 * Newlines start a new paragraph (each becomes its own bullet on
 * layouts that bullet their body placeholder).
 *
 * Throws when the slide has no body placeholder — pair with
 * `findSlideLayoutByType(pres, 'obj')` / `'tx'` to add the slide
 * onto a layout that has one.
 */
export const setSlideBody = (slide: SlideData, text: string): void => {
  const bodyShape = findSlidePlaceholder(slide, 'body');
  if (bodyShape === null) {
    throw new Error('setSlideBody: slide has no body placeholder');
  }
  setShapeText(bodyShape, text);
};

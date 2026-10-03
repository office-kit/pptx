// Editing a slide layout: what Google Slides calls the theme builder.
//
// A layout is a `<p:cSld>` like a slide is, so the background writer and
// the geometry writers are the slide ones. What differs is the write-back
// path — a layout handle commits into its own package part — and the fact
// that every slide bound to the layout inherits the result.

import { setPosition, setSize, setSolidFill } from '../../internal/drawingml/index.ts';
import { NS, type XmlElement, attr, firstChildElement, qname } from '../../internal/xml/index.ts';
import { LAYOUT_DOCUMENT, LAYOUT_PART, type SlideLayoutData } from '../_internal-symbols.ts';
import { NAME_CSLD, commitLayoutData } from './_helpers.ts';
import { writeBackgroundPr } from './slide-background.ts';
import type { ShapeBounds } from './shapes.ts';

const ATTR_NAME = qname('', 'name', '');

const requireCSld = (layout: SlideLayoutData, fn: string): XmlElement => {
  const cSld = firstChildElement(layout[LAYOUT_DOCUMENT].root, NAME_CSLD);
  if (cSld === null) throw new Error(`${fn}: the layout has no <p:cSld>`);
  return cSld;
};

/**
 * Renames the layout. This is the name PowerPoint shows in the
 * "New Slide" gallery and in the layout pane, so it is what an author
 * picks a layout by — it is not an identifier, and nothing references
 * it.
 */
export const setSlideLayoutName = (layout: SlideLayoutData, name: string): void => {
  const cSld = requireCSld(layout, 'setSlideLayoutName');
  cSld.attrs = cSld.attrs.filter((a) => a.name.localName !== 'name');
  cSld.attrs.push(attr(ATTR_NAME, name));
  commitLayoutData(layout);
};

/**
 * Gives the layout a solid background, replacing whatever it had.
 * Every slide on this layout that does not set a background of its own
 * follows.
 *
 * `color` takes the same forms as `setSlideBackground`: `#RRGGBB` or
 * `scheme:accent1`.
 */
export const setSlideLayoutBackground = (layout: SlideLayoutData, color: string): void => {
  const cSld = requireCSld(layout, 'setSlideLayoutBackground');
  writeBackgroundPr(cSld, (bgPr) => setSolidFill(bgPr, color));
  commitLayoutData(layout);
};

/**
 * Drops the layout's own background so it inherits the master's again.
 */
export const clearSlideLayoutBackground = (layout: SlideLayoutData): void => {
  const cSld = requireCSld(layout, 'clearSlideLayoutBackground');
  cSld.children = cSld.children.filter(
    (c) => !(c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'bg'),
  );
  commitLayoutData(layout);
};

// Same predicate as `getSlideLayoutPlaceholders`, so index `n` here and
// index `n` there are the same slot.
const placeholderShapes = (layout: SlideLayoutData) =>
  layout[LAYOUT_PART].shapes.filter(
    (shape) => shape.placeholderType !== null || shape.placeholderIdx !== null,
  );

/**
 * Moves and resizes one of the layout's placeholder slots, identified by
 * its position in `getSlideLayoutPlaceholders`.
 *
 * Slides inherit the new box wherever they have not pushed the
 * placeholder around themselves: a slide placeholder carrying its own
 * `<a:xfrm>` keeps it, exactly as PowerPoint behaves when you edit a
 * layout under slides that were already nudged.
 */
export const setSlideLayoutPlaceholderBounds = (
  layout: SlideLayoutData,
  index: number,
  bounds: ShapeBounds,
): void => {
  const shapes = placeholderShapes(layout);
  const shape = shapes[index];
  if (shape === undefined) {
    throw new Error(
      `setSlideLayoutPlaceholderBounds: the layout has ${shapes.length} placeholder(s); no index ${index}`,
    );
  }
  setPosition(shape.element, shape.kind, bounds.x, bounds.y);
  setSize(shape.element, shape.kind, bounds.w, bounds.h);
  commitLayoutData(layout);
};

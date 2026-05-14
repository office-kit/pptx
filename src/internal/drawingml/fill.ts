// Solid-fill mutation for shapes and slide backgrounds.
//
// The spec places the fill element after geometry (`prstGeom`/`custGeom`)
// and before the line (`a:ln`) on a shape's `<p:spPr>`. We re-insert
// rather than append so re-applying a fill stays clean across calls.
//
// `setSolidFill(host, color)` accepts the host element that wraps the
// fill (`p:spPr` for shapes, `p:bgPr` for backgrounds). It removes any
// previous fill choice (`noFill`/`solidFill`/`gradFill`/`blipFill`/
// `pattFill`/`grpFill`) before inserting the new `solidFill`.

import { NS, type XmlElement, elem, qname } from '../xml/index.ts';
import { buildColorElement } from './color.ts';

const NAME_SOLID_FILL = qname('a', 'solidFill', NS.dml);
const NAME_NO_FILL = qname('a', 'noFill', NS.dml);

const FILL_CHOICE_LOCAL_NAMES = new Set([
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
]);

const removeAnyFill = (host: XmlElement): void => {
  host.children = host.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        FILL_CHOICE_LOCAL_NAMES.has(c.name.localName)
      ),
  );
};

/**
 * Returns the index where the fill should be inserted on `host` per the
 * schema's child-element sequence. For `<p:spPr>` that's right after the
 * geometry element; for `<p:bgPr>` it's the start.
 *
 * The schema is more nuanced (xfrm → geometry → fill → ln → effects →
 * scene3d → sp3d → extLst) but PowerPoint tolerates any of those slots
 * being absent, and we only care about staying ahead of `a:ln`.
 */
const fillInsertionIndex = (host: XmlElement): number => {
  for (let i = 0; i < host.children.length; i++) {
    const c = host.children[i];
    if (c?.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (c.name.localName === 'ln') return i;
    if (c.name.localName === 'effectLst' || c.name.localName === 'effectDag') return i;
    if (c.name.localName === 'scene3d' || c.name.localName === 'sp3d') return i;
    if (c.name.localName === 'extLst') return i;
  }
  return host.children.length;
};

/** Sets `<a:solidFill>` on `host`, removing any previous fill choice. */
export const setSolidFill = (host: XmlElement, color: string): void => {
  removeAnyFill(host);
  const fill = elem(NAME_SOLID_FILL, { children: [buildColorElement(color)] });
  host.children.splice(fillInsertionIndex(host), 0, fill);
};

/** Sets `<a:noFill>` on `host`, removing any previous fill choice. */
export const setNoFill = (host: XmlElement): void => {
  removeAnyFill(host);
  host.children.splice(fillInsertionIndex(host), 0, elem(NAME_NO_FILL));
};

/**
 * Removes any fill choice from `host` entirely. The shape then inherits
 * its fill from the layout / master placeholder it descends from.
 */
export const clearFill = (host: XmlElement): void => {
  removeAnyFill(host);
};

// Helpers for mutating `<p:pic>` shapes. Reads the `r:embed` rId from the
// `<a:blip>` and exposes a setter that points the same rId at a different
// relationship target.
//
// The actual media-part swap lives in `parts/`; this module is purely the
// XML manipulation side.

import { NS, getAttrValue, qname } from '../xml/index.ts';
import type { XmlElement } from '../xml/index.ts';
import { firstChildElement } from '../xml/index.ts';

const NAME_BLIP_FILL = qname('p', 'blipFill', NS.pml);
const NAME_BLIP = qname('a', 'blip', NS.dml);
const ATTR_R_EMBED = qname('r', 'embed', NS.officeDocRels);

/**
 * Returns the `r:embed` rId on this picture shape's `a:blip`. Returns null
 * if the shape has no blip fill, no blip, or no embed attribute (which
 * happens for picture shapes that use an external `r:link` instead).
 */
export const getPictureEmbedRId = (picElement: XmlElement): string | null => {
  const blipFill = firstChildElement(picElement, NAME_BLIP_FILL);
  if (blipFill === null) return null;
  const blip = firstChildElement(blipFill, NAME_BLIP);
  if (blip === null) return null;
  return getAttrValue(blip, ATTR_R_EMBED);
};

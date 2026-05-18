// Hyperlink mutation: wrap every `<a:r>` in a `<a:txBody>` so its
// `<a:rPr>` carries `<a:hlinkClick r:id="rIdN"/>`. The rId must already
// exist on the part's `.rels` (the caller is responsible for allocating it
// and adding the rel before calling here).

import { NS, type XmlElement, attr, elem, firstChildElement, qname } from '../xml/index.ts';

const NAME_R = qname('a', 'r', NS.dml);
const NAME_RPR = qname('a', 'rPr', NS.dml);
const NAME_HLINK_CLICK = qname('a', 'hlinkClick', NS.dml);
const ATTR_R_ID = qname('r', 'id', NS.officeDocRels);

/**
 * Sets `<a:hlinkClick r:id="rIdN" [tooltip="…"]/>` inside the `<a:rPr>`
 * of every run in `txBody`. Pass `null` for `rId` to remove an existing
 * hyperlink. When `tooltip` is `undefined` no `tooltip=` attribute is
 * written; when it's a string the attribute is set. Creates `<a:rPr>`
 * if absent on a run.
 */
export const applyHyperlinkToAllRuns = (
  txBody: XmlElement,
  rId: string | null,
  tooltip?: string,
): void => {
  for (const p of txBody.children) {
    if (p.kind !== 'element' || p.name.namespaceURI !== NS.dml || p.name.localName !== 'p') {
      continue;
    }
    for (const r of p.children) {
      if (r.kind !== 'element' || r.name.namespaceURI !== NS.dml || r.name.localName !== 'r') {
        continue;
      }
      let rPr = firstChildElement(r, NAME_RPR);
      if (rPr === null) {
        rPr = elem(NAME_RPR);
        r.children.unshift(rPr);
      }
      // Drop any existing hlinkClick.
      rPr.children = rPr.children.filter(
        (c) =>
          !(
            c.kind === 'element' &&
            c.name.namespaceURI === NS.dml &&
            c.name.localName === 'hlinkClick'
          ),
      );
      if (rId !== null) {
        const attrs = [attr(ATTR_R_ID, rId)];
        if (tooltip !== undefined) {
          attrs.push(attr(qname('', 'tooltip', ''), tooltip));
        }
        // Per the schema, hlinkClick is one of the last children of rPr —
        // it follows the fill/typeface children. Append.
        rPr.children.push(elem(NAME_HLINK_CLICK, { attrs }));
      }
    }
  }
  // Force-touch NAME_R so it isn't elided as unused.
  void NAME_R;
};

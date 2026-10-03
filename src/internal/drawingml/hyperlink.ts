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
      applyHyperlinkToProperties(rPr, rId, tooltip);
    }
  }
  // Force-touch NAME_R so it isn't elided as unused.
  void NAME_R;
};

export const applyHyperlinkToProperties = (
  rPr: XmlElement,
  rId: string | null,
  tooltip?: string,
): void => {
  const attrs = rId === null ? null : [attr(ATTR_R_ID, rId)];
  if (attrs && tooltip !== undefined) attrs.push(attr(qname('', 'tooltip', ''), tooltip));
  replaceClickHyperlink(rPr, attrs ? elem(NAME_HLINK_CLICK, { attrs }) : null);
};

/** Replace a click link on rPr or cNvPr without disturbing its ordered siblings. */
export const replaceClickHyperlink = (parent: XmlElement, link: XmlElement | null): void => {
  parent.children = parent.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        c.name.localName === 'hlinkClick'
      ),
  );
  if (!link) return;
  // CT_TextCharacterProperties and CT_NonVisualDrawingProps both require
  // click links before hover links and extension lists; rPr also has rtl.
  const next = parent.children.findIndex(
    (c) =>
      c.kind === 'element' &&
      c.name.namespaceURI === NS.dml &&
      ['hlinkMouseOver', 'hlinkHover', 'rtl', 'extLst'].includes(c.name.localName),
  );
  parent.children.splice(next < 0 ? parent.children.length : next, 0, link);
};

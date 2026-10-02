import { readDrawingmlPercentage } from './_drawingml-percentage.ts';
import {
  NS,
  type XmlElement,
  firstChildElement,
  getAttrValue,
  qname,
} from '../../internal/xml/index.ts';

/** One paragraph-property layer's bullet detail; undefined means absent. */
export interface BulletStyleLayer {
  color?: string | null;
  colorFollowText?: boolean;
  sizePct?: number | null;
  sizePts?: number | null;
  sizeFollowText?: boolean;
  font?: string | null;
  fontFollowText?: boolean;
}

/** Read only bullet detail authored by this pPr, preserving explicit Tx sentinels. */
export const readBulletStyleLayer = (
  pPr: XmlElement | null,
  resolveColor: (element: XmlElement) => string | null,
): BulletStyleLayer => {
  if (!pPr) return {};
  const out: BulletStyleLayer = {};
  const buClr = firstChildElement(pPr, qname('a', 'buClr', NS.dml));
  const buClrTx = firstChildElement(pPr, qname('a', 'buClrTx', NS.dml));
  if (buClrTx) {
    out.color = null;
    out.colorFollowText = true;
  } else if (buClr) {
    const color = buClr.children.find(
      (child): child is XmlElement =>
        child.kind === 'element' && child.name.namespaceURI === NS.dml,
    );
    if (color) {
      out.color = resolveColor(color);
      out.colorFollowText = false;
    }
  }

  const buSzTx = firstChildElement(pPr, qname('a', 'buSzTx', NS.dml));
  const buSzPct = firstChildElement(pPr, qname('a', 'buSzPct', NS.dml));
  const buSzPts = firstChildElement(pPr, qname('a', 'buSzPts', NS.dml));
  if (buSzTx) {
    out.sizePct = null;
    out.sizePts = null;
    out.sizeFollowText = true;
  } else if (buSzPct) {
    const value = getAttrValue(buSzPct, qname('', 'val', ''));
    const size = value === null ? Number.NaN : readDrawingmlPercentage(value, Number.NaN);
    if (Number.isFinite(size)) {
      out.sizePct = size;
      out.sizePts = null;
      out.sizeFollowText = false;
    }
  } else if (buSzPts) {
    const value = getAttrValue(buSzPts, qname('', 'val', ''));
    const size = value === null ? Number.NaN : Number.parseInt(value, 10) / 100;
    if (Number.isFinite(size)) {
      out.sizePct = null;
      out.sizePts = size;
      out.sizeFollowText = false;
    }
  }

  const buFontTx = firstChildElement(pPr, qname('a', 'buFontTx', NS.dml));
  const buFont = firstChildElement(pPr, qname('a', 'buFont', NS.dml));
  if (buFontTx) {
    out.font = null;
    out.fontFollowText = true;
  } else if (buFont) {
    out.font = getAttrValue(buFont, qname('', 'typeface', ''));
    out.fontFollowText = false;
  }
  return out;
};

import {
  NS,
  type XmlElement,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  qname,
} from '../../internal/xml/index.ts';

const NAME_ALPHA_MOD_FIX_FN = qname('a', 'alphaModFix', NS.dml);
const ATTR_AMT_FN = qname('', 'amt', '');

export const readImageOpacity = (blip: XmlElement): number | null => {
  const alpha = firstChildElement(blip, qname('a', 'alphaModFix', NS.dml));
  if (!alpha) return null;
  const amt = getAttrValue(alpha, qname('', 'amt', ''));
  if (amt === null) return 1;
  const n = Number.parseInt(amt, 10);
  if (!Number.isFinite(n)) return null;
  return n / 100000;
};

export const writeImageOpacity = (blip: XmlElement, opacity: number | null): void => {
  if (opacity !== null && (!Number.isFinite(opacity) || opacity < 0 || opacity > 1)) {
    throw new RangeError(`opacity must be in [0, 1], got ${opacity}`);
  }

  blip.children = blip.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        c.name.localName === 'alphaModFix'
      ),
  );

  if (opacity !== null) {
    blip.children.push(
      elem(NAME_ALPHA_MOD_FIX_FN, {
        attrs: [attr(ATTR_AMT_FN, String(Math.round(opacity * 100000)))],
      }),
    );
  }
};

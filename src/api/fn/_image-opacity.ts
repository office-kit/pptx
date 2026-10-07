import {
  NS,
  type XmlElement,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  qname,
} from '../../internal/xml/index.ts';
import { readDrawingmlPercentage } from './_drawingml-percentage.ts';

const NAME_ALPHA_MOD_FIX_FN = qname('a', 'alphaModFix', NS.dml);
const ATTR_AMT_FN = qname('', 'amt', '');

/**
 * Adds a CT_Blip effect. Effects form an unordered choice, but they must all
 * precede the blip's `<a:extLst>` (where PowerPoint keeps `a14:imgProps` and
 * `a14:useLocalDpi`).
 */
export const insertBlipEffect = (blip: XmlElement, effect: XmlElement): void => {
  const extensionIndex = blip.children.findIndex(
    (child) =>
      child.kind === 'element' &&
      child.name.namespaceURI === NS.dml &&
      child.name.localName === 'extLst',
  );
  if (extensionIndex === -1) blip.children.push(effect);
  else blip.children.splice(extensionIndex, 0, effect);
};

export const readImageOpacity = (blip: XmlElement): number | null => {
  const alpha = firstChildElement(blip, qname('a', 'alphaModFix', NS.dml));
  if (!alpha) return null;
  const amt = getAttrValue(alpha, qname('', 'amt', ''));
  if (amt === null) return 1;
  const value = readDrawingmlPercentage(amt, Number.NaN);
  return Number.isFinite(value) ? value : null;
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
    insertBlipEffect(
      blip,
      elem(NAME_ALPHA_MOD_FIX_FN, {
        attrs: [attr(ATTR_AMT_FN, String(Math.round(opacity * 100000)))],
      }),
    );
  }
};

// The reference desktop app's built-in Picture Styles. The reference desktop app compiles them into
// the application and writes each one out as plain `p:spPr` content — geometry,
// an optional fill, `a:ln`, `a:effectLst` and sometimes `a:scene3d` / `a:sp3d`,
// all with literal sRGB colors — so a deck never names the style it used.

import { NS, type XmlElement, type XmlNode, parseXml } from '../xml/index.ts';
import {
  BUILTIN_PICTURE_STYLE_SP_PR,
  BUILTIN_PICTURE_STYLES,
} from './builtin-picture-styles.generated.ts';

export { BUILTIN_PICTURE_STYLES };

/** English name of one of the reference desktop app's built-in picture styles (its gallery tooltip). */
export type BuiltinPictureStyleName = (typeof BUILTIN_PICTURE_STYLES)[number];

const styleIndex = new Map<string, number>(BUILTIN_PICTURE_STYLES.map((name, i) => [name, i]));

export const isBuiltinPictureStyleName = (name: string): name is BuiltinPictureStyleName =>
  styleIndex.has(name);

/** Fresh `p:spPr` children (after `a:xfrm`) for the style, ready to insert. */
export const buildPictureStyleChildren = (name: BuiltinPictureStyleName): XmlElement[] => {
  const xml = BUILTIN_PICTURE_STYLE_SP_PR[styleIndex.get(name)!]!;
  const root = parseXml(`<a:r xmlns:a="${NS.dml}">${xml}</a:r>`).root;
  return root.children.filter((child): child is XmlElement => child.kind === 'element');
};

const isDml = (node: XmlNode, localName: string): node is XmlElement =>
  node.kind === 'element' && node.name.namespaceURI === NS.dml && node.name.localName === localName;

/**
 * Replaces everything a picture style controls in `spPr` — all children but
 * `a:xfrm` and `a:extLst` — with the style's content, as the reference desktop app does.
 */
export const applyPictureStyle = (spPr: XmlElement, name: BuiltinPictureStyleName): void => {
  const xfrm = spPr.children.filter((child) => isDml(child, 'xfrm'));
  const extLst = spPr.children.filter((child) => isDml(child, 'extLst'));
  spPr.children = [...xfrm, ...buildPictureStyleChildren(name), ...extLst];
};

const elementChildren = (element: XmlElement): XmlElement[] =>
  element.children.filter((child): child is XmlElement => child.kind === 'element');

// Compared by namespace and local name so a deck that binds DrawingML to
// another prefix still matches; attribute order carries no meaning in XML.
const sameElement = (a: XmlElement, b: XmlElement): boolean => {
  if (a.name.namespaceURI !== b.name.namespaceURI || a.name.localName !== b.name.localName)
    return false;
  if (a.attrs.length !== b.attrs.length) return false;
  const key = (attr: XmlElement['attrs'][number]) =>
    `${attr.name.namespaceURI} ${attr.name.localName}=${attr.value}`;
  const attrs = new Set(a.attrs.map(key));
  if (!b.attrs.every((attr) => attrs.has(key(attr)))) return false;
  const ac = elementChildren(a);
  const bc = elementChildren(b);
  return ac.length === bc.length && ac.every((child, i) => sameElement(child, bc[i]!));
};

// Detection runs on every selection change in an editor; parse each style once.
const parsedStyles = new Map<BuiltinPictureStyleName, readonly XmlElement[]>();
const comparable = (name: BuiltinPictureStyleName): readonly XmlElement[] => {
  let style = parsedStyles.get(name);
  if (!style) parsedStyles.set(name, (style = buildPictureStyleChildren(name)));
  return style;
};

/**
 * The built-in style whose content `spPr` carries exactly (ignoring `a:xfrm`
 * and `a:extLst`), or `null`. Any later edit — a different border width, an
 * extra effect — means the picture no longer carries a built-in style.
 */
export const detectPictureStyle = (spPr: XmlElement): BuiltinPictureStyleName | null => {
  const own = elementChildren(spPr).filter(
    (child) => !isDml(child, 'xfrm') && !isDml(child, 'extLst'),
  );
  for (const name of BUILTIN_PICTURE_STYLES) {
    const style = comparable(name);
    if (style.length === own.length && style.every((child, i) => sameElement(child, own[i]!)))
      return name;
  }
  return null;
};

import { partName } from '../../internal/opc/index.ts';
import {
  NS,
  cloneElement,
  getAttrValue,
  type XmlElement,
  attr,
  elem,
  firstChildElement,
  parseXml,
  qname,
  serializeXml,
} from '../../internal/xml/index.ts';
import { INTERNAL_PACKAGE, type SlideData } from '../_internal-symbols.ts';
import { NAME_CSLD, decode, encode } from './_helpers.ts';
import { getSlideMasterPartName } from './shape-read-base.ts';
import { readBackgroundStyle } from './background-style-read.ts';
import { backgroundOfCSld, type SlideBackground } from './slide-background.ts';
import { parseGradFill } from './shape-gradient-read.ts';
import { resolveDrawingColor, resolveDrawingColorOpacity } from './shape-color.ts';
import type { ReadGradientFill } from '../../internal/drawingml/index.ts';
import { themeRootFromPackage } from './theme.ts';

/**
 * Choose one of PowerPoint's twelve Background Styles (numbered row by row).
 * Changes the owning master, so every slide using that master inherits the
 * style. Existing slide and layout background overrides remain in place.
 * The theme must provide the referenced background fill style.
 */
export const setSlideMasterBackgroundStyle = (slide: SlideData, style: number): void => {
  const columns = 4;
  const styleCount = 12;
  if (!Number.isInteger(style) || style < 1 || style > styleCount)
    throw new RangeError('Background style must be an integer from 1 to 12');
  const masterName = getSlideMasterPartName(slide);
  if (!masterName) throw new Error('Slide has no master');
  const pkg = slide[INTERNAL_PACKAGE];
  const name = partName(masterName);
  const part = pkg.getPart(name);
  if (!part) throw new Error(`Missing slide master: ${name}`);
  const document = parseXml(decode(part.data));
  const row = Math.floor((style - 1) / columns);
  const theme = themeRootFromPackage(pkg, name);
  const elements = theme && firstChildElement(theme, qname('a', 'themeElements', NS.dml));
  const scheme = elements && firstChildElement(elements, qname('a', 'fmtScheme', NS.dml));
  const fills = scheme && firstChildElement(scheme, qname('a', 'bgFillStyleLst', NS.dml));
  if (!fills?.children.filter((child) => child.kind === 'element')[row])
    throw new Error(`Theme has no background fill style ${row + 1}`);
  configureStyle(document.root, style);
  part.data = encode(serializeXml(document));
};

function configureStyle(root: XmlElement, style: number): void {
  const columns = 4;
  const backgroundIndexStart = 1001;
  const row = Math.floor((style - 1) / columns);
  const column = (style - 1) % columns;
  const cSld = firstChildElement(root, NAME_CSLD);
  const colorMap = firstChildElement(root, qname('p', 'clrMap', NS.pml));
  if (!cSld || !colorMap) throw new Error('Slide master requires cSld and clrMap');
  const reference = elem(qname('p', 'bgRef', NS.pml), {
    attrs: [attr(qname('', 'idx', ''), String(backgroundIndexStart + row))],
    children: [
      elem(qname('a', 'schemeClr', NS.dml), {
        attrs: [attr(qname('', 'val', ''), column === 0 || column === 3 ? 'bg1' : 'bg2')],
      }),
    ],
  });
  const background = elem(qname('p', 'bg', NS.pml), { children: [reference] });
  cSld.children = cSld.children.filter(
    (child) =>
      !(
        child.kind === 'element' &&
        child.name.namespaceURI === NS.pml &&
        child.name.localName === 'bg'
      ),
  );
  cSld.children.unshift(background);
  // Mac PowerPoint swaps both text/background pairs for the two dark columns,
  // preserving the theme's accent and hyperlink mappings.
  const dark = column >= 2;
  const pairs: Record<string, string> = {
    bg1: dark ? 'dk1' : 'lt1',
    tx1: dark ? 'lt1' : 'dk1',
    bg2: dark ? 'dk2' : 'lt2',
    tx2: dark ? 'lt2' : 'dk2',
  };
  colorMap.attrs = colorMap.attrs.map((attribute) =>
    attribute.name.namespaceURI === '' && pairs[attribute.name.localName]
      ? { ...attribute, value: pairs[attribute.name.localName]! }
      : attribute,
  );
}

/** A theme background preset; complex fill kinds without a gradient retain their kind. */
export interface SlideMasterBackgroundStyle {
  readonly style: number;
  readonly selected: boolean;
  readonly fill: SlideBackground;
  readonly gradient: ReadGradientFill | null;
}

/**
 * Read the owning master's available Background Styles without changing the
 * presentation. Solid colors and gradient stop resolvedColor values include
 * the preset's color mapping and theme transforms. Missing theme rows are omitted.
 * Selection describes the master, independently of slide/layout overrides.
 */
export const getSlideMasterBackgroundStyles = (slide: SlideData): SlideMasterBackgroundStyle[] => {
  const masterName = getSlideMasterPartName(slide);
  if (!masterName) return [];
  const pkg = slide[INTERNAL_PACKAGE];
  const name = partName(masterName);
  const part = pkg.getPart(name);
  if (!part) throw new Error(`Missing slide master: ${name}`);
  const root = parseXml(decode(part.data)).root;
  const theme = themeRootFromPackage(pkg, name);
  const elements = theme && firstChildElement(theme, qname('a', 'themeElements', NS.dml));
  const scheme = elements && firstChildElement(elements, qname('a', 'fmtScheme', NS.dml));
  const fills = scheme && firstChildElement(scheme, qname('a', 'bgFillStyleLst', NS.dml));
  const rows = Math.min(3, fills?.children.filter((child) => child.kind === 'element').length ?? 0);
  const signature = (master: XmlElement): string => {
    const cSld = firstChildElement(master, NAME_CSLD);
    const bg = cSld && firstChildElement(cSld, qname('p', 'bg', NS.pml));
    const map = firstChildElement(master, qname('p', 'clrMap', NS.pml));
    const reference = bg && firstChildElement(bg, qname('p', 'bgRef', NS.pml));
    const color = reference && firstChildElement(reference, qname('a', 'schemeClr', NS.dml));
    return JSON.stringify([
      reference && getAttrValue(reference, qname('', 'idx', '')),
      color && getAttrValue(color, qname('', 'val', '')),
      color?.children.filter((child) => child.kind === 'element'),
      ...['bg1', 'bg2', 'tx1', 'tx2'].map((key) => map && getAttrValue(map, qname('', key, ''))),
    ]);
  };
  const selected = signature(root);
  return Array.from({ length: rows * 4 }, (_, index) => {
    const candidate = cloneElement(root);
    configureStyle(candidate, index + 1);
    const context = readBackgroundStyle(pkg, name, candidate);
    const cSld = firstChildElement(candidate, NAME_CSLD)!;
    let fill = backgroundOfCSld(cSld, context.properties);
    const solid =
      context.properties && firstChildElement(context.properties, qname('a', 'solidFill', NS.dml));
    const colorElement = solid?.children.find((child) => child.kind === 'element');
    if (fill.kind === 'solid' && colorElement?.kind === 'element') {
      const color = resolveDrawingColor(colorElement, context.theme, context.colorMap);
      const opacity = resolveDrawingColorOpacity(colorElement);
      if (color !== null) fill = { kind: 'solid', color, ...(opacity !== null ? { opacity } : {}) };
    }
    const gradient =
      context.properties && firstChildElement(context.properties, qname('a', 'gradFill', NS.dml));
    return {
      style: index + 1,
      selected: selected === signature(candidate),
      fill,
      gradient: gradient ? parseGradFill(gradient, context) : null,
    };
  });
};

import { partName } from '../../internal/opc/index.ts';
import {
  NS,
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
  const backgroundIndexStart = 1001;
  if (!Number.isInteger(style) || style < 1 || style > styleCount)
    throw new RangeError('Background style must be an integer from 1 to 12');
  const masterName = getSlideMasterPartName(slide);
  if (!masterName) throw new Error('Slide has no master');
  const pkg = slide[INTERNAL_PACKAGE];
  const name = partName(masterName);
  const part = pkg.getPart(name);
  if (!part) throw new Error(`Missing slide master: ${name}`);
  const document = parseXml(decode(part.data));
  const cSld = firstChildElement(document.root, NAME_CSLD);
  const colorMap = firstChildElement(document.root, qname('p', 'clrMap', NS.pml));
  if (!cSld || !colorMap) throw new Error('Slide master requires cSld and clrMap');
  const row = Math.floor((style - 1) / columns);
  const column = (style - 1) % columns;
  const theme = themeRootFromPackage(pkg, name);
  const elements = theme && firstChildElement(theme, qname('a', 'themeElements', NS.dml));
  const scheme = elements && firstChildElement(elements, qname('a', 'fmtScheme', NS.dml));
  const fills = scheme && firstChildElement(scheme, qname('a', 'bgFillStyleLst', NS.dml));
  if (!fills?.children.filter((child) => child.kind === 'element')[row])
    throw new Error(`Theme has no background fill style ${row + 1}`);
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
  part.data = encode(serializeXml(document));
};

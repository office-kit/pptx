import { resolveTarget, type PartName } from '../../internal/opc/index.ts';
import type { OpcPackage } from '../../internal/parts/index.ts';
import { REL_TYPES } from '../../internal/presentationml/index.ts';
import {
  NS,
  cloneElement,
  elem,
  firstChildElement,
  getAttrValue,
  parseXml,
  qname,
  type XmlElement,
} from '../../internal/xml/index.ts';
import { decode } from './_helpers.ts';
import { themeFromPackage, themeRootFromPackage } from './theme.ts';

const element = (root: XmlElement | null, name: string, ns: string = NS.pml): XmlElement | null =>
  root ? firstChildElement(root, qname('', name, ns)) : null;

/** Resolve a background style in its owning master's theme, without mutating XML. @internal */
export function readBackgroundStyle(pkg: OpcPackage, from: PartName, root: XmlElement) {
  const bg = element(element(root, 'cSld'), 'bg');
  let properties = element(bg, 'bgPr');
  const reference = element(bg, 'bgRef');
  if (!properties && !reference) return { properties: null, theme: null, colorMap: {} };
  const roots = [root];
  let masterName = from;
  for (const type of [REL_TYPES.slideLayout, REL_TYPES.slideMaster]) {
    const rel = pkg
      .getRels(masterName)
      ?.items.find((item) => item.type === type && item.targetMode !== 'External');
    if (!rel) continue;
    masterName = resolveTarget(masterName, rel.target);
    const part = pkg.getPart(masterName);
    if (part) roots.push(parseXml(decode(part.data)).root);
  }
  let colorMap: Record<string, string> = {};
  let masterColorMap: Record<string, string> = {};
  for (const source of [...roots].reverse()) {
    const masterMapping = element(source, 'clrMap');
    const override = element(source, 'clrMapOvr');
    if (element(override, 'masterClrMapping', NS.dml)) colorMap = { ...masterColorMap };
    const mapping = masterMapping ?? element(override, 'overrideClrMapping', NS.dml);
    for (const attribute of mapping?.attrs ?? []) {
      if (attribute.name.namespaceURI === '') colorMap[attribute.name.localName] = attribute.value;
    }
    if (masterMapping) masterColorMap = { ...colorMap };
  }
  const theme = themeFromPackage(pkg, masterName);
  if (!properties && reference) {
    const index = Number(getAttrValue(reference, qname('', 'idx', '')));
    const themeRoot = themeRootFromPackage(pkg, masterName);
    const matrix = element(element(themeRoot, 'themeElements', NS.dml), 'fmtScheme', NS.dml);
    const backgroundStyleStart = 1001;
    const list = element(
      matrix,
      index >= backgroundStyleStart ? 'bgFillStyleLst' : 'fillStyleLst',
      NS.dml,
    );
    const styles = list?.children.filter((child) => child.kind === 'element') ?? [];
    const style = styles[index >= backgroundStyleStart ? index - backgroundStyleStart : index - 1];
    if (style?.kind === 'element') {
      const fill = cloneElement(style);
      const placeholder = reference.children.find(
        (child) => child.kind === 'element' && child.name.namespaceURI === NS.dml,
      );
      if (placeholder?.kind === 'element') {
        const replacePlaceholder = (parent: XmlElement): void => {
          parent.children = parent.children.map((child) => {
            if (child.kind !== 'element') return child;
            if (
              child.name.namespaceURI === NS.dml &&
              child.name.localName === 'schemeClr' &&
              getAttrValue(child, qname('', 'val', '')) === 'phClr'
            ) {
              const color = cloneElement(placeholder);
              color.children.push(
                ...child.children.map((item) =>
                  item.kind === 'element' ? cloneElement(item) : { ...item },
                ),
              );
              return color;
            }
            replacePlaceholder(child);
            return child;
          });
        };
        replacePlaceholder(fill);
      }
      properties = elem(qname('p', 'bgPr', NS.pml), { children: [fill] });
    }
  }
  return { properties, theme, colorMap };
}

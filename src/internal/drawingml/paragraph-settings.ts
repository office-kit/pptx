import {
  NS,
  attr,
  elem,
  firstChildElement,
  insertChildByRank,
  qname,
  type XmlElement,
} from '../xml/index.ts';
import { alignToken, type ParagraphAlignment } from './text-body-mutation.ts';

export interface ParagraphSettings {
  align?: ParagraphAlignment;
  leftEmu?: number;
  rightEmu?: number;
  firstLineEmu?: number;
  beforePts?: number;
  afterPts?: number;
  lineSpacing?: { kind: 'pct' | 'pts'; value: number };
}

/** Build a validated patch separately so invalid input cannot partially change text. */
export function paragraphSettingsPatch(settings: ParagraphSettings): XmlElement {
  const patch = elem(qname('a', 'pPr', NS.dml));
  const numeric = (value: number, min: number, max: number) => {
    if (!Number.isFinite(value) || value < min || value > max)
      throw new RangeError('Invalid paragraph setting.');
    return Math.round(value);
  };
  if (settings.align !== undefined)
    patch.attrs.push(attr(qname('', 'algn', ''), alignToken(settings.align, 'paragraphSettings')));
  for (const [key, name] of [
    ['leftEmu', 'marL'],
    ['rightEmu', 'marR'],
    ['firstLineEmu', 'indent'],
  ] as const) {
    const value = settings[key];
    if (value !== undefined)
      patch.attrs.push(
        attr(
          qname('', name, ''),
          String(numeric(value, name === 'indent' ? -51206400 : 0, 51206400)),
        ),
      );
  }
  const spacing = (name: string, kind: 'pct' | 'pts', value: number) => {
    const amount = numeric(value * (kind === 'pct' ? 100000 : 100), 0, 2147483647);
    patch.children.push(
      elem(qname('a', name, NS.dml), {
        children: [
          elem(qname('a', kind === 'pct' ? 'spcPct' : 'spcPts', NS.dml), {
            attrs: [attr(qname('', 'val', ''), String(amount))],
          }),
        ],
      }),
    );
  };
  if (settings.lineSpacing !== undefined) {
    if (!['pct', 'pts'].includes(settings.lineSpacing.kind))
      throw new Error('Invalid paragraph line spacing.');
    spacing('lnSpc', settings.lineSpacing.kind, settings.lineSpacing.value);
  }
  if (settings.beforePts !== undefined) spacing('spcBef', 'pts', settings.beforePts);
  if (settings.afterPts !== undefined) spacing('spcAft', 'pts', settings.afterPts);
  return patch;
}

export function applyParagraphSettingsPatch(paragraph: XmlElement, patch: XmlElement): void {
  let properties = firstChildElement(paragraph, qname('a', 'pPr', NS.dml));
  if (!properties) {
    properties = elem(qname('a', 'pPr', NS.dml));
    paragraph.children.unshift(properties);
  }
  for (const value of patch.attrs) {
    properties.attrs = properties.attrs.filter(
      (existing) =>
        existing.name.localName !== value.name.localName ||
        existing.name.namespaceURI !== value.name.namespaceURI,
    );
    properties.attrs.push({ ...value });
  }
  const rank = (child: XmlElement) =>
    ({ lnSpc: 0, spcBef: 1, spcAft: 2 })[child.name.localName] ?? 3;
  for (const child of patch.children) {
    if (child.kind !== 'element') continue;
    properties.children = properties.children.filter(
      (existing) =>
        existing.kind !== 'element' ||
        existing.name.localName !== child.name.localName ||
        existing.name.namespaceURI !== NS.dml,
    );
    insertChildByRank(properties, structuredClone(child), rank);
  }
}

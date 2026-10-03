import { INTERNAL_PACKAGE, type PresentationData } from '../_internal-symbols.ts';
import { partName, resolveTarget, nextRelId } from '../../internal/opc/index.ts';
import { REL_TYPES } from '../../internal/presentationml/relationship-types.ts';
import {
  NS,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  parseXml,
  qname,
  serializeXml,
  type XmlElement,
} from '../../internal/xml/index.ts';
import { PRES_PART_NAME, decode, encode } from './_helpers.ts';
import { getCustomShows } from './custom-shows.ts';
import { getSlides } from './slide-query.ts';

export interface SlideShowProperties {
  mode:
    | { kind: 'present' }
    | { kind: 'browse'; showScrollbar: boolean }
    | { kind: 'kiosk'; restart: number };
  slides:
    | { kind: 'all' }
    | { kind: 'range'; start: number; end: number }
    | { kind: 'customShow'; id: number };
  loop: boolean;
  showNarration: boolean;
  showAnimation: boolean;
  useTimings: boolean;
}
const p = (name: string) => qname('p', name, NS.pml);
const value = (node: XmlElement, name: string) => getAttrValue(node, qname('', name, ''));
const flag = (node: XmlElement | undefined, name: string, fallback: boolean) => {
  const raw = node && value(node, name);
  return raw == null ? fallback : raw === '1' || raw === 'true';
};
function propertiesPart(pres: PresentationData) {
  const pkg = pres[INTERNAL_PACKAGE];
  const rels = pkg.getRels(PRES_PART_NAME) ?? { items: [] };
  const rel = rels.items.find((item) => item.type === REL_TYPES.presProps);
  if (rel?.targetMode === 'External') throw new Error('Presentation properties must be internal.');
  const part = rel && pkg.getPart(resolveTarget(PRES_PART_NAME, rel.target));
  const doc = part && parseXml(decode(part.data));
  if (
    doc &&
    (doc.root.name.namespaceURI !== NS.pml || doc.root.name.localName !== 'presentationPr')
  )
    throw new Error('Invalid presentation properties root.');
  return { pkg, rels, rel, part, doc };
}
/** Reads presentation-wide slideshow settings using the OOXML defaults. */
export function getSlideShowProperties(pres: PresentationData): SlideShowProperties {
  const { doc } = propertiesPart(pres);
  const show = doc ? (firstChildElement(doc.root, p('showPr')) ?? undefined) : undefined;
  const browse = show && firstChildElement(show, p('browse'));
  const kiosk = show && firstChildElement(show, p('kiosk'));
  const range = show && firstChildElement(show, p('sldRg'));
  const custom = show && firstChildElement(show, p('custShow'));
  return {
    mode: browse
      ? { kind: 'browse', showScrollbar: flag(browse, 'showScrollbar', true) }
      : kiosk
        ? { kind: 'kiosk', restart: Number(value(kiosk, 'restart') ?? 300000) }
        : { kind: 'present' },
    slides: range
      ? { kind: 'range', start: Number(value(range, 'st')), end: Number(value(range, 'end')) }
      : custom
        ? { kind: 'customShow', id: Number(value(custom, 'id')) }
        : { kind: 'all' },
    loop: flag(show, 'loop', false),
    showNarration: flag(show, 'showNarration', false),
    showAnimation: flag(show, 'showAnimation', true),
    useTimings: flag(show, 'useTimings', true),
  };
}
/** Replaces slideshow settings while preserving print/web options, pen color, and extensions. */
export function setSlideShowProperties(
  pres: PresentationData,
  settings: SlideShowProperties,
): void {
  const unsigned = (n: number) => Number.isInteger(n) && n >= 0 && n <= 0xffffffff;
  if (
    !settings ||
    !settings.mode ||
    !settings.slides ||
    !['present', 'browse', 'kiosk'].includes(settings.mode.kind) ||
    !['all', 'range', 'customShow'].includes(settings.slides.kind) ||
    ['loop', 'showNarration', 'showAnimation', 'useTimings'].some(
      (key) => typeof settings[key as keyof SlideShowProperties] !== 'boolean',
    )
  )
    throw new Error('Invalid slideshow settings.');
  if (settings.mode.kind === 'browse' && typeof settings.mode.showScrollbar !== 'boolean')
    throw new Error('Invalid slideshow scrollbar setting.');
  if (settings.mode.kind === 'kiosk' && !unsigned(settings.mode.restart))
    throw new Error('Invalid kiosk restart interval.');
  const slides = settings.slides;
  if (
    slides.kind === 'range' &&
    (!unsigned(slides.start) ||
      !unsigned(slides.end) ||
      slides.start < 1 ||
      slides.start > slides.end ||
      slides.end > getSlides(pres).length)
  )
    throw new Error('Invalid slideshow slide range.');
  if (
    slides.kind === 'customShow' &&
    (!unsigned(slides.id) || !getCustomShows(pres).some((show) => show.id === slides.id))
  )
    throw new Error('Custom show destination no longer exists.');
  const { pkg, rels, rel, part, doc: existing } = propertiesPart(pres);
  const doc = existing ?? parseXml(`<p:presentationPr xmlns:p="${NS.pml}" xmlns:a="${NS.dml}"/>`);
  doc.root.prefixDecls.set('p', NS.pml);
  let show = firstChildElement(doc.root, p('showPr'));
  if (!show) {
    show = elem(p('showPr'));
    const following = doc.root.children.findIndex(
      (child) =>
        child.kind === 'element' &&
        child.name.namespaceURI === NS.pml &&
        ['clrMru', 'extLst'].includes(child.name.localName),
    );
    doc.root.children.splice(following < 0 ? doc.root.children.length : following, 0, show);
  }
  const set = (node: XmlElement, key: string, val: string) => {
    node.attrs = node.attrs.filter(
      (item) => item.name.namespaceURI !== '' || item.name.localName !== key,
    );
    node.attrs.push(attr(qname('', key, ''), val));
  };
  for (const key of ['loop', 'showNarration', 'showAnimation', 'useTimings'] as const)
    set(show, key, settings[key] ? '1' : '0');
  const mode = firstChildElement(show, p(settings.mode.kind)) ?? elem(p(settings.mode.kind));
  if (settings.mode.kind === 'browse')
    set(mode, 'showScrollbar', settings.mode.showScrollbar ? '1' : '0');
  if (settings.mode.kind === 'kiosk') set(mode, 'restart', String(settings.mode.restart));
  const selection =
    slides.kind === 'all' ? 'sldAll' : slides.kind === 'range' ? 'sldRg' : 'custShow';
  const sequence = firstChildElement(show, p(selection)) ?? elem(p(selection));
  if (slides.kind === 'range') {
    set(sequence, 'st', String(slides.start));
    set(sequence, 'end', String(slides.end));
  }
  if (slides.kind === 'customShow') set(sequence, 'id', String(slides.id));
  show.children = [
    mode,
    sequence,
    ...show.children.filter(
      (child) =>
        child.kind !== 'element' ||
        child.name.namespaceURI !== NS.pml ||
        !['present', 'browse', 'kiosk', 'sldAll', 'sldRg', 'custShow'].includes(
          child.name.localName,
        ),
    ),
  ];
  const data = encode(serializeXml(doc));
  if (part) part.data = data;
  else {
    let name = rel ? resolveTarget(PRES_PART_NAME, rel.target) : partName('/ppt/presProps.xml');
    if (!rel) for (let n = 1; pkg.getPart(name); n++) name = partName(`/ppt/presProps${n}.xml`);
    pkg.addPart(
      name,
      'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml',
      data,
    );
    if (!rel) {
      rels.items.push({
        id: nextRelId(rels.items.map((item) => item.id)),
        type: REL_TYPES.presProps,
        target: name,
        targetMode: 'Internal',
      });
      pkg.setRels(PRES_PART_NAME, rels);
    }
  }
}

import { resolveTarget } from '../../internal/opc/index.ts';
import { REL_TYPES } from '../../internal/presentationml/relationship-types.ts';
import {
  NS,
  allChildElements,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  parseXml,
  qname,
  serializeXml,
  type XmlElement,
} from '../../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  SLIDE_PART_NAME,
  type PresentationData,
  type SlideData,
} from '../_internal-symbols.ts';
import { PRES_PART_NAME, decode, encode } from './_helpers.ts';
import { getSlides } from './slide-query.ts';

/** A named slide sequence. Repeated slides and empty sequences are supported. */
export interface CustomShow {
  readonly id: number;
  readonly name: string;
  readonly slides: readonly SlideData[];
}
const p = (name: string) => qname('p', name, NS.pml);
const id = qname('', 'id', '');
const name = qname('', 'name', '');
const rid = qname('r', 'id', NS.officeDocRels);
const is = (node: XmlElement, local: string) =>
  node.name.namespaceURI === NS.pml && node.name.localName === local;

function slideRelationships(pres: PresentationData) {
  return (
    pres[INTERNAL_PACKAGE]
      .getRels(PRES_PART_NAME)
      ?.items.filter((rel) => rel.type === REL_TYPES.slide && rel.targetMode === 'Internal') ?? []
  );
}

/** Reads custom shows in list order. Missing slide references are omitted. */
export function getCustomShows(pres: PresentationData): CustomShow[] {
  const part = pres[INTERNAL_PACKAGE].getPart(PRES_PART_NAME);
  if (!part) return [];
  const list = firstChildElement(parseXml(decode(part.data)).root, p('custShowLst'));
  if (!list) return [];
  const byPart = new Map(getSlides(pres).map((slide) => [slide[SLIDE_PART_NAME], slide]));
  const byRel = new Map(
    slideRelationships(pres).map(
      (rel) => [rel.id, byPart.get(resolveTarget(PRES_PART_NAME, rel.target))] as const,
    ),
  );
  return allChildElements(list, p('custShow')).map((show) => {
    const sequence = firstChildElement(show, p('sldLst'));
    return {
      id: Number(getAttrValue(show, id)),
      name: getAttrValue(show, name) ?? '',
      slides: sequence
        ? allChildElements(sequence, p('sld')).flatMap((node) => {
            const slide = byRel.get(getAttrValue(node, rid) ?? '');
            return slide ? [slide] : [];
          })
        : [],
    };
  });
}

/**
 * Replaces the custom-show list atomically, retaining metadata on surviving IDs.
 * Slide handles must belong to this presentation. IDs are stable unsigned integers.
 */
export function setCustomShows(pres: PresentationData, shows: readonly CustomShow[]): void {
  const pkg = pres[INTERNAL_PACKAGE];
  const part = pkg.getPart(PRES_PART_NAME);
  if (!part) throw new Error('presentation.xml missing');
  const doc = parseXml(decode(part.data));
  const existing = firstChildElement(doc.root, p('custShowLst'));
  const byPart = new Map(
    slideRelationships(pres).map(
      (rel) => [resolveTarget(PRES_PART_NAME, rel.target), rel.id] as const,
    ),
  );
  const slideParts = new Set(getSlides(pres).map((slide) => slide[SLIDE_PART_NAME]));
  const ids = new Set<number>();
  // Resolve and validate everything before changing either XML or package state.
  const resolved = shows.map((show) => {
    if (!Number.isInteger(show.id) || show.id < 0 || show.id > 0xffffffff || ids.has(show.id))
      throw new Error('Custom show IDs must be unique unsigned integers');
    if (typeof show.name !== 'string') throw new Error('Custom show name must be a string');
    ids.add(show.id);
    return {
      ...show,
      relationships: show.slides.map((slide) => {
        const rel = byPart.get(slide[SLIDE_PART_NAME]);
        if (slide[INTERNAL_PACKAGE] !== pkg || !slideParts.has(slide[SLIDE_PART_NAME]) || !rel)
          throw new Error('Custom show slide must belong to this presentation');
        return rel;
      }),
    };
  });
  const oldShows = existing ? allChildElements(existing, p('custShow')) : [];
  const children = resolved.map((show) => {
    const node =
      oldShows.find((old) => getAttrValue(old, id) === String(show.id)) ?? elem(p('custShow'));
    node.attrs = node.attrs.filter(
      (a) => !(a.name.namespaceURI === '' && ['id', 'name'].includes(a.name.localName)),
    );
    node.attrs.push(attr(id, String(show.id)), attr(name, show.name));
    const oldSequence = firstChildElement(node, p('sldLst'));
    const sequence = oldSequence ?? elem(p('sldLst'));
    sequence.children = show.relationships.map((rel) =>
      elem(p('sld'), { attrs: [attr(rid, rel)] }),
    );
    node.children = [
      sequence,
      ...node.children.filter((child) => child.kind !== 'element' || !is(child, 'sldLst')),
    ];
    return node;
  });
  if (existing) existing.children = children;
  else if (children.length) {
    const list = elem(p('custShowLst'), { children });
    const following = new Set([
      'photoAlbum',
      'custDataLst',
      'kinsoku',
      'defaultTextStyle',
      'modifyVerifier',
      'extLst',
    ]);
    const index = doc.root.children.findIndex(
      (child) =>
        child.kind === 'element' &&
        child.name.namespaceURI === NS.pml &&
        following.has(child.name.localName),
    );
    doc.root.children.splice(index < 0 ? doc.root.children.length : index, 0, list);
  } else return;
  part.data = encode(serializeXml(doc));
}

/** Internal deletion hook: remove all occurrences while retaining names and IDs. */
export function removeCustomShowSlideReferences(root: XmlElement, relationshipId: string): void {
  const list = firstChildElement(root, p('custShowLst'));
  if (!list) return;
  for (const show of allChildElements(list, p('custShow'))) {
    const sequence = firstChildElement(show, p('sldLst'));
    if (sequence)
      sequence.children = sequence.children.filter(
        (child) =>
          child.kind !== 'element' ||
          !is(child, 'sld') ||
          getAttrValue(child, rid) !== relationshipId,
      );
  }
}

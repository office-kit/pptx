// Slide size.

import { oneOf, slideSizeCoordinate } from '../../internal/bounds.ts';
import { SLIDE_SIZE_TYPES } from '../../internal/enum-values.ts';
import type { Emu } from '../units.ts';
import { scaleSlideContent } from '../../internal/drawingml/scale-slide-content.ts';
import {
  NS,
  attr,
  elem,
  firstChildElement,
  parseXml,
  qname,
  serializeXml,
} from '../../internal/xml/index.ts';
import { readPresentationPart } from '../../internal/presentationml/index.ts';
import {
  INTERNAL_PACKAGE,
  SLIDE_DOCUMENT,
  SLIDE_PART_NAME,
  type PresentationData,
  type SlideData,
} from '../_internal-symbols.ts';
import { PRES_PART_NAME, decode, encode, refreshSlideData } from './_helpers.ts';

// Slide size.

/**
 * Width × height of the slide canvas, in EMU. `type` is PowerPoint's
 * aspect-ratio hint (`screen4x3`, `screen16x9`, ...); the actual size
 * is always `width` × `height`.
 */
export interface SlideSize {
  readonly width: Emu;
  readonly height: Emu;
  readonly type?: string;
}

/** Returns the slide canvas size, or `null` if `presentation.xml` omits it. */
export const getSlideSize = (pres: PresentationData): SlideSize | null => {
  const pkg = pres[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (presPart === null) return null;
  const root = parseXml(decode(presPart.data)).root;
  const model = readPresentationPart(root);
  if (model.slideSize === null) return null;
  return {
    width: model.slideSize.cx as Emu,
    height: model.slideSize.cy as Emu,
    ...(model.slideSize.type !== undefined ? { type: model.slideSize.type } : {}),
  };
};

/** Notes page canvas dimensions, or null when the presentation omits them. */
export const getNotesSize = (pres: PresentationData): Omit<SlideSize, 'type'> | null => {
  const part = pres[INTERNAL_PACKAGE].getPart(PRES_PART_NAME);
  if (!part) return null;
  const size = readPresentationPart(parseXml(decode(part.data)).root).notesSize;
  return size ? { width: size.cx as Emu, height: size.cy as Emu } : null;
};

/** First displayed slide number; omitted presentation attributes default to 1. */
export const getFirstSlideNumber = (pres: PresentationData): number => {
  const part = pres[INTERNAL_PACKAGE].getPart(PRES_PART_NAME);
  if (!part) return 1;
  const root = parseXml(decode(part.data)).root;
  const value = root.attrs.find(
    (a) => a.name.localName === 'firstSlideNum' && a.name.namespaceURI === '',
  )?.value;
  return value === undefined ? 1 : Number(value);
};

const NAME_SLD_SZ_FN = qname('p', 'sldSz', NS.pml);
const ATTR_CX = qname('', 'cx', '');
const ATTR_CY = qname('', 'cy', '');
const ATTR_TYPE = qname('', 'type', '');
const BEFORE_SLIDE_SIZE = new Set([
  'sldMasterIdLst',
  'notesMasterIdLst',
  'handoutMasterIdLst',
  'sldIdLst',
]);

/**
 * Sets the slide canvas size. Creates `<p:sldSz>` when absent, replaces
 * its attributes when present. The `type` hint is preserved as given.
 * Dimensions round to whole EMU and must be within 1–56 inches.
 * `firstSlideNumber` optionally updates the displayed starting number (0–9999)
 * in the same transaction; omitted values preserve the existing setting.
 * `notesOrientation` swaps the notes canvas dimensions independently. It does
 * not rearrange shapes on notes pages or notes/handout masters.
 * By default this changes the canvas only. With `scaleContent`, explicit
 * slide/layout/master geometry and text measurements scale proportionally
 * to fit the new canvas, centered on it. Notes and media bytes stay unchanged.
 *
 * Schema ordering: `<p:sldSz>` follows `<p:sldIdLst>` per ECMA-376
 * §19.2.1.26; we insert at the correct position when bootstrapping.
 */
export const setSlideSize = (
  pres: PresentationData,
  opts: SlideSize,
  options: {
    scaleContent?: boolean;
    firstSlideNumber?: number;
    notesOrientation?: 'portrait' | 'landscape';
  } = {},
): void => {
  if (
    options.firstSlideNumber !== undefined &&
    (!Number.isInteger(options.firstSlideNumber) ||
      options.firstSlideNumber < 0 ||
      options.firstSlideNumber > 9999)
  )
    throw new RangeError('firstSlideNumber must be an integer from 0 to 9999');
  if (options.notesOrientation !== undefined)
    oneOf(options.notesOrientation, ['portrait', 'landscape'], 'notesOrientation');
  if (options.scaleContent !== undefined && typeof options.scaleContent !== 'boolean')
    throw new TypeError('setSlideSize: scaleContent must be boolean');
  if (opts.type !== undefined) oneOf(opts.type, SLIDE_SIZE_TYPES, 'setSlideSize: type');
  const width = slideSizeCoordinate(opts.width, 'setSlideSize: width');
  const height = slideSizeCoordinate(opts.height, 'setSlideSize: height');
  const pkg = pres[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (!presPart) throw new Error('presentation.xml is missing');
  const doc = parseXml(decode(presPart.data));
  if (options.firstSlideNumber !== undefined) {
    doc.root.attrs = doc.root.attrs.filter(
      (a) => !(a.name.localName === 'firstSlideNum' && a.name.namespaceURI === ''),
    );
    doc.root.attrs.push(attr(qname('', 'firstSlideNum', ''), String(options.firstSlideNumber)));
  }
  const previous = getSlideSize(pres);
  if (options.scaleContent && !previous)
    throw new Error('Cannot scale content without an existing slide size.');

  let sldSz = firstChildElement(doc.root, NAME_SLD_SZ_FN);
  if (sldSz === null) {
    sldSz = elem(NAME_SLD_SZ_FN);
    const after = doc.root.children.findIndex(
      (node) =>
        node.kind === 'element' &&
        node.name.namespaceURI === NS.pml &&
        !BEFORE_SLIDE_SIZE.has(node.name.localName),
    );
    doc.root.children.splice(after < 0 ? doc.root.children.length : after, 0, sldSz);
  }

  sldSz.attrs = [attr(ATTR_CX, String(width)), attr(ATTR_CY, String(height))];
  if (opts.type !== undefined) sldSz.attrs.push(attr(ATTR_TYPE, opts.type));

  if (options.notesOrientation !== undefined) {
    const name = qname('p', 'notesSz', NS.pml);
    let notes = firstChildElement(doc.root, name);
    const previous = getNotesSize(pres) ?? { width: 6858000, height: 9144000 };
    if (
      !Number.isSafeInteger(previous.width) ||
      !Number.isSafeInteger(previous.height) ||
      previous.width <= 0 ||
      previous.height <= 0
    )
      throw new Error('Invalid notes page dimensions');
    if (!notes) {
      notes = elem(name);
      doc.root.children.splice(doc.root.children.indexOf(sldSz) + 1, 0, notes);
    }
    const short = Math.min(previous.width, previous.height);
    const long = Math.max(previous.width, previous.height);
    notes.attrs = notes.attrs.filter(
      (a) => !(a.name.namespaceURI === '' && ['cx', 'cy'].includes(a.name.localName)),
    );
    notes.attrs.push(
      attr(ATTR_CX, String(options.notesOrientation === 'landscape' ? long : short)),
      attr(ATTR_CY, String(options.notesOrientation === 'landscape' ? short : long)),
    );
  }

  const pending = new Map([[presPart, doc]]);
  if (options.scaleContent && previous) {
    const factor = Math.min(width / previous.width, height / previous.height);
    const dx = (width - previous.width * factor) / 2;
    const dy = (height - previous.height * factor) / 2;
    for (const part of pkg.parts) {
      if (
        /^application\/vnd\.openxmlformats-officedocument\.presentationml\.(slide|slideLayout|slideMaster)\+xml$/.test(
          part.contentType,
        )
      )
        pending.set(part, parseXml(decode(part.data)));
    }
    for (const document of pending.values()) scaleSlideContent(document.root, factor, dx, dy);
  }
  // Parse and serialize every affected part before committing any bytes.
  const writes = [...pending].map(([part, document]) => ({
    part,
    document,
    data: encode(serializeXml(document)),
  }));
  for (const { part, data, document } of writes) {
    part.data = data;
    for (const slide of (pres._slidesCache ?? []) as SlideData[]) {
      if (slide[SLIDE_PART_NAME] !== part.name) continue;
      slide[SLIDE_DOCUMENT] = document;
      refreshSlideData(slide);
    }
  }
};

import { emu as emuValue } from '../units.ts';

/** 10in × 7.5in (`screen4x3`). */
export const SLIDE_SIZE_4_3: SlideSize = {
  width: emuValue(9144000),
  height: emuValue(6858000),
  type: 'screen4x3',
};

/** 13.333in × 7.5in (`screen16x9`) — Office 2013+ default. */
export const SLIDE_SIZE_16_9: SlideSize = {
  width: emuValue(12192000),
  height: emuValue(6858000),
  type: 'screen16x9',
};

/** 13.333in × 8.33in (`screen16x10`). */
export const SLIDE_SIZE_16_10: SlideSize = {
  width: emuValue(12192000),
  height: emuValue(7620000),
  type: 'screen16x10',
};

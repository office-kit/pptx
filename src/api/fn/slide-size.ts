// Slide size.

import { boundedInt, oneOf } from '../../internal/bounds.ts';
import { SLIDE_SIZE_TYPES } from '../../internal/enum-values.ts';
import type { Emu } from '../units.ts';
import {
  NS,
  type XmlDocument,
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
} from '../_internal-symbols.ts';
import { PRES_PART_NAME, decode, encode, refreshSlideData } from './_helpers.ts';

import { getSlides } from './slide-query.ts';
import { scaleSlideContent } from './_scale-slide-content.ts';

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

const NAME_SLD_SZ_FN = qname('p', 'sldSz', NS.pml);
const ATTR_CX = qname('', 'cx', '');
const ATTR_CY = qname('', 'cy', '');
const ATTR_TYPE = qname('', 'type', '');
const NAME_SLD_ID_LST_FN = qname('p', 'sldIdLst', NS.pml);

/**
 * Sets the slide canvas size. Creates `<p:sldSz>` when absent, replaces
 * its attributes when present. Dimensions must be 1–56 inches in EMU;
 * invalid dimensions throw before changing the presentation.
 * The `type` hint is preserved as given. `content: 'fit'` uniformly scales
 * content to fit the new page and centers it, including layout/master artwork
 * and physical text formatting. The default keeps existing content unchanged.
 *
 * Schema ordering: `<p:sldSz>` follows `<p:sldIdLst>` per ECMA-376
 * §19.2.1.26; we insert at the correct position when bootstrapping.
 */
export const setSlideSize = (
  pres: PresentationData,
  opts: SlideSize,
  options: { content?: 'keep' | 'fit' } = {},
): void => {
  if (options.content !== undefined)
    oneOf(options.content, ['keep', 'fit'] as const, 'setSlideSize: content');
  const width = boundedInt(opts.width, 'slideSize', 'setSlideSize: width');
  const height = boundedInt(opts.height, 'slideSize', 'setSlideSize: height');
  if (opts.type !== undefined) oneOf(opts.type, SLIDE_SIZE_TYPES, 'setSlideSize: type');
  const pkg = pres[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (!presPart) throw new Error('presentation.xml is missing');
  const doc = parseXml(decode(presPart.data));

  // Stage every part before committing, so an invalid scaled font/coordinate
  // cannot leave half of the presentation resized.
  const updates: Array<{ part: (typeof pkg.parts)[number]; doc: XmlDocument; data: Uint8Array }> =
    [];
  if (options.content === 'fit') {
    const previous = readPresentationPart(doc.root).slideSize;
    if (!previous || previous.cx <= 0 || previous.cy <= 0)
      throw new Error('setSlideSize: the current slide size is unavailable');
    const scale = Math.min(width / previous.cx, height / previous.cy);
    const dx = (width - previous.cx * scale) / 2;
    const dy = (height - previous.cy * scale) / 2;
    if (scale !== 1 || dx !== 0 || dy !== 0) {
      scaleSlideContent(doc.root, scale, 0, 0);
      for (const part of pkg.parts) {
        if (
          !/^application\/vnd\.openxmlformats-officedocument\.(presentationml\.(slide|slideLayout|slideMaster)|drawingml\.(theme|chart))\+xml$/.test(
            part.contentType,
          )
        )
          continue;
        const changed = parseXml(decode(part.data));
        scaleSlideContent(changed.root, scale, dx, dy);
        updates.push({ part, doc: changed, data: encode(serializeXml(changed)) });
      }
    }
  }

  let sldSz = firstChildElement(doc.root, NAME_SLD_SZ_FN);
  if (sldSz === null) {
    sldSz = elem(NAME_SLD_SZ_FN);
    const sldIdLst = firstChildElement(doc.root, NAME_SLD_ID_LST_FN);
    if (sldIdLst !== null) {
      const idx = doc.root.children.indexOf(sldIdLst);
      doc.root.children.splice(idx + 1, 0, sldSz);
    } else {
      doc.root.children.push(sldSz);
    }
  }

  sldSz.attrs = [attr(ATTR_CX, String(width)), attr(ATTR_CY, String(height))];
  if (opts.type !== undefined) sldSz.attrs.push(attr(ATTR_TYPE, opts.type));

  const presentationBytes = encode(serializeXml(doc));
  const slides = updates.length ? getSlides(pres) : [];
  for (const update of updates) update.part.data = update.data;
  const byName = new Map(updates.map((update) => [update.part.name, update.doc]));
  for (const slide of slides) {
    const changed = byName.get(slide[SLIDE_PART_NAME]);
    if (changed) {
      slide[SLIDE_DOCUMENT] = changed;
      refreshSlideData(slide);
    }
  }
  presPart.data = presentationBytes;
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

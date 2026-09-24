import {
  NS,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  qname,
  type XmlElement,
} from '../../internal/xml/index.ts';
import { SHAPE_ELEMENT, SHAPE_SNAPSHOT, type SlideShapeData } from '../_internal-symbols.ts';
import { emu, type Emu } from '../units.ts';
import { commitAndRefresh } from './_helpers.ts';

export type ImageTileAlignment = 'tl' | 't' | 'tr' | 'l' | 'ctr' | 'r' | 'bl' | 'b' | 'br';
export type ImageTileFlip = 'none' | 'x' | 'y' | 'xy';

/** Image placement inside the shape; percentages use fractions (0.25 = 25%). */
export type ImageFillLayout = {
  /** Omitted on write preserves the existing rotation setting. */
  rotateWithShape?: boolean;
} & (
  | { mode: 'stretch'; left?: number; top?: number; right?: number; bottom?: number }
  | {
      mode: 'tile';
      offsetX?: Emu;
      offsetY?: Emu;
      scaleX?: number;
      scaleY?: number;
      alignment?: ImageTileAlignment;
      flip?: ImageTileFlip;
    }
);

const fillElement = (shape: SlideShapeData): XmlElement | null => {
  const element = shape[SHAPE_ELEMENT];
  if (shape[SHAPE_SNAPSHOT].kind === 'picture') {
    return firstChildElement(element, qname('p', 'blipFill', NS.pml));
  }
  const spPr = firstChildElement(element, qname('p', 'spPr', NS.pml));
  return spPr ? firstChildElement(spPr, qname('a', 'blipFill', NS.dml)) : null;
};
const value = (element: XmlElement, name: string): string | null =>
  getAttrValue(element, qname('', name, ''));
const percentage = (element: XmlElement, name: string, fallback: number): number => {
  const raw = value(element, name)?.trim();
  if (!raw) return fallback;
  const parsed = Number(raw.endsWith('%') ? raw.slice(0, -1) : raw);
  return Number.isFinite(parsed) ? parsed / (raw.endsWith('%') ? 100 : 100000) : fallback;
};
const coordinate = (element: XmlElement, name: string): Emu => {
  const raw = value(element, name)?.trim() ?? '0';
  // ST_Coordinate also permits physical units in Transitional documents.
  const units: Record<string, number> = {
    mm: 36000,
    cm: 360000,
    in: 914400,
    pt: 12700,
    pc: 152400,
    pi: 152400,
  };
  const match = /^([+-]?(?:\d+(?:\.\d*)?|\.\d+))(mm|cm|in|pt|pc|pi)?$/.exec(raw);
  return emu(match ? Number(match[1]) * (match[2] ? units[match[2]]! : 1) : 0);
};
const alignment = (raw: string | null): ImageTileAlignment => {
  switch (raw?.trim()) {
    case 't':
      return 't';
    case 'tr':
      return 'tr';
    case 'l':
      return 'l';
    case 'ctr':
      return 'ctr';
    case 'r':
      return 'r';
    case 'bl':
      return 'bl';
    case 'b':
      return 'b';
    case 'br':
      return 'br';
    default:
      return 'tl';
  }
};
const flip = (raw: string | null): ImageTileFlip => {
  switch (raw?.trim()) {
    case 'x':
      return 'x';
    case 'y':
      return 'y';
    case 'xy':
      return 'xy';
    default:
      return 'none';
  }
};

/** Reads direct picture/image-fill placement; returns null when there is no image fill. */
export const getShapeImageFillLayout = (shape: SlideShapeData): ImageFillLayout | null => {
  const fill = fillElement(shape);
  if (!fill) return null;
  const rotation = value(fill, 'rotWithShape')?.trim();
  const rotateWithShape = rotation !== '0' && rotation !== 'false';
  const tile = firstChildElement(fill, qname('a', 'tile', NS.dml));
  if (tile)
    return {
      mode: 'tile',
      rotateWithShape,
      offsetX: coordinate(tile, 'tx'),
      offsetY: coordinate(tile, 'ty'),
      scaleX: percentage(tile, 'sx', 1),
      scaleY: percentage(tile, 'sy', 1),
      alignment: alignment(value(tile, 'algn')),
      flip: flip(value(tile, 'flip')),
    };
  const stretch = firstChildElement(fill, qname('a', 'stretch', NS.dml));
  const rect = stretch && firstChildElement(stretch, qname('a', 'fillRect', NS.dml));
  return {
    mode: 'stretch',
    rotateWithShape,
    left: rect ? percentage(rect, 'l', 0) : 0,
    top: rect ? percentage(rect, 't', 0) : 0,
    right: rect ? percentage(rect, 'r', 0) : 0,
    bottom: rect ? percentage(rect, 'b', 0) : 0,
  };
};

// ECMA-376 ST_PercentageDecimal is xsd:int; ST_Coordinate has narrower bounds than xsd:long.
const MIN_PERCENTAGE = -2147483648;
const MAX_PERCENTAGE = 2147483647;
const MIN_COORDINATE = -27273042329600;
const MAX_COORDINATE = 27273042316900;
const numberAttribute = (name: string, input: number, scale: number, min: number, max: number) => {
  const stored = Math.round(input * scale);
  if (!Number.isFinite(input) || stored < min || stored > max) {
    throw new RangeError(`Image fill ${name} is outside the supported range`);
  }
  return attr(qname('', name, ''), String(stored));
};
const percentAttribute = (name: string, input: number) =>
  numberAttribute(name, input, 100000, MIN_PERCENTAGE, MAX_PERCENTAGE);

/**
 * Changes picture/image-fill placement without replacing its media, crop, or effects.
 * Replaces the tile/stretch mode; omitted layout fields use native defaults.
 * Invalid inputs leave the shape unchanged.
 */
export const setShapeImageFillLayout = (shape: SlideShapeData, layout: ImageFillLayout): void => {
  const fill = fillElement(shape);
  if (!fill)
    throw new Error('setShapeImageFillLayout requires a picture or a shape with an image fill');
  let mode: XmlElement;
  if (layout.mode === 'tile') {
    const tileAlignment = layout.alignment ?? 'tl';
    const tileFlip = layout.flip ?? 'none';
    if (alignment(tileAlignment) !== tileAlignment || flip(tileFlip) !== tileFlip) {
      throw new RangeError('Invalid image tile alignment or flip');
    }
    mode = elem(qname('a', 'tile', NS.dml), {
      attrs: [
        numberAttribute('tx', layout.offsetX ?? 0, 1, MIN_COORDINATE, MAX_COORDINATE),
        numberAttribute('ty', layout.offsetY ?? 0, 1, MIN_COORDINATE, MAX_COORDINATE),
        percentAttribute('sx', layout.scaleX ?? 1),
        percentAttribute('sy', layout.scaleY ?? 1),
        attr(qname('', 'flip', ''), tileFlip),
        attr(qname('', 'algn', ''), tileAlignment),
      ],
    });
  } else if (layout.mode === 'stretch') {
    mode = elem(qname('a', 'stretch', NS.dml), {
      children: [
        elem(qname('a', 'fillRect', NS.dml), {
          attrs: [
            percentAttribute('l', layout.left ?? 0),
            percentAttribute('t', layout.top ?? 0),
            percentAttribute('r', layout.right ?? 0),
            percentAttribute('b', layout.bottom ?? 0),
          ],
        }),
      ],
    });
  } else {
    throw new RangeError('Invalid image fill mode');
  }
  if (layout.rotateWithShape !== undefined && typeof layout.rotateWithShape !== 'boolean') {
    throw new TypeError('rotateWithShape must be a boolean');
  }
  fill.children = fill.children.filter(
    (child) =>
      !(
        child.kind === 'element' &&
        child.name.namespaceURI === NS.dml &&
        (child.name.localName === 'tile' || child.name.localName === 'stretch')
      ),
  );
  fill.children.push(mode);
  if (layout.rotateWithShape !== undefined) {
    fill.attrs = fill.attrs.filter(
      (attribute) =>
        !(attribute.name.localName === 'rotWithShape' && attribute.name.namespaceURI === ''),
    );
    fill.attrs.push(attr(qname('', 'rotWithShape', ''), layout.rotateWithShape ? '1' : '0'));
  }
  commitAndRefresh(shape);
};

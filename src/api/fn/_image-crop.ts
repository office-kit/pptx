import {
  firstChildElement,
  getAttrValue,
  qname,
  NS,
  type XmlElement,
} from '../../internal/xml/index.ts';
import type { ImageCrop } from './shape-image-effects.ts';

export function readImageCrop(blipFill: XmlElement): ImageCrop | null {
  const srcRect = firstChildElement(blipFill, qname('a', 'srcRect', NS.dml));
  if (!srcRect) return null;
  const parseSide = (local: string): number => {
    const v = getAttrValue(srcRect, qname('', local, ''));
    if (v === null) return 0;
    const n = Number.parseInt(v, 10);
    return Number.isFinite(n) ? n / 100000 : 0;
  };
  return {
    left: parseSide('l'),
    top: parseSide('t'),
    right: parseSide('r'),
    bottom: parseSide('b'),
  };
}

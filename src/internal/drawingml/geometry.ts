// Read-only geometry extraction from DrawingML transforms.
//
// `<a:xfrm>` carries the position (`<a:off x cy/>`) and size
// (`<a:ext cx cy/>`) of a shape. PresentationML wraps the transform
// differently depending on the shape kind:
//
//   - `p:sp`, `p:pic`, `p:cxnSp` → `<p:spPr><a:xfrm>...</a:xfrm></p:spPr>`
//   - `p:grpSp` → `<p:grpSpPr><a:xfrm>...</a:xfrm></p:grpSpPr>`
//   - `p:graphicFrame` → `<p:xfrm>...</p:xfrm>` (note: p:, not a:)
//
// Each can be absent. Placeholders inheriting position from the layout
// frequently omit `<a:xfrm>` on the slide; effective geometry would then
// resolve up the inheritance chain. That resolution lives in the
// presentationml/ layer once the layout reader is in place.

import { NS, firstChildElement, getAttrValue, qname } from '../xml/index.ts';
import type { XmlElement } from '../xml/index.ts';

const NAME_SP_PR = qname('p', 'spPr', NS.pml);
const NAME_GRP_SP_PR = qname('p', 'grpSpPr', NS.pml);
const NAME_A_XFRM = qname('a', 'xfrm', NS.dml);
const NAME_P_XFRM = qname('p', 'xfrm', NS.pml);
const NAME_OFF = qname('a', 'off', NS.dml);
const NAME_EXT = qname('a', 'ext', NS.dml);
const ATTR_X = qname('', 'x', '');
const ATTR_Y = qname('', 'y', '');
const ATTR_CX = qname('', 'cx', '');
const ATTR_CY = qname('', 'cy', '');

export interface Position {
  /** Left edge offset from the slide origin, in EMU. */
  readonly x: number;
  /** Top edge offset from the slide origin, in EMU. */
  readonly y: number;
}

export interface Size {
  /** Width in EMU. */
  readonly w: number;
  /** Height in EMU. */
  readonly h: number;
}

const parseIntOr = (raw: string | null): number | null =>
  raw === null ? null : Number.isFinite(Number.parseInt(raw, 10)) ? Number.parseInt(raw, 10) : null;

const findTransform = (shape: XmlElement, kind: ShapeKindForGeometry): XmlElement | null => {
  switch (kind) {
    case 'shape':
    case 'picture':
    case 'connector': {
      const spPr = firstChildElement(shape, NAME_SP_PR);
      if (spPr === null) return null;
      return firstChildElement(spPr, NAME_A_XFRM);
    }
    case 'group': {
      const grpSpPr = firstChildElement(shape, NAME_GRP_SP_PR);
      if (grpSpPr === null) return null;
      return firstChildElement(grpSpPr, NAME_A_XFRM);
    }
    case 'graphicFrame':
      // Graphic frames use a PML xfrm wrapper (no spPr).
      return firstChildElement(shape, NAME_P_XFRM);
  }
};

/**
 * Subset of `ShapeKind` from presentationml/, redeclared here to avoid an
 * upward import. They're kept in lock-step intentionally.
 */
export type ShapeKindForGeometry = 'shape' | 'picture' | 'group' | 'graphicFrame' | 'connector';

/** Returns the shape's position (in EMU) or `null` if `<a:off>` is absent. */
export const readPosition = (shape: XmlElement, kind: ShapeKindForGeometry): Position | null => {
  const xfrm = findTransform(shape, kind);
  if (xfrm === null) return null;
  const off = firstChildElement(xfrm, NAME_OFF);
  if (off === null) return null;
  const x = parseIntOr(getAttrValue(off, ATTR_X));
  const y = parseIntOr(getAttrValue(off, ATTR_Y));
  if (x === null || y === null) return null;
  return { x, y };
};

/** Returns the shape's size (in EMU) or `null` if `<a:ext>` is absent. */
export const readSize = (shape: XmlElement, kind: ShapeKindForGeometry): Size | null => {
  const xfrm = findTransform(shape, kind);
  if (xfrm === null) return null;
  const ext = firstChildElement(xfrm, NAME_EXT);
  if (ext === null) return null;
  const w = parseIntOr(getAttrValue(ext, ATTR_CX));
  const h = parseIntOr(getAttrValue(ext, ATTR_CY));
  if (w === null || h === null) return null;
  return { w, h };
};

const ATTR_ROT_R = qname('', 'rot', '');
const ATTR_FLIP_H_R = qname('', 'flipH', '');
const ATTR_FLIP_V_R = qname('', 'flipV', '');

/**
 * Returns the shape's rotation in degrees, or `0` if no rotation is set.
 * ECMA-376 stores rotation in 60000ths of a degree; we divide.
 */
export const readRotation = (shape: XmlElement, kind: ShapeKindForGeometry): number => {
  const xfrm = findTransform(shape, kind);
  if (xfrm === null) return 0;
  const raw = parseIntOr(getAttrValue(xfrm, ATTR_ROT_R));
  if (raw === null) return 0;
  return raw / 60000;
};

/** Returns the shape's flip state (h/v boolean) or `null` when no xfrm exists. */
export const readFlip = (
  shape: XmlElement,
  kind: ShapeKindForGeometry,
): { horizontal: boolean; vertical: boolean } | null => {
  const xfrm = findTransform(shape, kind);
  if (xfrm === null) return null;
  return {
    horizontal: getAttrValue(xfrm, ATTR_FLIP_H_R) === '1',
    vertical: getAttrValue(xfrm, ATTR_FLIP_V_R) === '1',
  };
};

// Mutating counterpart to `geometry.ts`. Writes `<a:off>` / `<a:ext>` on
// the shape's transform, creating the `<a:xfrm>` (and the host element)
// when they don't yet exist.
//
// The shape-kind dispatch matches `geometry.ts` exactly — both modules
// agree on where each shape kind keeps its transform.

import { NS, attr, elem, firstChildElement, qname } from '../xml/index.ts';
import type { XmlElement } from '../xml/index.ts';
import type { ShapeKindForGeometry } from './geometry.ts';

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

/**
 * Ensures the shape has an `<a:xfrm>` (or `<p:xfrm>` for graphic frames)
 * to write into, creating the host element and its parent if absent.
 * Returns the xfrm element.
 */
const ensureTransform = (shape: XmlElement, kind: ShapeKindForGeometry): XmlElement => {
  if (kind === 'graphicFrame') {
    let xfrm = firstChildElement(shape, NAME_P_XFRM);
    if (xfrm === null) {
      xfrm = elem(NAME_P_XFRM);
      // graphicFrame children sequence: nvGraphicFramePr → xfrm → graphic.
      // Insert after nvGraphicFramePr if present, else at the front.
      const idx = shape.children.findIndex(
        (c) =>
          c.kind === 'element' &&
          c.name.namespaceURI === NS.pml &&
          c.name.localName === 'nvGraphicFramePr',
      );
      shape.children.splice(idx + 1, 0, xfrm);
    }
    return xfrm;
  }
  const hostName = kind === 'group' ? NAME_GRP_SP_PR : NAME_SP_PR;
  let host = firstChildElement(shape, hostName);
  if (host === null) {
    host = elem(hostName);
    // The host element comes after nvSpPr / nvPicPr / nvGrpSpPr / nvCxnSpPr.
    // Easiest: append; PowerPoint tolerates ordering at the spPr level.
    shape.children.push(host);
  }
  let xfrm = firstChildElement(host, NAME_A_XFRM);
  if (xfrm === null) {
    xfrm = elem(NAME_A_XFRM);
    host.children.unshift(xfrm);
  }
  return xfrm;
};

/** Sets the shape's `<a:off>` to `(x, y)` in EMU. */
export const setPosition = (
  shape: XmlElement,
  kind: ShapeKindForGeometry,
  x: number,
  y: number,
): void => {
  const xfrm = ensureTransform(shape, kind);
  let off = firstChildElement(xfrm, NAME_OFF);
  if (off === null) {
    off = elem(NAME_OFF);
    xfrm.children.unshift(off);
  }
  off.attrs = [attr(ATTR_X, String(x)), attr(ATTR_Y, String(y))];
};

/** Sets the shape's `<a:ext>` to `(w, h)` in EMU. */
export const setSize = (
  shape: XmlElement,
  kind: ShapeKindForGeometry,
  w: number,
  h: number,
): void => {
  const xfrm = ensureTransform(shape, kind);
  let ext = firstChildElement(xfrm, NAME_EXT);
  if (ext === null) {
    ext = elem(NAME_EXT);
    // ext should come AFTER off.
    const offIdx = xfrm.children.findIndex(
      (c) => c.kind === 'element' && c.name.localName === 'off' && c.name.namespaceURI === NS.dml,
    );
    if (offIdx >= 0) xfrm.children.splice(offIdx + 1, 0, ext);
    else xfrm.children.push(ext);
  }
  ext.attrs = [attr(ATTR_CX, String(w)), attr(ATTR_CY, String(h))];
};

const ATTR_ROT = qname('', 'rot', '');
const ATTR_FLIP_H = qname('', 'flipH', '');
const ATTR_FLIP_V = qname('', 'flipV', '');

/**
 * Sets the shape's rotation. `degrees` is degrees (positive clockwise),
 * fractional values allowed. PowerPoint serializes the value in
 * 60000ths of a degree per ECMA-376 ST_Angle, which we mirror.
 *
 * The full range is `0..360`; values outside that range are normalized.
 */
export const setRotation = (
  shape: XmlElement,
  kind: ShapeKindForGeometry,
  degrees: number,
): void => {
  const xfrm = ensureTransform(shape, kind);
  // Normalize into `[0, 360)`.
  const normalized = ((degrees % 360) + 360) % 360;
  const value = Math.round(normalized * 60000);
  xfrm.attrs = xfrm.attrs.filter((a) => a.name.localName !== 'rot');
  if (value !== 0) xfrm.attrs.push(attr(ATTR_ROT, String(value)));
};

/** Sets `flipH` / `flipV` boolean attributes on the shape's transform. */
export const setFlip = (
  shape: XmlElement,
  kind: ShapeKindForGeometry,
  options: { horizontal?: boolean; vertical?: boolean },
): void => {
  const xfrm = ensureTransform(shape, kind);
  if (options.horizontal !== undefined) {
    xfrm.attrs = xfrm.attrs.filter((a) => a.name.localName !== 'flipH');
    if (options.horizontal) xfrm.attrs.push(attr(ATTR_FLIP_H, '1'));
  }
  if (options.vertical !== undefined) {
    xfrm.attrs = xfrm.attrs.filter((a) => a.name.localName !== 'flipV');
    if (options.vertical) xfrm.attrs.push(attr(ATTR_FLIP_V, '1'));
  }
};

// Builds a group shape (`<p:grpSp>`) that wraps already-built shape-tree
// children (`<p:sp>` / `<p:pic>` / `<p:cxnSp>` / `<p:graphicFrame>` /
// nested `<p:grpSp>`) under a single transform.
//
// `<a:chOff>` / `<a:chExt>` define the coordinate space the children's own
// `<a:xfrm>` values live in. At creation time the children already carry
// real slide-space coordinates, so the child space is set 1:1 with the
// outer `<a:off>` / `<a:ext>` (`chOff == off`, `chExt == ext`) — the same
// convention PowerPoint itself uses right after a fresh "Group" action,
// before the group is subsequently moved or resized.

import { emuCoordinate, emuExtent } from '../bounds.ts';
import { type XmlElement, NS, attr, elem, qname } from '../xml/index.ts';

const NAME_GRP_SP = qname('p', 'grpSp', NS.pml);
const NAME_NV_GRP_SP_PR = qname('p', 'nvGrpSpPr', NS.pml);
const NAME_C_NV_PR = qname('p', 'cNvPr', NS.pml);
const NAME_C_NV_GRP_SP_PR = qname('p', 'cNvGrpSpPr', NS.pml);
const NAME_NV_PR = qname('p', 'nvPr', NS.pml);
const NAME_GRP_SP_PR = qname('p', 'grpSpPr', NS.pml);
const NAME_A_XFRM = qname('a', 'xfrm', NS.dml);
const NAME_OFF = qname('a', 'off', NS.dml);
const NAME_EXT = qname('a', 'ext', NS.dml);
const NAME_CH_OFF = qname('a', 'chOff', NS.dml);
const NAME_CH_EXT = qname('a', 'chExt', NS.dml);
const ATTR_ID = qname('', 'id', '');
const ATTR_NAME = qname('', 'name', '');
const ATTR_X = qname('', 'x', '');
const ATTR_Y = qname('', 'y', '');
const ATTR_CX = qname('', 'cx', '');
const ATTR_CY = qname('', 'cy', '');

export interface GroupOptions {
  id: number;
  name?: string;
  /** Slide-space bounds (EMU) — the union of the grouped children's bounds. */
  x: number;
  y: number;
  w: number;
  h: number;
  /** Already-built shape-tree elements to nest inside the group. */
  children: ReadonlyArray<XmlElement>;
}

/** Returns a `<p:grpSp>` wrapping `opts.children` under one transform. */
export const buildGroup = (opts: GroupOptions): XmlElement => {
  const name = opts.name ?? `Group ${opts.id}`;

  const cNvPr = elem(NAME_C_NV_PR, {
    attrs: [attr(ATTR_ID, String(opts.id)), attr(ATTR_NAME, name)],
  });
  const nvGrpSpPr = elem(NAME_NV_GRP_SP_PR, {
    children: [cNvPr, elem(NAME_C_NV_GRP_SP_PR), elem(NAME_NV_PR)],
  });

  const x = emuCoordinate(opts.x, 'groupShapes: x');
  const y = emuCoordinate(opts.y, 'groupShapes: y');
  const cx = emuExtent(opts.w, 'groupShapes: w');
  const cy = emuExtent(opts.h, 'groupShapes: h');

  const xfrm = elem(NAME_A_XFRM, {
    children: [
      elem(NAME_OFF, { attrs: [attr(ATTR_X, String(x)), attr(ATTR_Y, String(y))] }),
      elem(NAME_EXT, { attrs: [attr(ATTR_CX, String(cx)), attr(ATTR_CY, String(cy))] }),
      elem(NAME_CH_OFF, { attrs: [attr(ATTR_X, String(x)), attr(ATTR_Y, String(y))] }),
      elem(NAME_CH_EXT, { attrs: [attr(ATTR_CX, String(cx)), attr(ATTR_CY, String(cy))] }),
    ],
  });
  const grpSpPr = elem(NAME_GRP_SP_PR, { children: [xfrm] });

  return elem(NAME_GRP_SP, { children: [nvGrpSpPr, grpSpPr, ...opts.children] });
};

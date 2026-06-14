// Builds the XML element for a picture shape (`<p:pic>`).
//
// PowerPoint emits pictures with this skeleton:
//
//   <p:pic>
//     <p:nvPicPr>
//       <p:cNvPr id="X" name="Picture X"/>
//       <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
//       <p:nvPr/>
//     </p:nvPicPr>
//     <p:blipFill>
//       <a:blip r:embed="rIdN"/>
//       <a:stretch><a:fillRect/></a:stretch>
//     </p:blipFill>
//     <p:spPr>
//       <a:xfrm><a:off x y/><a:ext cx cy/></a:xfrm>
//       <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
//     </p:spPr>
//   </p:pic>
//
// The `r:embed` rId is allocated by the caller and must already exist as
// a relationship on the slide's `.rels` part pointing at a media part.

import { type XmlElement, NS, attr, elem, qname } from '../xml/index.ts';

const NAME_PIC = qname('p', 'pic', NS.pml);
const NAME_NV_PIC_PR = qname('p', 'nvPicPr', NS.pml);
const NAME_C_NV_PR = qname('p', 'cNvPr', NS.pml);
const NAME_C_NV_PIC_PR = qname('p', 'cNvPicPr', NS.pml);
const NAME_NV_PR = qname('p', 'nvPr', NS.pml);
const NAME_BLIP_FILL = qname('p', 'blipFill', NS.pml);
const NAME_BLIP = qname('a', 'blip', NS.dml);
const NAME_STRETCH = qname('a', 'stretch', NS.dml);
const NAME_FILL_RECT = qname('a', 'fillRect', NS.dml);
const NAME_SP_PR = qname('p', 'spPr', NS.pml);
const NAME_A_XFRM = qname('a', 'xfrm', NS.dml);
const NAME_OFF = qname('a', 'off', NS.dml);
const NAME_EXT = qname('a', 'ext', NS.dml);
const NAME_PRST_GEOM = qname('a', 'prstGeom', NS.dml);
const NAME_AV_LST = qname('a', 'avLst', NS.dml);
const NAME_PIC_LOCKS = qname('a', 'picLocks', NS.dml);
const ATTR_ID = qname('', 'id', '');
const ATTR_NAME = qname('', 'name', '');
const ATTR_X = qname('', 'x', '');
const ATTR_Y = qname('', 'y', '');
const ATTR_CX = qname('', 'cx', '');
const ATTR_CY = qname('', 'cy', '');
const ATTR_PRST = qname('', 'prst', '');
const ATTR_R_EMBED = qname('r', 'embed', NS.officeDocRels);
const ATTR_NO_CHANGE_ASPECT = qname('', 'noChangeAspect', '');

export interface PictureOptions {
  id: number;
  name?: string;
  rEmbed: string;
  x: number;
  y: number;
  w: number;
  h: number;
  /** When true, emits `<a:picLocks noChangeAspect="1"/>`. PowerPoint sets
   * this for pictures inserted via "Insert > Picture" but not for ones
   * dropped onto the canvas; default true to mirror the common case. */
  lockAspect?: boolean;
}

/** Returns a `<p:pic>` element ready to be appended to a slide's `<p:spTree>`. */
export const buildPicture = (opts: PictureOptions): XmlElement => {
  const name = opts.name ?? `Picture ${opts.id}`;
  const lockAspect = opts.lockAspect ?? true;

  const picLocks = elem(NAME_PIC_LOCKS, {
    attrs: lockAspect ? [attr(ATTR_NO_CHANGE_ASPECT, '1')] : [],
  });
  const cNvPicPr = elem(NAME_C_NV_PIC_PR, { children: [picLocks] });
  const cNvPr = elem(NAME_C_NV_PR, {
    attrs: [attr(ATTR_ID, String(opts.id)), attr(ATTR_NAME, name)],
  });
  const nvPr = elem(NAME_NV_PR);
  const nvPicPr = elem(NAME_NV_PIC_PR, { children: [cNvPr, cNvPicPr, nvPr] });

  const blip = elem(NAME_BLIP, { attrs: [attr(ATTR_R_EMBED, opts.rEmbed)] });
  const stretch = elem(NAME_STRETCH, { children: [elem(NAME_FILL_RECT)] });
  const blipFill = elem(NAME_BLIP_FILL, { children: [blip, stretch] });

  // Round to whole EMU — `fit: 'contain'` scaling produces fractional offsets
  // (`as Emu` cast), and fractional ST_Coordinate values corrupt the file.
  const off = elem(NAME_OFF, {
    attrs: [attr(ATTR_X, String(Math.round(opts.x))), attr(ATTR_Y, String(Math.round(opts.y)))],
  });
  const ext = elem(NAME_EXT, {
    attrs: [attr(ATTR_CX, String(Math.round(opts.w))), attr(ATTR_CY, String(Math.round(opts.h)))],
  });
  const xfrm = elem(NAME_A_XFRM, { children: [off, ext] });
  const prstGeom = elem(NAME_PRST_GEOM, {
    attrs: [attr(ATTR_PRST, 'rect')],
    children: [elem(NAME_AV_LST)],
  });
  const spPr = elem(NAME_SP_PR, { children: [xfrm, prstGeom] });

  return elem(NAME_PIC, { children: [nvPicPr, blipFill, spPr] });
};

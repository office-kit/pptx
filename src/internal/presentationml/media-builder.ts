// Builds and reads the media flavour of a picture shape (`<p:pic>`).
//
// PowerPoint stores an inserted video / audio clip as an ordinary picture (the
// poster frame) whose non-visual properties point at the media:
//
//   <p:pic>
//     <p:nvPicPr>
//       <p:cNvPr id="X" name="...">
//         <a:hlinkClick r:id="" action="ppaction://media"/>
//       </p:cNvPr>
//       <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
//       <p:nvPr>
//         <a:videoFile r:link="rIdV"/>            (or <a:audioFile>)
//         <p:extLst>
//           <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">
//             <p14:media r:embed="rIdM"/>
//           </p:ext>
//         </p:extLst>
//       </p:nvPr>
//     </p:nvPicPr>
//     <p:blipFill><a:blip r:embed="rIdPoster"/>...</p:blipFill>
//     <p:spPr>...</p:spPr>
//   </p:pic>
//
// `r:link` is the ECMA-376 (2006) reference and `p14:media r:embed` the
// PowerPoint 2010+ one; both relationships target the same media part. An
// online video has only the `r:link`, pointing at an external URL.
//
// The slide also needs a media time node under `<p:timing>`:
//
//   <p:video>                                      (or <p:audio>)
//     <p:cMediaNode vol="80000">
//       <p:cTn id="N" fill="hold" display="0">
//         <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
//       </p:cTn>
//       <p:tgtEl><p:spTgt spid="X"/></p:tgtEl>
//     </p:cMediaNode>
//   </p:video>
//
// Without that node PowerPoint shows the poster but no play button / seek bar
// in the slide show, so the clip cannot be controlled.

import {
  NS,
  type XmlElement,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  qname,
} from '../xml/index.ts';
import { type PictureOptions, buildPicture } from './picture-builder.ts';

/** Which media time node / DrawingML file element a clip uses. */
export type MediaFileKind = 'video' | 'audio';

const NAME_NV_PIC_PR = qname('p', 'nvPicPr', NS.pml);
const NAME_C_NV_PR = qname('p', 'cNvPr', NS.pml);
const NAME_NV_PR = qname('p', 'nvPr', NS.pml);
const NAME_HLINK_CLICK = qname('a', 'hlinkClick', NS.dml);
const NAME_VIDEO_FILE = qname('a', 'videoFile', NS.dml);
const NAME_AUDIO_FILE = qname('a', 'audioFile', NS.dml);
const NAME_EXT_LST = qname('p', 'extLst', NS.pml);
const NAME_EXT = qname('p', 'ext', NS.pml);
const NAME_P14_MEDIA = qname('p14', 'media', NS.p14);
const NAME_TIMING = qname('p', 'timing', NS.pml);
const NAME_TN_LST = qname('p', 'tnLst', NS.pml);
const NAME_PAR = qname('p', 'par', NS.pml);
const NAME_C_TN = qname('p', 'cTn', NS.pml);
const NAME_CHILD_TN_LST = qname('p', 'childTnLst', NS.pml);
const NAME_VIDEO = qname('p', 'video', NS.pml);
const NAME_AUDIO = qname('p', 'audio', NS.pml);
const NAME_C_MEDIA_NODE = qname('p', 'cMediaNode', NS.pml);
const NAME_ST_COND_LST = qname('p', 'stCondLst', NS.pml);
const NAME_COND = qname('p', 'cond', NS.pml);
const NAME_TGT_EL = qname('p', 'tgtEl', NS.pml);
const NAME_SP_TGT = qname('p', 'spTgt', NS.pml);

const ATTR_R_ID = qname('r', 'id', NS.officeDocRels);
const ATTR_R_LINK = qname('r', 'link', NS.officeDocRels);
const ATTR_R_EMBED = qname('r', 'embed', NS.officeDocRels);
const ATTR_ACTION = qname('', 'action', '');
const ATTR_URI = qname('', 'uri', '');
const ATTR_ID = qname('', 'id', '');
const ATTR_DUR = qname('', 'dur', '');
const ATTR_RESTART = qname('', 'restart', '');
const ATTR_NODE_TYPE = qname('', 'nodeType', '');
const ATTR_VOL = qname('', 'vol', '');
const ATTR_FILL = qname('', 'fill', '');
const ATTR_DISPLAY = qname('', 'display', '');
const ATTR_DELAY = qname('', 'delay', '');
const ATTR_SPID = qname('', 'spid', '');

const MEDIA_ACTION = 'ppaction://media';
// The `<p:ext>` URI PowerPoint 2010+ uses for its `<p14:media>` reference.
const P14_MEDIA_EXT_URI = '{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}';
// PowerPoint's default clip volume (ST_PositiveFixedPercentage, 80%).
const DEFAULT_MEDIA_VOLUME = '80000';

export interface MediaPictureOptions extends PictureOptions {
  kind: MediaFileKind;
  /** rId of the `video` / `audio` relationship (internal part or external URL). */
  rLink: string;
  /** rId of the `media` relationship to the embedded part. Omit for linked media. */
  rMedia?: string;
}

/** Returns a media `<p:pic>`; `opts.rEmbed` is the poster image relationship. */
export const buildMediaPicture = (opts: MediaPictureOptions): XmlElement => {
  const pic = buildPicture(opts);
  // buildPicture always emits nvPicPr / cNvPr / nvPr.
  const nvPicPr = firstChildElement(pic, NAME_NV_PIC_PR)!;
  const cNvPr = firstChildElement(nvPicPr, NAME_C_NV_PR)!;
  const nvPr = firstChildElement(nvPicPr, NAME_NV_PR)!;

  // The empty r:id is intentional: `ppaction://media` carries no target, and
  // PowerPoint itself writes `r:id=""` here.
  cNvPr.children.push(
    elem(NAME_HLINK_CLICK, { attrs: [attr(ATTR_R_ID, ''), attr(ATTR_ACTION, MEDIA_ACTION)] }),
  );

  nvPr.children.push(
    elem(opts.kind === 'video' ? NAME_VIDEO_FILE : NAME_AUDIO_FILE, {
      attrs: [attr(ATTR_R_LINK, opts.rLink)],
    }),
  );
  if (opts.rMedia !== undefined) {
    // Declared locally, as PowerPoint does: the slide root only binds a / r / p.
    const media = elem(NAME_P14_MEDIA, {
      attrs: [attr(ATTR_R_EMBED, opts.rMedia)],
      prefixDecls: new Map([['p14', NS.p14]]),
    });
    const ext = elem(NAME_EXT, { attrs: [attr(ATTR_URI, P14_MEDIA_EXT_URI)], children: [media] });
    nvPr.children.push(elem(NAME_EXT_LST, { children: [ext] }));
  }
  return pic;
};

/** Returns the `<p:video>` / `<p:audio>` time node that targets shape `spid`. */
export const buildMediaTimingNode = (
  kind: MediaFileKind,
  spid: number,
  cTnId: number,
): XmlElement => {
  const cTn = elem(NAME_C_TN, {
    attrs: [attr(ATTR_ID, String(cTnId)), attr(ATTR_FILL, 'hold'), attr(ATTR_DISPLAY, '0')],
    children: [
      elem(NAME_ST_COND_LST, {
        children: [elem(NAME_COND, { attrs: [attr(ATTR_DELAY, 'indefinite')] })],
      }),
    ],
  });
  const tgtEl = elem(NAME_TGT_EL, {
    children: [elem(NAME_SP_TGT, { attrs: [attr(ATTR_SPID, String(spid))] })],
  });
  const cMediaNode = elem(NAME_C_MEDIA_NODE, {
    attrs: [attr(ATTR_VOL, DEFAULT_MEDIA_VOLUME)],
    children: [cTn, tgtEl],
  });
  return elem(kind === 'video' ? NAME_VIDEO : NAME_AUDIO, { children: [cMediaNode] });
};

/**
 * Returns a `<p:timing>` holding only the root time node, with `nodes` as its
 * children. `nodes` must be non-empty (`<p:childTnLst>` requires a child).
 */
export const buildTimingRoot = (nodes: ReadonlyArray<XmlElement>): XmlElement => {
  const cTn = elem(NAME_C_TN, {
    attrs: [
      attr(ATTR_ID, '1'),
      attr(ATTR_DUR, 'indefinite'),
      attr(ATTR_RESTART, 'never'),
      attr(ATTR_NODE_TYPE, 'tmRoot'),
    ],
    children: [elem(NAME_CHILD_TN_LST, { children: [...nodes] })],
  });
  return elem(NAME_TIMING, {
    children: [elem(NAME_TN_LST, { children: [elem(NAME_PAR, { children: [cTn] })] })],
  });
};

/** `true` for a `<p:video>` / `<p:audio>` time node. */
export const isMediaTimingNode = (el: XmlElement): boolean =>
  el.name.namespaceURI === NS.pml &&
  (el.name.localName === 'video' || el.name.localName === 'audio');

/** Shape id a media time node targets, or `null` when it has no `<p:spTgt>`. */
export const mediaTimingNodeTarget = (node: XmlElement): number | null => {
  const cMediaNode = firstChildElement(node, NAME_C_MEDIA_NODE);
  const tgtEl = cMediaNode ? firstChildElement(cMediaNode, NAME_TGT_EL) : null;
  const spTgt = tgtEl ? firstChildElement(tgtEl, NAME_SP_TGT) : null;
  const raw = spTgt ? getAttrValue(spTgt, ATTR_SPID) : null;
  if (raw === null) return null;
  const spid = Number.parseInt(raw, 10);
  return Number.isFinite(spid) ? spid : null;
};

/** The relationship ids a media picture references. */
export interface PictureMediaRef {
  readonly kind: MediaFileKind;
  /** `r:link` (or `r:embed` for `<a:wavAudioFile>`) on the DrawingML file element. */
  readonly fileRId: string | null;
  /** `r:embed` on `<p14:media>`. */
  readonly mediaRId: string | null;
}

// Every EG_Media member that carries a relationship (`<a:audioCd>` does not).
// `<a:wavAudioFile>` is the only one that uses `r:embed`.
const MEDIA_FILE_ELEMENTS: Readonly<Record<string, MediaFileKind>> = {
  videoFile: 'video',
  quickTimeFile: 'video',
  audioFile: 'audio',
  wavAudioFile: 'audio',
};

/** Reads the media reference off a `<p:pic>`, or `null` for a plain picture. */
export const readPictureMediaRef = (pic: XmlElement): PictureMediaRef | null => {
  const nvPicPr = firstChildElement(pic, NAME_NV_PIC_PR);
  const nvPr = nvPicPr ? firstChildElement(nvPicPr, NAME_NV_PR) : null;
  if (!nvPr) return null;
  for (const child of nvPr.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.dml) continue;
    const kind = MEDIA_FILE_ELEMENTS[child.name.localName];
    if (kind === undefined) continue;
    const fileRId =
      child.name.localName === 'wavAudioFile'
        ? getAttrValue(child, ATTR_R_EMBED)
        : getAttrValue(child, ATTR_R_LINK);
    return { kind, fileRId, mediaRId: readP14MediaRId(nvPr) };
  }
  return null;
};

const readP14MediaRId = (nvPr: XmlElement): string | null => {
  const extLst = firstChildElement(nvPr, NAME_EXT_LST);
  if (!extLst) return null;
  for (const ext of extLst.children) {
    if (ext.kind !== 'element') continue;
    const media = firstChildElement(ext, NAME_P14_MEDIA);
    if (media) return getAttrValue(media, ATTR_R_EMBED);
  }
  return null;
};

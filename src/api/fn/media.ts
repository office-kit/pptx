// Video / audio / online-video authoring and read-back.

import type { Emu } from '../units.ts';
import { emuCoordinate, emuExtent } from '../../internal/bounds.ts';
import {
  type AudioFormat,
  type ImageFormat,
  type PartName,
  type TargetMode,
  type VideoFormat,
  contentTypeForAudioFormat,
  contentTypeForFormat,
  contentTypeForVideoFormat,
  detectAudioFormat,
  detectImageFormat,
  detectVideoFormat,
  emptyRels,
  extensionForFormat,
  nextRelId,
  partName,
  resolveTarget,
} from '../../internal/opc/index.ts';
import type { OpcPackage } from '../../internal/parts/index.ts';
import {
  REL_TYPES,
  buildMediaPicture,
  readPictureMediaRef,
} from '../../internal/presentationml/index.ts';
import {
  INTERNAL_PACKAGE,
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  SHAPE_SNAPSHOT,
  SLIDE_DOCUMENT,
  SLIDE_PART_NAME,
  SLIDE_SHAPES,
  type SlideData,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import {
  appendAndReturnNewShape,
  commitSlideData,
  nextShapeId,
  refreshSlideData,
  setOpcDefault,
} from './_helpers.ts';
import {
  addMediaTimingNode,
  findMediaTimingNodeWithAncestors,
  type MediaTimingPath,
} from './_media-timing.ts';
import {
  NS,
  type XmlElement,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  qname,
  walkElements,
} from '../../internal/xml/index.ts';

export type { AudioFormat, VideoFormat };

/** What `addSlideMedia` embeds or links. */
export type SlideMediaSource =
  | {
      kind: 'video';
      data: Uint8Array;
      /** Container of `data`. Detected from its signature when omitted. */
      format?: VideoFormat;
    }
  | {
      kind: 'audio';
      data: Uint8Array;
      /** Container of `data`. Detected from its signature when omitted. */
      format?: AudioFormat;
    }
  | {
      kind: 'online';
      /** `http(s)` URL of the video. YouTube page URLs are rewritten to the embed form. */
      url: string;
    };

export type SlideMediaOptions = SlideMediaSource & {
  x: Emu;
  y: Emu;
  w: Emu;
  h: Emu;
  name?: string;
  /** Poster frame shown until the clip plays. A neutral play-button image when omitted. */
  poster?: Uint8Array;
  /** Format of `poster`. Detected from its signature when omitted. */
  posterFormat?: ImageFormat;
};

/** A picture's media, as `getShapeMedia` reports it. */
export type ShapeMedia =
  | {
      readonly kind: 'video' | 'audio';
      readonly partName: string;
      readonly contentType: string;
      /** Live view into the package part — copy before mutating. */
      readonly bytes: Uint8Array;
    }
  | {
      /** Media linked by URL instead of embedded (an online video, or a linked file). */
      readonly kind: 'online';
      readonly url: string;
    };

// 320×180 1-bit PNG: white play triangle on dark grey. `<a:blip>` must point at
// a real image, so a clip added without a poster still needs one. `atob` is a
// Web-standard global in every supported runtime (Node >= 16, browsers).
const DEFAULT_POSTER_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAUAAAAC0AQMAAADfKmdSAAAABlBMVEUmJib///+/9FOwAAAAvUlEQVR42u3XMQ7DIBAEQOfn' +
  'fhpP4QmUFAgiJUqdKa6xtFtPA8J36+tKkiRJkiR5RF4Mm8KucCrcCg/DpnAoXAq3wnMr7AqnwqXwMGwKu8KpcCukK//ApnAo' +
  'XAq3QjnNF3aFU+FSeBg2hUPhVLgV/r/yH2wKRzls1YfZ5Rc+yh9F+cOd5R/XXT0AVvmQatWDlEczD3teH7yQypcmr2Fe7FoV' +
  'uHxwneGCxJWrvBZy0eTqymU4fxhJkiRJkjwqbzMmJGbQfp8bAAAAAElFTkSuQmCC';

const defaultPoster = (): Uint8Array =>
  Uint8Array.from(atob(DEFAULT_POSTER_BASE64), (ch) => ch.charCodeAt(0));

const YOUTUBE_HOSTS = new Set(['youtube.com', 'www.youtube.com', 'm.youtube.com']);
const YOUTUBE_SHORT_HOST = 'youtu.be';
const YOUTUBE_ID = /^[A-Za-z0-9_-]+$/;

const youtubeVideoId = (url: URL): string | null => {
  const segments = url.pathname.split('/').filter((s) => s !== '');
  if (url.hostname === YOUTUBE_SHORT_HOST) return segments[0] ?? null;
  if (!YOUTUBE_HOSTS.has(url.hostname)) return null;
  if (segments[0] === 'watch') return url.searchParams.get('v');
  if (segments[0] === 'shorts' || segments[0] === 'live') return segments[1] ?? null;
  return null;
};

// PowerPoint loads the relationship target directly in its embedded player, so
// a YouTube *page* URL (watch / youtu.be / shorts) would render the whole web
// page instead of the clip. Rewrite those to the embed endpoint, in the exact
// form PowerPoint's own "Insert > Online Video" stores. Anything else —
// including a URL already in embed form — is kept as given.
const resolveOnlineVideoUrl = (input: string): string => {
  let url: URL;
  try {
    url = new URL(input);
  } catch {
    throw new Error(`addSlideMedia: url is not an absolute URL: ${JSON.stringify(input)}`);
  }
  if (url.protocol !== 'http:' && url.protocol !== 'https:') {
    throw new Error(`addSlideMedia: url must be http(s), got ${JSON.stringify(input)}`);
  }
  const id = youtubeVideoId(url);
  if (id === null || !YOUTUBE_ID.test(id)) return url.href;
  return `https://www.youtube.com/embed/${id}?feature=oembed`;
};

const MEDIA_DIR = '/ppt/media/';

const nextMediaPartName = (pkg: OpcPackage, stem: string, extension: string): PartName => {
  let nextN = 1;
  const pattern = new RegExp(`^/ppt/media/${stem}(\\d+)\\.`);
  for (const p of pkg.parts) {
    const m = p.name.match(pattern);
    if (m?.[1] !== undefined) {
      const n = Number.parseInt(m[1], 10);
      if (n >= nextN) nextN = n + 1;
    }
  }
  return partName(`${MEDIA_DIR}${stem}${nextN}.${extension}`);
};

const bytesEqual = (a: Uint8Array, b: Uint8Array): boolean => {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) return false;
  }
  return true;
};

// Clips are large and decks often repeat one (the same jingle on every slide),
// so identical bytes share a single part. Posters are deliberately NOT shared:
// `setShapeImage` rewrites a picture's image part in place, which would change
// the poster of every other clip sharing it.
const internClipPart = (
  pkg: OpcPackage,
  extension: string,
  contentType: string,
  bytes: Uint8Array,
): PartName => {
  for (const p of pkg.parts) {
    if (p.name.startsWith(MEDIA_DIR) && p.contentType === contentType && bytesEqual(p.data, bytes))
      return p.name;
  }
  const name = nextMediaPartName(pkg, 'media', extension);
  setOpcDefault(pkg, extension, contentType);
  pkg.addPart(name, contentType, bytes);
  return name;
};

type ResolvedSource =
  | { readonly url: string }
  | { readonly data: Uint8Array; readonly extension: string; readonly contentType: string };

const resolveSource = (source: SlideMediaSource): ResolvedSource => {
  switch (source.kind) {
    case 'online':
      return { url: resolveOnlineVideoUrl(source.url) };
    case 'video': {
      const format = source.format ?? detectVideoFormat(source.data);
      if (format === null) {
        throw new Error(
          'addSlideMedia: could not detect the video format. Pass options.format explicitly.',
        );
      }
      const contentType = contentTypeForVideoFormat(format);
      return { data: source.data, extension: format, contentType };
    }
    case 'audio': {
      const format = source.format ?? detectAudioFormat(source.data);
      if (format === null) {
        throw new Error(
          'addSlideMedia: could not detect the audio format. Pass options.format explicitly.',
        );
      }
      const contentType = contentTypeForAudioFormat(format);
      return { data: source.data, extension: format, contentType };
    }
  }
};

const DEFAULT_NAME_PREFIX = { video: 'Video', audio: 'Audio', online: 'Online Media' } as const;

/**
 * Adds a video, an audio clip or an online (URL) video to the slide. Returns
 * the new picture shape — the clip's poster frame.
 *
 * Embedded clips are stored once per package: adding the same bytes again
 * reuses the existing `/ppt/media/mediaN.<ext>` part. The container format is
 * detected from the bytes (mp4 / m4v / mov / webm / avi / wmv, mp3 / wav / m4a /
 * ogg / wma); pass `format` for anything the signature check cannot tell.
 * Whether a clip plays depends on the codecs of the viewing application.
 *
 * The shape is an ordinary picture: `getShapeImageBytes` / `setShapeImage`
 * read and replace the poster, `getShapeMedia` reads the clip.
 */
export const addSlideMedia = (slide: SlideData, opts: SlideMediaOptions): SlideShapeData => {
  const x = emuCoordinate(opts.x, 'addSlideMedia: x');
  const y = emuCoordinate(opts.y, 'addSlideMedia: y');
  const w = emuExtent(opts.w, 'addSlideMedia: w');
  const h = emuExtent(opts.h, 'addSlideMedia: h');

  // Resolve every input before touching the package so a rejected call leaves
  // no orphan part behind.
  const source = resolveSource(opts);
  const poster = opts.poster ?? defaultPoster();
  const posterFormat = opts.posterFormat ?? detectImageFormat(poster);
  if (posterFormat === null) {
    throw new Error(
      'addSlideMedia: could not detect the poster image format. Pass options.posterFormat explicitly.',
    );
  }

  const pkg = slide[INTERNAL_PACKAGE];
  const slidePartName = slide[SLIDE_PART_NAME];
  const rels = pkg.getRels(slidePartName) ?? emptyRels();
  const usedIds = rels.items.map((r) => r.id);
  // Two clips on one slide that share a part (or URL) share its relationships.
  const relate = (type: string, target: string, targetMode: TargetMode): string => {
    const existing = rels.items.find(
      (r) => r.type === type && r.target === target && r.targetMode === targetMode,
    );
    if (existing) return existing.id;
    const id = nextRelId(usedIds);
    usedIds.push(id);
    rels.items.push({ id, type, target, targetMode });
    return id;
  };
  const relativeTarget = (name: PartName): string => `../media/${name.slice(MEDIA_DIR.length)}`;

  const kind = opts.kind === 'audio' ? 'audio' : 'video';
  const fileRelType = kind === 'audio' ? REL_TYPES.audio : REL_TYPES.video;
  let rLink: string;
  let rMedia: string | undefined;
  if ('url' in source) {
    rLink = relate(fileRelType, source.url, 'External');
  } else {
    const clipPart = internClipPart(pkg, source.extension, source.contentType, source.data);
    // PowerPoint orders the `media` rel before the `video` / `audio` one.
    rMedia = relate(REL_TYPES.media, relativeTarget(clipPart), 'Internal');
    rLink = relate(fileRelType, relativeTarget(clipPart), 'Internal');
  }

  const posterExtension = extensionForFormat(posterFormat);
  const posterContentType = contentTypeForFormat(posterFormat);
  const posterPart = nextMediaPartName(pkg, 'image', posterExtension);
  setOpcDefault(pkg, posterExtension, posterContentType);
  pkg.addPart(posterPart, posterContentType, poster);
  const rPoster = relate(REL_TYPES.image, relativeTarget(posterPart), 'Internal');
  pkg.setRels(slidePartName, rels);

  const id = nextShapeId(slide);
  const pic = buildMediaPicture({
    id,
    name: opts.name ?? `${DEFAULT_NAME_PREFIX[opts.kind]} ${id}`,
    kind,
    rLink,
    ...(rMedia !== undefined ? { rMedia } : {}),
    rEmbed: rPoster,
    x,
    y,
    w,
    h,
  });
  addMediaTimingNode(slide, kind, id);
  return appendAndReturnNewShape(slide, pic);
};

/**
 * Returns the clip behind a media picture, or `null` when the shape is not
 * one. Reads PowerPoint-, PptxGenJS- and python-pptx-authored media alike: the
 * embedded part is taken from `<p14:media r:embed>` when present and from the
 * DrawingML `r:link` otherwise; a relationship with an external target is
 * reported as `'online'`.
 */
export const getShapeMedia = (shape: SlideShapeData): ShapeMedia | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') return null;
  const ref = readPictureMediaRef(shape[SHAPE_ELEMENT]);
  if (ref === null) return null;
  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
  if (rels === null) return null;
  const relsById = new Map(rels.items.map((r) => [r.id, r]));

  for (const rId of [ref.mediaRId, ref.fileRId]) {
    const rel = rId === null ? undefined : relsById.get(rId);
    if (rel === undefined) continue;
    if (rel.targetMode === 'External') return { kind: 'online', url: rel.target };
    const name = rel.target.startsWith('/')
      ? partName(rel.target)
      : resolveTarget(slide[SLIDE_PART_NAME], rel.target);
    const part = pkg.getPart(name);
    if (part) {
      return {
        kind: ref.kind,
        partName: part.name,
        contentType: part.contentType,
        bytes: part.data,
      };
    }
  }
  return null;
};

/** Every shape on the slide that `getShapeMedia` reports a clip for. */
export const findShapesWithMedia = (slide: SlideData): ReadonlyArray<SlideShapeData> =>
  slide[SLIDE_SHAPES].filter((shape) => getShapeMedia(shape) !== null);

/** Durations removed from the source clip, in milliseconds. */
export interface MediaTrim {
  readonly startMs: number;
  /** Duration removed from the end of the source clip. */
  readonly endMs: number;
}

export interface MediaFade {
  readonly inMs: number;
  readonly outMs: number;
}

/** A named position in an embedded media clip, measured from its beginning. */
export interface MediaBookmark {
  readonly name: string;
  readonly timeMs: number;
}

/**
 * How a clip plays in the slide show — the attributes of its
 * `<p:cMediaNode>` and the start condition of its time node.
 *
 * Trimming and fades are stored in the PowerPoint 2010 `p14:media` extension.
 * `trim.endMs` is the duration removed from the end of the clip, matching the
 * OOXML `p14:trim@end` meaning (it is not the playback end position).
 */
export interface MediaPlayback {
  /** Starts with the slide instead of waiting for a click. */
  readonly autoplay: boolean;
  /** Milliseconds to wait after the slide starts before automatic playback. */
  readonly delayMs?: number;
  /** Plays again from the beginning until the slide moves on. */
  readonly loop: boolean;
  /** Playback volume, 0–1. PowerPoint's own default is 0.8. */
  readonly volume: number;
  readonly muted: boolean;
  /** Video only: plays filling the screen. */
  readonly fullScreen: boolean;
  /** Hides the clip once it has played (`showWhenStopped="0"`). */
  readonly hideWhenStopped: boolean;
  /** Number of slides across which the clip should keep playing (`numSld`).
   * When absent, OOXML defaults to one slide. */
  readonly slideCount?: number;
  /** Returns playback to the beginning after natural completion (`fill="remove"`). */
  readonly rewindAfterPlaying?: boolean;
  readonly trim?: MediaTrim;
  readonly fade?: MediaFade;
  /** Named positions in the clip, in the order stored in the media extension. */
  readonly bookmarks?: readonly MediaBookmark[];
}

const NAME_C_MEDIA_NODE = qname('p', 'cMediaNode', NS.pml);
const NAME_C_TN = qname('p', 'cTn', NS.pml);
const NAME_ST_COND_LST = qname('p', 'stCondLst', NS.pml);
const NAME_COND = qname('p', 'cond', NS.pml);
const ATTR_VOL = qname('', 'vol', '');
const ATTR_MUTE = qname('', 'mute', '');
const ATTR_FULL_SCRN = qname('', 'fullScrn', '');
const ATTR_FILL = qname('', 'fill', '');
const ATTR_SHOW_WHEN_STOPPED = qname('', 'showWhenStopped', '');
const ATTR_NUM_SLD = qname('', 'numSld', '');
const MAX_MEDIA_SLIDE_COUNT = 0xffffffff;
const ATTR_REPEAT_COUNT = qname('', 'repeatCount', '');
const ATTR_DELAY = qname('', 'delay', '');
const ATTR_EVT = qname('', 'evt', '');
const ATTR_MASTER_REL = qname('', 'masterRel', '');
const ATTR_CMD = qname('', 'cmd', '');
const ATTR_NODE_TYPE = qname('', 'nodeType', '');
const ATTR_ID = qname('', 'id', '');
const ATTR_VAL = qname('', 'val', '');
const NAME_C_BHVR = qname('p', 'cBhvr', NS.pml);
const NAME_TGT_EL = qname('p', 'tgtEl', NS.pml);
const NAME_SP_TGT = qname('p', 'spTgt', NS.pml);
const NAME_TN = qname('p', 'tn', NS.pml);
const NAME_CHILD_TN_LST = qname('p', 'childTnLst', NS.pml);
const NAME_NV_PIC_PR = qname('p', 'nvPicPr', NS.pml);
const NAME_NV_PR = qname('p', 'nvPr', NS.pml);
const NAME_EXT_LST = qname('p', 'extLst', NS.pml);
const NAME_EXT = qname('p', 'ext', NS.pml);
const NAME_P14_MEDIA = qname('p14', 'media', NS.p14);
const NAME_P14_TRIM = qname('p14', 'trim', NS.p14);
const NAME_P14_FADE = qname('p14', 'fade', NS.p14);
const NAME_P14_BMK_LST = qname('p14', 'bmkLst', NS.p14);
const NAME_P14_BMK = qname('p14', 'bmk', NS.p14);
const NAME_P14_BMK_TGT = qname('p14', 'bmkTgt', NS.p14);
const ATTR_URI = qname('', 'uri', '');
const ATTR_ST = qname('', 'st', '');
const ATTR_END = qname('', 'end', '');
const ATTR_IN = qname('', 'in', '');
const ATTR_OUT = qname('', 'out', '');
const ATTR_NAME = qname('', 'name', '');
const ATTR_TIME = qname('', 'time', '');
const ATTR_SPID = qname('', 'spid', '');
const ATTR_BMK_NAME = qname('', 'bmkName', '');
const P14_MEDIA_EXT_URI = '{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}';

// ST_PositiveFixedPercentage accepts both `80000` and `80%`; PowerPoint writes
// the integer form, and the schema's own default is spelled `50%`.
const percentFraction = (raw: string | null, fallback: number): number => {
  if (raw === null) return fallback;
  const value = Number.parseFloat(raw.endsWith('%') ? raw.slice(0, -1) : raw);
  if (!Number.isFinite(value)) return fallback;
  return raw.endsWith('%') ? value / 100 : value / 100000;
};

const xsdBoolean = (raw: string | null, fallback: boolean): boolean =>
  raw === null ? fallback : raw === '1' || raw === 'true';

const setOrRemove = (
  el: XmlElement,
  name: ReturnType<typeof qname>,
  value: string | null,
): void => {
  el.attrs = el.attrs.filter(
    (a) => a.name.localName !== name.localName || a.name.namespaceURI !== name.namespaceURI,
  );
  if (value !== null) el.attrs.push(attr(name, value));
};

// Bookmark attributes are edited in place so unrelated attributes keep their
// original order and any future PowerPoint metadata remains untouched.
const setOrRemovePreservingOrder = (
  el: XmlElement,
  name: ReturnType<typeof qname>,
  value: string | null,
): void => {
  const index = el.attrs.findIndex(
    (a) => a.name.localName === name.localName && a.name.namespaceURI === name.namespaceURI,
  );
  if (value === null) {
    if (index >= 0) el.attrs.splice(index, 1);
  } else if (index >= 0) {
    el.attrs[index] = attr(name, value);
  } else {
    el.attrs.push(attr(name, value));
  }
};

const p14MediaOf = (shape: SlideShapeData): XmlElement | null => {
  const nvPicPr = firstChildElement(shape[SHAPE_ELEMENT], NAME_NV_PIC_PR);
  const nvPr = nvPicPr === null ? null : firstChildElement(nvPicPr, NAME_NV_PR);
  const extLst = nvPr === null ? null : firstChildElement(nvPr, NAME_EXT_LST);
  if (extLst === null) return null;
  for (const child of extLst.children) {
    if (child.kind !== 'element' || !qnameSame(child.name, NAME_EXT)) continue;
    if (getAttrValue(child, ATTR_URI) !== P14_MEDIA_EXT_URI) continue;
    const media = firstChildElement(child, NAME_P14_MEDIA);
    if (media !== null) return media;
  }
  return null;
};

const qnameSame = (a: ReturnType<typeof qname>, b: ReturnType<typeof qname>): boolean =>
  a.namespaceURI === b.namespaceURI && a.localName === b.localName;

// ST_UniversalTimeOffset is expressed in milliseconds when no suffix is
// present. PowerPoint commonly writes fractional milliseconds in this form.
const universalTimeMs = (raw: string | null): number | null => {
  if (raw === null) return null;
  const match = /^([+]?(?:\d+(?:\.\d*)?|\.\d+))(ms|s|min|h|µs|ns)?$/.exec(raw);
  if (match === null) return null;
  const value = Number(match[1]);
  if (!Number.isFinite(value)) return null;
  const multiplier =
    match[2] === 's'
      ? 1000
      : match[2] === 'min'
        ? 60000
        : match[2] === 'h'
          ? 3600000
          : match[2] === 'µs'
            ? 0.001
            : match[2] === 'ns'
              ? 0.000001
              : 1;
  const result = value * multiplier;
  return Number.isFinite(result) ? result : null;
};

const mediaAdjustments = (
  shape: SlideShapeData,
): Pick<MediaPlayback, 'trim' | 'fade' | 'bookmarks'> => {
  const media = p14MediaOf(shape);
  if (media === null) return {};
  const trim = firstChildElement(media, NAME_P14_TRIM);
  const fade = firstChildElement(media, NAME_P14_FADE);
  const startMs = trim === null ? null : universalTimeMs(getAttrValue(trim, ATTR_ST));
  const endMs = trim === null ? null : universalTimeMs(getAttrValue(trim, ATTR_END));
  const inMs = fade === null ? null : universalTimeMs(getAttrValue(fade, ATTR_IN));
  const outMs = fade === null ? null : universalTimeMs(getAttrValue(fade, ATTR_OUT));
  const bookmarkList = firstChildElement(media, NAME_P14_BMK_LST);
  const bookmarks =
    bookmarkList === null
      ? undefined
      : bookmarkList.children
          .filter(
            (child): child is XmlElement =>
              child.kind === 'element' && qnameSame(child.name, NAME_P14_BMK),
          )
          .map((bookmark) => {
            const name = getAttrValue(bookmark, ATTR_NAME);
            const timeMs = universalTimeMs(getAttrValue(bookmark, ATTR_TIME));
            // Both attributes are required by CT_MediaBookmark. Do not turn
            // malformed external XML into a plausible typed value: callers
            // must repair the source before reading or editing playback.
            if (name === null || timeMs === null || timeMs < 0)
              throw new Error('getShapeMediaPlayback: malformed media bookmark');
            return { name, timeMs };
          });
  return {
    ...(trim === null || (startMs === null && endMs === null)
      ? {}
      : { trim: { startMs: startMs ?? 0, endMs: endMs ?? 0 } }),
    ...(fade === null || (inMs === null && outMs === null)
      ? {}
      : { fade: { inMs: inMs ?? 0, outMs: outMs ?? 0 } }),
    ...(bookmarks === undefined ? {} : { bookmarks }),
  };
};

const validateAdjustment = (value: MediaTrim | MediaFade, name: 'trim' | 'fade'): void => {
  if (typeof value !== 'object' || value === null) {
    throw new Error(`setShapeMediaPlayback: ${name} durations must be finite and nonnegative`);
  }
  if (name === 'trim' && !('startMs' in value)) {
    throw new Error(`setShapeMediaPlayback: ${name} durations must be finite and nonnegative`);
  }
  if (name === 'fade' && !('inMs' in value)) {
    throw new Error(`setShapeMediaPlayback: ${name} durations must be finite and nonnegative`);
  }
  const values =
    name === 'trim'
      ? 'startMs' in value
        ? [value.startMs, value.endMs]
        : []
      : 'inMs' in value
        ? [value.inMs, value.outMs]
        : [];
  if (values.some((item) => !Number.isFinite(item) || item < 0)) {
    throw new Error(`setShapeMediaPlayback: ${name} durations must be finite and nonnegative`);
  }
};

// XML Schema decimals do not allow JavaScript's exponent notation. Preserve
// the exact decimal spelling of finite values such as 0.0000001ms.
const universalTimeString = (value: number): string => {
  const raw = String(value);
  if (!/[eE]/.test(raw)) return raw;
  const match = /^([+-]?)(\d+)(?:\.(\d+))?[eE]([+-]?\d+)$/.exec(raw);
  if (match === null) return raw;
  const digits = `${match[2]}${match[3] ?? ''}`;
  const point = match[2]!.length + Number(match[4]);
  if (point <= 0) return `${match[1]}0.${'0'.repeat(-point)}${digits}`;
  if (point >= digits.length) return `${match[1]}${digits}${'0'.repeat(point - digits.length)}`;
  return `${match[1]}${digits.slice(0, point)}.${digits.slice(point)}`;
};

const removeMediaChild = (media: XmlElement, name: ReturnType<typeof qname>): void => {
  media.children = media.children.filter(
    (child) => child.kind !== 'element' || !qnameSame(child.name, name),
  );
};

const clearMediaChild = (
  media: XmlElement,
  name: ReturnType<typeof qname>,
  attrs: ReadonlyArray<ReturnType<typeof qname>>,
): void => {
  const child = firstChildElement(media, name);
  if (child === null) return;
  for (const attribute of attrs) setOrRemove(child, attribute, null);
  // Unknown attributes, children, and namespace declarations belong to the
  // source document. Keep the element when any of them remain.
  if (child.attrs.length === 0 && child.children.length === 0 && child.prefixDecls.size === 0) {
    removeMediaChild(media, name);
  }
};

const adjustmentRank = (name: ReturnType<typeof qname>): number =>
  qnameSame(name, NAME_P14_TRIM)
    ? 0
    : qnameSame(name, NAME_P14_FADE)
      ? 1
      : qnameSame(name, NAME_P14_BMK_LST)
        ? 2
        : qnameSame(name, NAME_EXT_LST)
          ? 3
          : 4;

const upsertMediaChild = (
  media: XmlElement,
  name: ReturnType<typeof qname>,
  attrs: ReadonlyArray<[ReturnType<typeof qname>, string]>,
): void => {
  let child = firstChildElement(media, name);
  if (child === null) {
    // The containing `<p14:media>` owns the namespace declaration.
    child = elem(name);
    const rank = adjustmentRank(name);
    const index = media.children.findIndex(
      (candidate) => candidate.kind === 'element' && adjustmentRank(candidate.name) > rank,
    );
    if (index < 0) media.children.push(child);
    else media.children.splice(index, 0, child);
  }
  for (const [attribute, value] of attrs) setOrRemove(child, attribute, value);
};

const setMediaAdjustments = (
  media: XmlElement,
  trim: MediaTrim | undefined,
  fade: MediaFade | undefined,
): void => {
  if (trim !== undefined) {
    if (trim.startMs === 0 && trim.endMs === 0)
      clearMediaChild(media, NAME_P14_TRIM, [ATTR_ST, ATTR_END]);
    else
      upsertMediaChild(media, NAME_P14_TRIM, [
        [ATTR_ST, universalTimeString(trim.startMs)],
        [ATTR_END, universalTimeString(trim.endMs)],
      ]);
  }
  if (fade !== undefined) {
    if (fade.inMs === 0 && fade.outMs === 0)
      clearMediaChild(media, NAME_P14_FADE, [ATTR_IN, ATTR_OUT]);
    else
      upsertMediaChild(media, NAME_P14_FADE, [
        [ATTR_IN, universalTimeString(fade.inMs)],
        [ATTR_OUT, universalTimeString(fade.outMs)],
      ]);
  }
};

const validateBookmarks = (value: readonly MediaBookmark[]): void => {
  if (!Array.isArray(value)) {
    throw new Error('setShapeMediaPlayback: bookmarks must be an array');
  }
  const names = new Set<string>();
  const times = new Set<number>();
  for (const bookmark of value) {
    if (
      typeof bookmark !== 'object' ||
      bookmark === null ||
      typeof bookmark.name !== 'string' ||
      !Number.isFinite(bookmark.timeMs) ||
      bookmark.timeMs < 0
    ) {
      throw new Error(
        'setShapeMediaPlayback: bookmarks require a string name and finite nonnegative timeMs',
      );
    }
    if (names.has(bookmark.name)) {
      throw new Error('setShapeMediaPlayback: bookmark names must be unique');
    }
    if (times.has(bookmark.timeMs)) {
      throw new Error('setShapeMediaPlayback: bookmark times must be unique');
    }
    names.add(bookmark.name);
    times.add(bookmark.timeMs);
  }
};

const validateBookmarkReferences = (
  shape: SlideShapeData,
  bookmarks: readonly MediaBookmark[],
): void => {
  const root = shape[SHAPE_SLIDE][SLIDE_DOCUMENT].root;
  const spid = String(shape[SHAPE_SNAPSHOT].id);
  const names = new Set(bookmarks.map((bookmark) => bookmark.name));
  let dangling: string | null = null;
  walkElements(root, (element) => {
    if (!qnameSame(element.name, NAME_P14_BMK_TGT)) return;
    if (getAttrValue(element, ATTR_SPID) !== spid) return;
    const name = getAttrValue(element, ATTR_BMK_NAME);
    if (name !== null && !names.has(name)) dangling = name;
  });
  if (dangling !== null) {
    throw new Error(
      `setShapeMediaPlayback: cannot remove or rename referenced bookmark ${JSON.stringify(dangling)}`,
    );
  }
};

const setMediaBookmarks = (media: XmlElement, bookmarks: readonly MediaBookmark[]): void => {
  const list = firstChildElement(media, NAME_P14_BMK_LST);
  if (list === null && bookmarks.length === 0) return;

  if (list === null) {
    const created = elem(NAME_P14_BMK_LST, {
      children: bookmarks.map((bookmark) =>
        elem(NAME_P14_BMK, {
          attrs: [
            attr(ATTR_NAME, bookmark.name),
            attr(ATTR_TIME, universalTimeString(bookmark.timeMs)),
          ],
        }),
      ),
    });
    const rank = adjustmentRank(NAME_P14_BMK_LST);
    const index = media.children.findIndex(
      (candidate) => candidate.kind === 'element' && adjustmentRank(candidate.name) > rank,
    );
    if (index < 0) media.children.push(created);
    else media.children.splice(index, 0, created);
    return;
  }

  const existing = list.children.filter(
    (child): child is XmlElement => child.kind === 'element' && qnameSame(child.name, NAME_P14_BMK),
  );
  const byIdentity = new Map<string, XmlElement>();
  const byName = new Map<string, XmlElement>();
  const byTime = new Map<number, XmlElement>();
  for (const candidate of existing) {
    const name = getAttrValue(candidate, ATTR_NAME) ?? '';
    if (!byName.has(name)) byName.set(name, candidate);
    const rawTime = getAttrValue(candidate, ATTR_TIME);
    const time = universalTimeMs(rawTime);
    if (time === null) continue;
    const key = `${name}\u0000${time}`;
    if (!byIdentity.has(key)) byIdentity.set(key, candidate);
    if (!byTime.has(time)) byTime.set(time, candidate);
  }
  const unused = new Set(existing);
  const assignments: Array<XmlElement | undefined> = bookmarks.map(() => undefined);

  // Reserve all strong matches before assigning positional fallbacks. This
  // prevents a newly inserted bookmark at the front from stealing a source
  // node that a later unchanged bookmark should retain.
  bookmarks.forEach((bookmark, index) => {
    const match = byIdentity.get(`${bookmark.name}\u0000${bookmark.timeMs}`);
    if (match !== undefined && unused.has(match)) {
      assignments[index] = match;
      unused.delete(match);
    }
  });
  bookmarks.forEach((bookmark, index) => {
    if (assignments[index] !== undefined) return;
    const match = byName.get(bookmark.name);
    if (match !== undefined && unused.has(match)) {
      assignments[index] = match;
      unused.delete(match);
    }
  });
  bookmarks.forEach((bookmark, index) => {
    if (assignments[index] !== undefined) return;
    const match = byTime.get(bookmark.timeMs);
    if (match !== undefined && unused.has(match)) {
      assignments[index] = match;
      unused.delete(match);
    }
  });
  let fallbackIndex = 0;
  bookmarks.forEach((_, index) => {
    if (assignments[index] !== undefined) return;
    while (fallbackIndex < existing.length) {
      const candidate = existing[fallbackIndex++];
      if (candidate !== undefined && unused.has(candidate)) {
        assignments[index] = candidate;
        unused.delete(candidate);
        return;
      }
    }
  });

  const next = bookmarks.map((bookmark, index) => {
    const match = assignments[index];
    if (match !== undefined) {
      if (getAttrValue(match, ATTR_NAME) !== bookmark.name)
        setOrRemovePreservingOrder(match, ATTR_NAME, bookmark.name);
      const rawTime = getAttrValue(match, ATTR_TIME);
      if (universalTimeMs(rawTime) !== bookmark.timeMs)
        setOrRemovePreservingOrder(match, ATTR_TIME, universalTimeString(bookmark.timeMs));
      return match;
    }
    return elem(NAME_P14_BMK, {
      attrs: [
        attr(ATTR_NAME, bookmark.name),
        attr(ATTR_TIME, universalTimeString(bookmark.timeMs)),
      ],
    });
  });
  const unknownChildren = list.children.filter(
    (child) => child.kind !== 'element' || !qnameSame(child.name, NAME_P14_BMK),
  );
  const firstUnknownElement = unknownChildren.findIndex((child) => child.kind === 'element');
  if (firstUnknownElement < 0) list.children = [...unknownChildren, ...next];
  else
    list.children = [
      ...unknownChildren.slice(0, firstUnknownElement),
      ...next,
      ...unknownChildren.slice(firstUnknownElement),
    ];

  // An empty, untouched list has no schema-visible purpose. Keep it when it
  // carries an extension, attributes, or other future content.
  if (
    next.length === 0 &&
    list.attrs.length === 0 &&
    list.children.length === 0 &&
    list.prefixDecls.size === 0
  ) {
    removeMediaChild(media, NAME_P14_BMK_LST);
  }
};

const mediaNodeOf = (
  shape: SlideShapeData,
): (MediaTimingPath & { readonly media: XmlElement }) | null => {
  const found = findMediaTimingNodeWithAncestors(shape[SHAPE_SLIDE], shape[SHAPE_SNAPSHOT].id);
  const node = found?.node ?? null;
  const media = node === null ? null : firstChildElement(node, NAME_C_MEDIA_NODE);
  return found !== null && media !== null ? { ...found, media } : null;
};

type MediaStart = { readonly automatic: boolean; readonly delayMs: number } | null;

type MediaCommandTiming = 'background' | 'interactive' | 'other' | null;

const mediaCommandTiming = (shape: SlideShapeData): MediaCommandTiming => {
  const root = shape[SHAPE_SLIDE][SLIDE_DOCUMENT].root;
  const spid = String(shape[SHAPE_SNAPSHOT].id);
  const matches: Array<{ element: XmlElement; ancestors: ReadonlyArray<XmlElement> }> = [];
  const walk = (element: XmlElement, ancestors: ReadonlyArray<XmlElement>): void => {
    const nextAncestors =
      element.name.namespaceURI === NS.pml && element.name.localName === 'cTn'
        ? [...ancestors, element]
        : ancestors;
    if (
      element.name.namespaceURI === NS.pml &&
      element.name.localName === 'cmd' &&
      getAttrValue(element, ATTR_CMD)?.startsWith('playFrom(')
    ) {
      const behavior = firstChildElement(element, NAME_C_BHVR);
      const target = behavior === null ? null : firstChildElement(behavior, NAME_TGT_EL);
      const spTarget = target === null ? null : firstChildElement(target, NAME_SP_TGT);
      if (spTarget !== null && getAttrValue(spTarget, qname('', 'spid', '')) === spid) {
        matches.push({ element, ancestors: nextAncestors });
      }
    }
    for (const child of element.children) {
      if (child.kind === 'element') walk(child, nextAncestors);
    }
  };
  walk(root, []);
  if (matches.length === 0) return null;
  if (matches.length !== 1) return 'other';

  const { element: commandElement, ancestors } = matches[0]!;
  const mainSequenceIndex = ancestors.findIndex(
    (ancestor) => getAttrValue(ancestor, ATTR_NODE_TYPE) === 'mainSeq',
  );
  const interactiveSequenceIndex = ancestors.findIndex(
    (ancestor) => getAttrValue(ancestor, ATTR_NODE_TYPE) === 'interactiveSeq',
  );
  const effectIndex = ancestors.findIndex(
    (ancestor) =>
      getAttrValue(ancestor, ATTR_NODE_TYPE) ===
      (mainSequenceIndex >= 0 ? 'afterEffect' : 'clickEffect'),
  );
  const sequenceIndex = mainSequenceIndex >= 0 ? mainSequenceIndex : interactiveSequenceIndex;
  if (sequenceIndex < 0 || effectIndex <= sequenceIndex) return 'other';
  const mainSequence = ancestors[sequenceIndex]!;
  const startGroup = ancestors[sequenceIndex + 1];
  if (startGroup === undefined) return 'other';
  const firstCtnIn = (element: XmlElement): XmlElement | null => {
    if (element.name.namespaceURI === NS.pml && element.name.localName === 'cTn') return element;
    for (const child of element.children) {
      if (child.kind !== 'element') continue;
      const first = firstCtnIn(child);
      if (first !== null) return first;
    }
    return null;
  };
  const mainChildren = firstChildElement(mainSequence, NAME_CHILD_TN_LST);
  if (mainChildren === null || firstCtnIn(mainChildren) !== startGroup) return 'other';
  const startList = firstChildElement(startGroup, NAME_ST_COND_LST);
  if (startList === null) return 'other';
  const startConditions = startList.children.filter(
    (child): child is XmlElement =>
      child.kind === 'element' &&
      child.name.namespaceURI === NS.pml &&
      child.name.localName === 'cond',
  );
  const sequenceStartList = firstChildElement(mainSequence, NAME_ST_COND_LST);
  const sequenceStartConditions =
    sequenceStartList?.children.filter(
      (child): child is XmlElement =>
        child.kind === 'element' &&
        child.name.namespaceURI === NS.pml &&
        child.name.localName === 'cond',
    ) ?? [];
  const hasNativeStart = startConditions.some((condition) => {
    if (getAttrValue(condition, ATTR_EVT) !== 'onBegin') return false;
    if (getAttrValue(condition, ATTR_DELAY) !== '0') return false;
    const reference = firstChildElement(condition, NAME_TN);
    return (
      reference !== null &&
      getAttrValue(reference, ATTR_VAL) === getAttrValue(mainSequence, ATTR_ID)
    );
  });
  const hasIndefiniteStart = startConditions.some(
    (condition) =>
      getAttrValue(condition, ATTR_EVT) === null &&
      getAttrValue(condition, ATTR_DELAY) === 'indefinite',
  );
  const interactiveTarget =
    sequenceStartConditions.length === 1
      ? firstChildElement(
          firstChildElement(sequenceStartConditions[0]!, NAME_TGT_EL) ?? elem(NAME_TGT_EL),
          NAME_SP_TGT,
        )
      : null;
  const hasInteractiveStart =
    sequenceStartConditions.length === 1 &&
    getAttrValue(sequenceStartConditions[0]!, ATTR_EVT) === 'onClick' &&
    getAttrValue(sequenceStartConditions[0]!, ATTR_DELAY) === '0' &&
    getAttrValue(interactiveTarget ?? elem(NAME_SP_TGT), qname('', 'spid', '')) === spid;
  if (
    mainSequenceIndex >= 0 &&
    (startConditions.length !== 2 || !hasNativeStart || !hasIndefiniteStart)
  )
    return 'other';
  if (
    interactiveSequenceIndex >= 0 &&
    (!hasInteractiveStart ||
      getAttrValue(ancestors[effectIndex]!, ATTR_NODE_TYPE) !== 'clickEffect')
  )
    return 'other';
  const zeroDelay = (cTn: XmlElement): boolean => {
    const list = firstChildElement(cTn, NAME_ST_COND_LST);
    if (list === null) return false;
    const conditions = list.children.filter(
      (child): child is XmlElement =>
        child.kind === 'element' &&
        child.name.namespaceURI === NS.pml &&
        child.name.localName === 'cond',
    );
    return (
      conditions.length === 1 &&
      getAttrValue(conditions[0]!, ATTR_EVT) === null &&
      getAttrValue(conditions[0]!, ATTR_DELAY) === '0' &&
      conditions[0]!.children.every((child) => child.kind !== 'element')
    );
  };
  if (interactiveSequenceIndex >= 0 && !zeroDelay(startGroup)) return 'other';
  if (!ancestors.slice(sequenceIndex + 2, effectIndex + 1).every((cTn) => zeroDelay(cTn))) {
    return 'other';
  }
  if (getAttrValue(commandElement, ATTR_CMD) !== 'playFrom(0.0)') return 'other';
  return mainSequenceIndex >= 0 ? 'background' : 'interactive';
};

// Mac PowerPoint uses these two dedicated command trees for background audio
// and When Clicked On. Match the complete tree before replacing it: changing a
// shared sequence's start condition would also change unrelated animations.
const nativeMediaSequence = (
  kind: 'background' | 'interactive',
  ids: readonly string[],
  spid: string,
): XmlElement => {
  const p = (
    name: string,
    attrs: Record<string, string> = {},
    children: XmlElement[] = [],
  ): XmlElement =>
    elem(qname('p', name, NS.pml), {
      attrs: Object.entries(attrs).map(([name, value]) => attr(qname('', name, ''), value)),
      children,
    });
  const click = (): XmlElement =>
    p('cond', { evt: 'onClick', delay: '0' }, [p('tgtEl', {}, [p('spTgt', { spid })])]);
  const zero = (): XmlElement => p('stCondLst', {}, [p('cond', { delay: '0' })]);
  const child = (node: XmlElement): XmlElement => p('childTnLst', {}, [node]);
  const background = kind === 'background';
  const effect = p('par', {}, [
    p(
      'cTn',
      {
        id: ids[3]!,
        presetID: '1',
        presetClass: 'mediacall',
        presetSubtype: '0',
        fill: 'hold',
        nodeType: background ? 'afterEffect' : 'clickEffect',
      },
      [
        zero(),
        child(
          p('cmd', { type: 'call', cmd: 'playFrom(0.0)' }, [
            p('cBhvr', {}, [
              p('cTn', { id: ids[4]!, dur: '1', fill: 'hold' }),
              p('tgtEl', {}, [p('spTgt', { spid })]),
            ]),
          ]),
        ),
      ],
    ),
  ]);
  const group = p('par', {}, [
    p('cTn', { id: ids[1]!, fill: 'hold' }, [
      background
        ? p('stCondLst', {}, [
            p('cond', { delay: 'indefinite' }),
            p('cond', { evt: 'onBegin', delay: '0' }, [p('tn', { val: ids[0]! })]),
          ])
        : zero(),
      child(p('par', {}, [p('cTn', { id: ids[2]!, fill: 'hold' }, [zero(), child(effect)])])),
    ]),
  ]);
  const sequence = background
    ? p('cTn', { id: ids[0]!, dur: 'indefinite', nodeType: 'mainSeq' }, [child(group)])
    : p(
        'cTn',
        {
          id: ids[0]!,
          restart: 'whenNotActive',
          fill: 'hold',
          evtFilter: 'cancelBubble',
          nodeType: 'interactiveSeq',
        },
        [
          p('stCondLst', {}, [click()]),
          p('endSync', { evt: 'end', delay: '0' }, [p('rtn', { val: 'all' })]),
          child(group),
        ],
      );
  const slideCondition = (name: string, evt: string): XmlElement =>
    p(name, {}, [p('cond', { evt, delay: '0' }, [p('tgtEl', {}, [p('sldTgt')])])]);
  return p('seq', { concurrent: '1', nextAc: 'seek' }, [
    sequence,
    ...(background
      ? [slideCondition('prevCondLst', 'onPrev'), slideCondition('nextCondLst', 'onNext')]
      : [p('nextCondLst', {}, [click()])]),
  ]);
};

const sameTimingTree = (actual: XmlElement, expected: XmlElement): boolean => {
  const sameName = (a: XmlElement['name'], b: XmlElement['name']): boolean =>
    a.namespaceURI === b.namespaceURI && a.localName === b.localName;
  if (!sameName(actual.name, expected.name) || actual.attrs.length !== expected.attrs.length)
    return false;
  const attrs = new Map(
    actual.attrs.map(({ name, value }) => [`${name.namespaceURI}:${name.localName}`, value]),
  );
  if (
    !expected.attrs.every(
      ({ name, value }) => attrs.get(`${name.namespaceURI}:${name.localName}`) === value,
    )
  )
    return false;
  if (
    actual.children.some((c) => c.kind !== 'element' && (c.kind !== 'text' || c.data.trim() !== ''))
  )
    return false;
  const children = actual.children.filter((c): c is XmlElement => c.kind === 'element');
  const expectedChildren = expected.children.filter((c): c is XmlElement => c.kind === 'element');
  return (
    children.length === expectedChildren.length &&
    children.every((c, i) => sameTimingTree(c, expectedChildren[i]!))
  );
};

const prepareDedicatedMediaTiming = (
  shape: SlideShapeData,
  from: 'background' | 'interactive',
): (() => void) => {
  const root = shape[SHAPE_SLIDE][SLIDE_DOCUMENT].root;
  const candidates: Array<{ node: XmlElement; parent: XmlElement }> = [];
  const walk = (node: XmlElement, visit: (node: XmlElement, parent: XmlElement) => void): void => {
    for (const child of node.children) {
      if (child.kind !== 'element') continue;
      visit(child, node);
      walk(child, visit);
    }
  };
  const spid = String(shape[SHAPE_SNAPSHOT].id);
  walk(root, (node, parent) => {
    if (node.name.namespaceURI !== NS.pml || node.name.localName !== 'seq') return;
    const cTn = firstChildElement(node, NAME_C_TN);
    if (
      cTn !== null &&
      getAttrValue(cTn, ATTR_NODE_TYPE) === (from === 'background' ? 'mainSeq' : 'interactiveSeq')
    ) {
      candidates.push({ node, parent });
    }
  });
  const timing = firstChildElement(root, qname('p', 'timing', NS.pml));
  const list = timing && firstChildElement(timing, qname('p', 'tnLst', NS.pml));
  const parallel = list && firstChildElement(list, qname('p', 'par', NS.pml));
  const rootTime = parallel && firstChildElement(parallel, NAME_C_TN);
  const rootChildren = rootTime && firstChildElement(rootTime, NAME_CHILD_TN_LST);
  for (const { node, parent } of candidates) {
    if (
      parent !== rootChildren ||
      rootTime === null ||
      getAttrValue(rootTime, ATTR_NODE_TYPE) !== 'tmRoot'
    )
      continue;
    const ids: string[] = [];
    walk(node, (child) => {
      if (child.name.namespaceURI === NS.pml && child.name.localName === 'cTn') {
        const id = getAttrValue(child, ATTR_ID);
        if (id !== null) ids.push(id);
      }
    });
    if (
      ids.length !== 5 ||
      new Set(ids).size !== 5 ||
      !sameTimingTree(node, nativeMediaSequence(from, ids, spid))
    )
      continue;
    const idSet = new Set(ids);
    let shared = false;
    const inspectExternal = (element: XmlElement): void => {
      if (element === node) return;
      if (element.name.namespaceURI === NS.pml) {
        if (element.name.localName === 'tn' && idSet.has(getAttrValue(element, ATTR_VAL) ?? ''))
          shared = true;
        if (
          element.name.localName === 'cTn' &&
          (idSet.has(getAttrValue(element, ATTR_ID) ?? '') ||
            (from === 'interactive' && getAttrValue(element, ATTR_NODE_TYPE) === 'mainSeq'))
        )
          shared = true;
      }
      for (const child of element.children) if (child.kind === 'element') inspectExternal(child);
    };
    inspectExternal(root);
    if (shared) break;
    const replacement = nativeMediaSequence(
      from === 'background' ? 'interactive' : 'background',
      ids,
      spid,
    );
    replacement.prefixDecls = new Map([['p', NS.pml]]);
    return () => {
      parent.children[parent.children.indexOf(node)] = replacement;
    };
  }
  throw new Error('setShapeMediaPlayback: dedicated media command timing is unsupported or shared');
};

// Only the simple, event-free form can be reduced to an effective slide
// start. A click or multiple-condition ancestor stays event-driven and is
// deliberately not guessed as autoplay.
const effectiveMediaStart = (
  cTns: ReadonlyArray<XmlElement>,
  requireLastCondition = false,
): MediaStart => {
  let delayMs = 0;
  for (const [index, cTn] of cTns.entries()) {
    if (getAttrValue(cTn, ATTR_MASTER_REL) !== null) return null;
    const list = firstChildElement(cTn, NAME_ST_COND_LST);
    if (list === null) {
      if (requireLastCondition && index === cTns.length - 1) return null;
      continue;
    }
    const conditions = list.children.filter(
      (child): child is XmlElement =>
        child.kind === 'element' &&
        child.name.namespaceURI === NS.pml &&
        child.name.localName === 'cond',
    );
    if (conditions.length !== 1) return null;
    const condition = conditions[0]!;
    if (
      condition.children.some(
        (child) =>
          child.kind === 'element' &&
          child.name.namespaceURI === NS.pml &&
          ['tn', 'rtn', 'tgtEl'].includes(child.name.localName),
      )
    )
      return null;
    if (getAttrValue(condition, ATTR_EVT) !== null) return { automatic: false, delayMs: 0 };
    const raw = getAttrValue(condition, ATTR_DELAY);
    if (raw === 'indefinite') return { automatic: false, delayMs: 0 };
    const value = raw === null ? 0 : Number(raw);
    if (!Number.isSafeInteger(value) || value < 0) return null;
    delayMs += value;
    if (!Number.isSafeInteger(delayMs)) return null;
  }
  return { automatic: true, delayMs };
};

/**
 * Reads how the shape's clip plays, or `null` when the shape carries no media
 * time node — which is what a picture that is not a clip looks like, and also
 * what a clip pasted in without its node looks like (it shows no controls).
 */
export const getShapeMediaPlayback = (shape: SlideShapeData): MediaPlayback | null => {
  const found = mediaNodeOf(shape);
  if (found === null) return null;
  const { node, media } = found;
  const cTn = firstChildElement(media, NAME_C_TN);
  const commandTiming = mediaCommandTiming(shape);
  const ownStart = cTn === null ? null : effectiveMediaStart([cTn], true);
  const parentStart = effectiveMediaStart(found.ancestors);
  const delayMs = (parentStart?.delayMs ?? 0) + (ownStart?.delayMs ?? 0);
  const autoplay =
    commandTiming === 'background' ||
    (cTn !== null &&
      !found.duplicateTarget &&
      !found.hasDependentTimingAncestor &&
      parentStart?.automatic === true &&
      ownStart?.automatic === true &&
      Number.isSafeInteger(delayMs));
  const playback: MediaPlayback = {
    // A delay is measured after the condition's trigger. An `evt` condition
    // (for example `onClick`) therefore remains event-triggered even when its
    // delay is zero; only an event-free condition starts with the slide.
    autoplay,
    loop: cTn !== null && getAttrValue(cTn, ATTR_REPEAT_COUNT) === 'indefinite',
    volume: percentFraction(getAttrValue(media, ATTR_VOL), 0.5),
    muted: xsdBoolean(getAttrValue(media, ATTR_MUTE), false),
    fullScreen: xsdBoolean(getAttrValue(node, ATTR_FULL_SCRN), false),
    hideWhenStopped: !xsdBoolean(getAttrValue(media, ATTR_SHOW_WHEN_STOPPED), true),
  };
  const rawSlideCount = getAttrValue(media, ATTR_NUM_SLD);
  const parsedSlideCount = rawSlideCount === null ? null : Number(rawSlideCount);
  const slideCount =
    parsedSlideCount !== null &&
    Number.isSafeInteger(parsedSlideCount) &&
    parsedSlideCount >= 0 &&
    parsedSlideCount <= MAX_MEDIA_SLIDE_COUNT
      ? parsedSlideCount
      : undefined;
  const rewindAfterPlaying = cTn !== null && getAttrValue(cTn, ATTR_FILL) === 'remove';
  const optionalPlayback = {
    ...(slideCount === undefined ? {} : { slideCount }),
    ...(rewindAfterPlaying ? { rewindAfterPlaying: true } : {}),
    ...mediaAdjustments(shape),
  };
  // Zero is the ordinary immediate-start form and remains absent to preserve
  // the existing result shape. Only a finite, event-free start can carry this
  // user-facing delay.
  if (autoplay && delayMs > 0) {
    return { ...playback, delayMs, ...optionalPlayback };
  }
  return { ...playback, ...optionalPlayback };
};

/**
 * Updates how the shape's clip plays. Omitted properties keep their current
 * value. Throws when the shape has no media time node — `addSlideMedia` writes
 * one, and a picture that is not a clip never plays.
 *
 * `delayMs` is a nonnegative safe integer and requires automatic playback.
 * Setting `autoplay: true` alone preserves an existing automatic delay;
 * switching from event playback starts immediately unless `delayMs` is supplied.
 *
 * `fullScreen` is a video attribute (`CT_TLMediaNodeVideo`); asking for it on
 * an audio clip throws rather than writing an attribute the schema rejects.
 */
export const setShapeMediaPlayback = (
  shape: SlideShapeData,
  options: Partial<MediaPlayback>,
): void => {
  const found = mediaNodeOf(shape);
  if (found === null) throw new Error('setShapeMediaPlayback: the shape has no media time node');
  const { node, media } = found;
  const commandTiming = mediaCommandTiming(shape);
  const convertingCommand =
    (commandTiming === 'background' && options.autoplay === false) ||
    (commandTiming === 'interactive' && options.autoplay === true);
  const delayMs = options.delayMs;
  if (delayMs !== undefined && (!Number.isSafeInteger(delayMs) || delayMs < 0)) {
    throw new Error('setShapeMediaPlayback: delayMs must be a nonnegative safe integer');
  }
  if (delayMs !== undefined && options.autoplay === false) {
    throw new Error('setShapeMediaPlayback: delayMs cannot be used with autoplay:false');
  }
  const cTn = firstChildElement(media, NAME_C_TN);
  const stCondLst = cTn && firstChildElement(cTn, NAME_ST_COND_LST);
  const currentStart = stCondLst && firstChildElement(stCondLst, NAME_COND);
  const currentAutoplay =
    currentStart !== null &&
    getAttrValue(currentStart, ATTR_EVT) === null &&
    getAttrValue(currentStart, ATTR_DELAY) !== 'indefinite';
  if (
    found.ancestors.length > 1 &&
    cTn !== null &&
    stCondLst !== null &&
    (options.autoplay !== undefined || delayMs !== undefined) &&
    effectiveMediaStart([cTn], true) === null
  ) {
    throw new Error('setShapeMediaPlayback: media start conditions are unsupported');
  }
  if (found.duplicateTarget) {
    throw new Error('setShapeMediaPlayback: multiple media nodes target this shape');
  }
  if (
    cTn === null &&
    (options.autoplay !== undefined ||
      options.delayMs !== undefined ||
      options.loop !== undefined ||
      options.rewindAfterPlaying !== undefined)
  ) {
    throw new Error('setShapeMediaPlayback: media timing node has no cTn');
  }
  if (
    commandTiming === 'other' &&
    (options.autoplay !== undefined || options.delayMs !== undefined)
  ) {
    throw new Error('setShapeMediaPlayback: media command timing is unsupported');
  }
  if (
    (commandTiming === 'background' || commandTiming === 'interactive') &&
    delayMs !== undefined
  ) {
    throw new Error('setShapeMediaPlayback: background media timing cannot be changed');
  }
  if (
    (options.autoplay !== undefined || delayMs !== undefined) &&
    cTn !== null &&
    (getAttrValue(cTn, ATTR_MASTER_REL) !== null ||
      (currentStart !== null &&
        currentStart.children.some(
          (child) =>
            child.kind === 'element' &&
            child.name.namespaceURI === NS.pml &&
            ['tn', 'rtn', 'tgtEl'].includes(child.name.localName),
        )))
  ) {
    throw new Error('setShapeMediaPlayback: media start depends on another timing node');
  }
  let ownDelayMs = delayMs;
  if (options.autoplay !== undefined || delayMs !== undefined) {
    if (found.hasDependentTimingAncestor) {
      throw new Error('setShapeMediaPlayback: media dependent timing is unsupported');
    }
    const parent = effectiveMediaStart(found.ancestors);
    if (parent === null) {
      throw new Error('setShapeMediaPlayback: nested media has unsupported start conditions');
    }
    if (!parent.automatic && options.autoplay !== false) {
      throw new Error('setShapeMediaPlayback: media is event-triggered');
    }
    if (delayMs !== undefined) {
      if (!parent.automatic || delayMs < parent.delayMs) {
        throw new Error(
          'setShapeMediaPlayback: delayMs cannot be represented with nested start conditions',
        );
      }
      ownDelayMs = delayMs - parent.delayMs;
    }
  }
  if (delayMs !== undefined && options.autoplay === undefined && !currentAutoplay) {
    throw new Error('setShapeMediaPlayback: delayMs requires autoplay');
  }
  if (
    options.volume !== undefined &&
    (!Number.isFinite(options.volume) || options.volume < 0 || options.volume > 1)
  ) {
    throw new Error('setShapeMediaPlayback: volume must be between 0 and 1');
  }
  if (options.fullScreen !== undefined && node.name.localName !== 'video') {
    throw new Error('setShapeMediaPlayback: fullScreen applies to video only');
  }
  if (
    options.slideCount !== undefined &&
    (!Number.isSafeInteger(options.slideCount) ||
      options.slideCount < 0 ||
      options.slideCount > MAX_MEDIA_SLIDE_COUNT)
  ) {
    throw new Error('setShapeMediaPlayback: slideCount must be an unsigned 32-bit integer');
  }
  if (options.trim !== undefined) validateAdjustment(options.trim, 'trim');
  if (options.fade !== undefined) validateAdjustment(options.fade, 'fade');
  if (options.bookmarks !== undefined) validateBookmarks(options.bookmarks);
  const p14Media =
    options.trim !== undefined || options.fade !== undefined || options.bookmarks !== undefined
      ? p14MediaOf(shape)
      : null;
  if (
    (options.trim !== undefined || options.fade !== undefined || options.bookmarks !== undefined) &&
    p14Media === null
  ) {
    throw new Error('setShapeMediaPlayback: trim, fade, and bookmarks require embedded media');
  }
  if (options.bookmarks !== undefined) validateBookmarkReferences(shape, options.bookmarks);

  const applyCommandTiming = convertingCommand
    ? prepareDedicatedMediaTiming(
        shape,
        commandTiming === 'background' ? 'background' : 'interactive',
      )
    : null;
  applyCommandTiming?.();

  if (options.volume !== undefined) {
    setOrRemove(media, ATTR_VOL, String(Math.round(options.volume * 100000)));
  }
  if (options.muted !== undefined) setOrRemove(media, ATTR_MUTE, options.muted ? '1' : '0');
  if (options.hideWhenStopped !== undefined) {
    setOrRemove(media, ATTR_SHOW_WHEN_STOPPED, options.hideWhenStopped ? '0' : '1');
  }
  if (options.slideCount !== undefined) {
    // One is the schema default; omit it so newly-authored files stay as
    // compact as PowerPoint's ordinary single-slide playback form.
    setOrRemove(media, ATTR_NUM_SLD, options.slideCount === 1 ? null : String(options.slideCount));
  }
  if (options.fullScreen !== undefined) {
    setOrRemove(node, ATTR_FULL_SCRN, options.fullScreen ? '1' : '0');
  }
  if (p14Media !== null) setMediaAdjustments(p14Media, options.trim, options.fade);
  if (p14Media !== null && options.bookmarks !== undefined)
    setMediaBookmarks(p14Media, options.bookmarks);

  if (cTn !== null) {
    // Mac PowerPoint stores Rewind After Playing as remove (on) or hold (off).
    if (options.rewindAfterPlaying !== undefined) {
      setOrRemove(cTn, ATTR_FILL, options.rewindAfterPlaying ? 'remove' : 'hold');
    }
    if (options.loop !== undefined) {
      setOrRemove(cTn, ATTR_REPEAT_COUNT, options.loop ? 'indefinite' : null);
    }
    if (
      !convertingCommand &&
      commandTiming !== 'background' &&
      commandTiming !== 'interactive' &&
      (options.autoplay !== undefined || delayMs !== undefined)
    ) {
      // CT_TLCommonTimeNodeData is a sequence: `<p:stCondLst>` comes first.
      let stCondLst = firstChildElement(cTn, NAME_ST_COND_LST);
      if (stCondLst === null) {
        stCondLst = elem(NAME_ST_COND_LST);
        cTn.children.unshift(stCondLst);
      }
      const nextDelay =
        options.autoplay === false
          ? 'indefinite'
          : ownDelayMs !== undefined
            ? String(ownDelayMs)
            : currentAutoplay
              ? (getAttrValue(currentStart!, ATTR_DELAY) ?? '0')
              : '0';
      if (currentAutoplay && currentStart !== null && options.autoplay !== false) {
        setOrRemove(currentStart, ATTR_DELAY, nextDelay);
      } else {
        stCondLst.children = [
          elem(NAME_COND, {
            attrs: [attr(ATTR_DELAY, nextDelay)],
          }),
        ];
      }
    }
  }
  commitSlideData(shape[SHAPE_SLIDE]);
  refreshSlideData(shape[SHAPE_SLIDE]);
};

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
import { addMediaTimingNode, findMediaTimingNode } from './_media-timing.ts';
import {
  NS,
  type XmlElement,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  qname,
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

/**
 * How a clip plays in the slide show — the attributes of its
 * `<p:cMediaNode>` and the start condition of its time node.
 *
 * Trimming (`p14:trim`) is not part of this: PowerPoint stores it in a 2010
 * extension rather than in the core schema, and a reader that does not know
 * the extension plays the whole clip.
 */
export interface MediaPlayback {
  /** Starts with the slide instead of waiting for a click. */
  readonly autoplay: boolean;
  /** Plays again from the beginning until the slide moves on. */
  readonly loop: boolean;
  /** Playback volume, 0–1. PowerPoint's own default is 0.8. */
  readonly volume: number;
  readonly muted: boolean;
  /** Video only: plays filling the screen. */
  readonly fullScreen: boolean;
  /** Hides the clip once it has played (`showWhenStopped="0"`). */
  readonly hideWhenStopped: boolean;
}

const NAME_C_MEDIA_NODE = qname('p', 'cMediaNode', NS.pml);
const NAME_C_TN = qname('p', 'cTn', NS.pml);
const NAME_ST_COND_LST = qname('p', 'stCondLst', NS.pml);
const NAME_COND = qname('p', 'cond', NS.pml);
const ATTR_VOL = qname('', 'vol', '');
const ATTR_MUTE = qname('', 'mute', '');
const ATTR_FULL_SCRN = qname('', 'fullScrn', '');
const ATTR_SHOW_WHEN_STOPPED = qname('', 'showWhenStopped', '');
const ATTR_REPEAT_COUNT = qname('', 'repeatCount', '');
const ATTR_DELAY = qname('', 'delay', '');

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
  el.attrs = el.attrs.filter((a) => a.name.localName !== name.localName);
  if (value !== null) el.attrs.push(attr(name, value));
};

const mediaNodeOf = (shape: SlideShapeData): { node: XmlElement; media: XmlElement } | null => {
  const node = findMediaTimingNode(shape[SHAPE_SLIDE], shape[SHAPE_SNAPSHOT].id);
  const media = node === null ? null : firstChildElement(node, NAME_C_MEDIA_NODE);
  return node !== null && media !== null ? { node, media } : null;
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
  const stCondLst = cTn && firstChildElement(cTn, NAME_ST_COND_LST);
  const start = stCondLst && firstChildElement(stCondLst, NAME_COND);
  return {
    // `indefinite` is "wait to be started"; every other delay starts on its own.
    autoplay: start !== null && getAttrValue(start, ATTR_DELAY) !== 'indefinite',
    loop: cTn !== null && getAttrValue(cTn, ATTR_REPEAT_COUNT) === 'indefinite',
    volume: percentFraction(getAttrValue(media, ATTR_VOL), 0.5),
    muted: xsdBoolean(getAttrValue(media, ATTR_MUTE), false),
    fullScreen: xsdBoolean(getAttrValue(node, ATTR_FULL_SCRN), false),
    hideWhenStopped: !xsdBoolean(getAttrValue(media, ATTR_SHOW_WHEN_STOPPED), true),
  };
};

/**
 * Updates how the shape's clip plays. Omitted properties keep their current
 * value. Throws when the shape has no media time node — `addSlideMedia` writes
 * one, and a picture that is not a clip never plays.
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
  if (options.volume !== undefined && (options.volume < 0 || options.volume > 1)) {
    throw new Error('setShapeMediaPlayback: volume must be between 0 and 1');
  }
  if (options.fullScreen !== undefined && node.name.localName !== 'video') {
    throw new Error('setShapeMediaPlayback: fullScreen applies to video only');
  }

  if (options.volume !== undefined) {
    setOrRemove(media, ATTR_VOL, String(Math.round(options.volume * 100000)));
  }
  if (options.muted !== undefined) setOrRemove(media, ATTR_MUTE, options.muted ? '1' : '0');
  if (options.hideWhenStopped !== undefined) {
    setOrRemove(media, ATTR_SHOW_WHEN_STOPPED, options.hideWhenStopped ? '0' : '1');
  }
  if (options.fullScreen !== undefined) {
    setOrRemove(node, ATTR_FULL_SCRN, options.fullScreen ? '1' : '0');
  }

  const cTn = firstChildElement(media, NAME_C_TN);
  if (cTn !== null) {
    if (options.loop !== undefined) {
      setOrRemove(cTn, ATTR_REPEAT_COUNT, options.loop ? 'indefinite' : null);
    }
    if (options.autoplay !== undefined) {
      // CT_TLCommonTimeNodeData is a sequence: `<p:stCondLst>` comes first.
      let stCondLst = firstChildElement(cTn, NAME_ST_COND_LST);
      if (stCondLst === null) {
        stCondLst = elem(NAME_ST_COND_LST);
        cTn.children.unshift(stCondLst);
      }
      stCondLst.children = [
        elem(NAME_COND, {
          attrs: [attr(ATTR_DELAY, options.autoplay ? '0' : 'indefinite')],
        }),
      ];
    }
  }
  commitSlideData(shape[SHAPE_SLIDE]);
  refreshSlideData(shape[SHAPE_SLIDE]);
};

// Video / audio container detection from raw bytes, plus the content type each
// container is registered under. Counterpart of `image-format.ts`.
//
// A format token doubles as the part's file extension. Detection is by
// container signature only — we never look at the codecs inside, so whether
// the clip actually plays depends on the codecs the viewer has installed.

export type VideoFormat = 'mp4' | 'm4v' | 'mov' | 'webm' | 'avi' | 'wmv';
export type AudioFormat = 'mp3' | 'wav' | 'm4a' | 'ogg' | 'wma';

const VIDEO_CONTENT_TYPES: Readonly<Record<VideoFormat, string>> = {
  mp4: 'video/mp4',
  // PowerPoint registers .m4v under the mp4 type rather than video/x-m4v.
  m4v: 'video/mp4',
  mov: 'video/quicktime',
  webm: 'video/webm',
  avi: 'video/x-msvideo',
  wmv: 'video/x-ms-wmv',
};

const AUDIO_CONTENT_TYPES: Readonly<Record<AudioFormat, string>> = {
  mp3: 'audio/mpeg',
  // PowerPoint writes the legacy `x-wav` token, not RFC 2361's `audio/wav`.
  wav: 'audio/x-wav',
  m4a: 'audio/mp4',
  ogg: 'audio/ogg',
  wma: 'audio/x-ms-wma',
};

export const contentTypeForVideoFormat = (format: VideoFormat): string =>
  VIDEO_CONTENT_TYPES[format];
export const contentTypeForAudioFormat = (format: AudioFormat): string =>
  AUDIO_CONTENT_TYPES[format];

const hasBytesAt = (bytes: Uint8Array, at: number, signature: ReadonlyArray<number>): boolean => {
  if (bytes.length < at + signature.length) return false;
  for (let i = 0; i < signature.length; i++) {
    if (bytes[at + i] !== signature[i]) return false;
  }
  return true;
};

const hasAsciiAt = (bytes: Uint8Array, at: number, text: string): boolean =>
  hasBytesAt(
    bytes,
    at,
    Array.from(text, (ch) => ch.charCodeAt(0)),
  );

// ISO base media (mp4 / mov / m4a / m4v): a `ftyp` box first, major brand at 8.
const isoBrand = (bytes: Uint8Array): string | null =>
  hasAsciiAt(bytes, 4, 'ftyp') && bytes.length >= 12
    ? String.fromCharCode(bytes[8]!, bytes[9]!, bytes[10]!, bytes[11]!)
    : null;

const isRiff = (bytes: Uint8Array, form: string): boolean =>
  hasAsciiAt(bytes, 0, 'RIFF') && hasAsciiAt(bytes, 8, form);

// ASF header object GUID — the container of both .wmv and .wma. The two are
// byte-identical at this level, so the caller's video/audio intent decides.
const ASF_HEADER = [0x30, 0x26, 0xb2, 0x75, 0x8e, 0x66, 0xcf, 0x11];

const EBML_HEADER = [0x1a, 0x45, 0xdf, 0xa3];
// The EBML DocType string sits within the first few dozen bytes of the header.
const EBML_DOCTYPE_WINDOW = 64;
const isWebm = (bytes: Uint8Array): boolean => {
  if (!hasBytesAt(bytes, 0, EBML_HEADER)) return false;
  // Matroska (.mkv) shares the EBML magic; only the DocType tells them apart.
  const limit = Math.min(bytes.length, EBML_DOCTYPE_WINDOW) - 4;
  for (let i = 4; i <= limit; i++) {
    if (hasAsciiAt(bytes, i, 'webm')) return true;
  }
  return false;
};

/** Detects a video container, or `null` when no known signature matches. */
export const detectVideoFormat = (bytes: Uint8Array): VideoFormat | null => {
  const brand = isoBrand(bytes);
  if (brand !== null) {
    if (brand === 'qt  ') return 'mov';
    if (brand === 'M4V ') return 'm4v';
    return 'mp4';
  }
  if (isWebm(bytes)) return 'webm';
  if (isRiff(bytes, 'AVI ')) return 'avi';
  if (hasBytesAt(bytes, 0, ASF_HEADER)) return 'wmv';
  return null;
};

/** Detects an audio container, or `null` when no known signature matches. */
export const detectAudioFormat = (bytes: Uint8Array): AudioFormat | null => {
  if (isoBrand(bytes) !== null) return 'm4a';
  if (isRiff(bytes, 'WAVE')) return 'wav';
  if (hasAsciiAt(bytes, 0, 'OggS')) return 'ogg';
  if (hasBytesAt(bytes, 0, ASF_HEADER)) return 'wma';
  // MP3: an ID3v2 tag, or a bare MPEG audio frame (11 set sync bits).
  if (hasAsciiAt(bytes, 0, 'ID3')) return 'mp3';
  if (bytes.length >= 2 && bytes[0] === 0xff && (bytes[1]! & 0xe0) === 0xe0) return 'mp3';
  return null;
};

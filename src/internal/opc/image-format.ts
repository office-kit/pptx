// Image format detection from raw bytes. Used by the picture-replacement
// path to pick the right content type and file extension without forcing the
// caller to spell them out.
//
// Detection is by magic bytes — sufficient for the formats PowerPoint
// accepts. We do not run a full validator (that's the image library's job).

export type ImageFormat = 'png' | 'jpeg' | 'gif' | 'bmp' | 'tiff' | 'webp' | 'svg';

const startsWith = (bytes: Uint8Array, signature: ReadonlyArray<number>): boolean => {
  if (bytes.length < signature.length) return false;
  for (let i = 0; i < signature.length; i++) {
    if (bytes[i] !== signature[i]) return false;
  }
  return true;
};

const decoder = new TextDecoder('utf-8', { fatal: false });

/**
 * Detects the image format from raw bytes. Returns `null` if no known
 * signature matches — callers should treat that as a hard error and surface
 * a message asking for a recognized format.
 */
export const detectImageFormat = (bytes: Uint8Array): ImageFormat | null => {
  if (startsWith(bytes, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) return 'png';
  if (startsWith(bytes, [0xff, 0xd8, 0xff])) return 'jpeg';
  if (startsWith(bytes, [0x47, 0x49, 0x46, 0x38])) return 'gif';
  if (startsWith(bytes, [0x42, 0x4d])) return 'bmp';
  if (startsWith(bytes, [0x49, 0x49, 0x2a, 0x00])) return 'tiff';
  if (startsWith(bytes, [0x4d, 0x4d, 0x00, 0x2a])) return 'tiff';
  if (
    bytes.length >= 12 &&
    startsWith(bytes, [0x52, 0x49, 0x46, 0x46]) &&
    bytes[8] === 0x57 &&
    bytes[9] === 0x45 &&
    bytes[10] === 0x42 &&
    bytes[11] === 0x50
  ) {
    return 'webp';
  }
  // SVG is textual — sniff up to the first 1KB.
  const head = decoder.decode(bytes.subarray(0, Math.min(bytes.length, 1024)));
  if (/<svg[\s>]/.test(head)) return 'svg';
  return null;
};

/** Natural pixel dimensions of an image. */
export interface ImagePixelSize {
  readonly width: number;
  readonly height: number;
}

const readUint16Be = (bytes: Uint8Array, at: number): number => (bytes[at]! << 8) | bytes[at + 1]!;

const readUint32Be = (bytes: Uint8Array, at: number): number =>
  // `>>> 0` keeps the result an unsigned 32-bit int (a 4-byte PNG dimension
  // with the high bit set would otherwise read as negative).
  ((bytes[at]! << 24) | (bytes[at + 1]! << 16) | (bytes[at + 2]! << 8) | bytes[at + 3]!) >>> 0;

// PNG: the IHDR chunk is the first chunk and always at a fixed offset —
// 8-byte signature, 4-byte length, 4-byte "IHDR" tag, then width / height
// as big-endian uint32 (PNG spec §11.2.2).
const pngSize = (bytes: Uint8Array): ImagePixelSize | null => {
  if (bytes.length < 24) return null;
  const width = readUint32Be(bytes, 16);
  const height = readUint32Be(bytes, 20);
  if (width <= 0 || height <= 0) return null;
  return { width, height };
};

// JPEG: walk the marker segments until a Start-Of-Frame marker, whose
// payload carries the sample dimensions. SOF markers are 0xC0..0xCF except
// the non-frame markers 0xC4 (DHT), 0xC8 (JPG), 0xCC (DAC).
const JPEG_SOF_EXCLUDED = new Set([0xc4, 0xc8, 0xcc]);
const jpegSize = (bytes: Uint8Array): ImagePixelSize | null => {
  // Skip the SOI (0xFFD8); then each segment is 0xFF, marker, 2-byte length.
  let offset = 2;
  while (offset + 9 < bytes.length) {
    if (bytes[offset] !== 0xff) {
      offset++;
      continue;
    }
    const marker = bytes[offset + 1]!;
    // Padding / standalone markers (RSTn, SOI, EOI, TEM) carry no length.
    if (marker === 0xff || (marker >= 0xd0 && marker <= 0xd9) || marker === 0x01) {
      offset += 2;
      continue;
    }
    const segmentLength = readUint16Be(bytes, offset + 2);
    if (segmentLength < 2) return null;
    if (marker >= 0xc0 && marker <= 0xcf && !JPEG_SOF_EXCLUDED.has(marker)) {
      // SOF payload: 1-byte precision, 2-byte height, 2-byte width.
      const height = readUint16Be(bytes, offset + 5);
      const width = readUint16Be(bytes, offset + 7);
      if (width <= 0 || height <= 0) return null;
      return { width, height };
    }
    offset += 2 + segmentLength;
  }
  return null;
};

/**
 * Reads an image's natural pixel dimensions from its header. Supports PNG
 * and JPEG — the two formats whose headers carry dimensions cheaply and
 * unambiguously. Returns `null` for every other format (and for truncated
 * / malformed headers), letting callers fall back rather than fail: an
 * aspect-ratio-preserving placement that can't measure the image just
 * stretches it as before.
 */
export const readImagePixelSize = (bytes: Uint8Array): ImagePixelSize | null => {
  const format = detectImageFormat(bytes);
  if (format === 'png') return pngSize(bytes);
  if (format === 'jpeg') return jpegSize(bytes);
  return null;
};

/**
 * Returns the conventional file-extension token (no leading dot) for the
 * given format. `jpeg` maps to `jpg` because that's what PowerPoint emits.
 */
export const extensionForFormat = (format: ImageFormat): string => {
  switch (format) {
    case 'jpeg':
      return 'jpg';
    case 'png':
      return 'png';
    case 'gif':
      return 'gif';
    case 'bmp':
      return 'bmp';
    case 'tiff':
      return 'tiff';
    case 'webp':
      return 'webp';
    case 'svg':
      return 'svg';
  }
};

/** Returns the IANA media type for the given image format. */
export const contentTypeForFormat = (format: ImageFormat): string => {
  switch (format) {
    case 'png':
      return 'image/png';
    case 'jpeg':
      return 'image/jpeg';
    case 'gif':
      return 'image/gif';
    case 'bmp':
      return 'image/bmp';
    case 'tiff':
      return 'image/tiff';
    case 'webp':
      return 'image/webp';
    case 'svg':
      return 'image/svg+xml';
  }
};

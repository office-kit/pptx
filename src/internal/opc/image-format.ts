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

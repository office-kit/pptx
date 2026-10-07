import { inflateSync } from 'node:zlib';
import {
  getSlideBackground,
  getSlideBackgroundImageBytes,
  getSlideLayout,
  getSlideLayoutBackground,
  getSlideLayoutBackgroundImageBytes,
  getSlideMasterBackgroundImageBytes,
} from '../../../../dist/index.js';
import {
  TEXTURES,
  TEXTURE_SIZE,
  texturePixels,
} from '../../../../packages/editor/src/core/textures.ts';

const PNG_HEADER_LENGTH = 8;
const RGB = 3;
const RGBA = 4;

// The editor's texture PNGs are unfiltered 8-bit RGB, so the inflated IDAT is
// each row's filter byte followed by its pixels.
function rgbRows(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const data = [];
  for (let offset = PNG_HEADER_LENGTH; offset + 8 <= bytes.length; ) {
    const length = view.getUint32(offset);
    const type = new TextDecoder('latin1').decode(bytes.subarray(offset + 4, offset + 8));
    if (type === 'IDAT') data.push(bytes.subarray(offset + 8, offset + 8 + length));
    offset += 12 + length;
  }
  return data.length ? inflateSync(Buffer.concat(data)) : null;
}

/** Which generated gallery texture the picture is, by its pixels; null for any other picture. */
export function textureIdOf(bytes) {
  const rows = bytes && rgbRows(bytes);
  if (!rows || rows.length !== TEXTURE_SIZE * (1 + TEXTURE_SIZE * RGB)) return null;
  return (
    TEXTURES.find(({ id }) => {
      const pixels = texturePixels(id);
      for (let y = 0; y < TEXTURE_SIZE; y++) {
        const row = y * (1 + TEXTURE_SIZE * RGB) + 1;
        for (let x = 0; x < TEXTURE_SIZE; x++) {
          for (let channel = 0; channel < RGB; channel++) {
            if (rows[row + x * RGB + channel] !== pixels[(y * TEXTURE_SIZE + x) * RGBA + channel])
              return false;
          }
        }
      }
      return true;
    })?.id ?? null
  );
}

/** The visible background picture, following slide → layout → master inheritance. */
export function slideBackgroundImageBytes(pres, slide) {
  if (getSlideBackground(slide).kind !== 'inherit') return getSlideBackgroundImageBytes(slide);
  const layout = getSlideLayout(slide);
  if (!layout) return null;
  if (getSlideLayoutBackground(layout).kind !== 'inherit')
    return getSlideLayoutBackgroundImageBytes(pres, layout);
  return getSlideMasterBackgroundImageBytes(pres, layout);
}

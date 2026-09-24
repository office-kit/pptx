// CSS for the parts of `<a:bodyPr>` that HTML expresses directly: vertical
// writing and multi-column text bodies. The renderer's `<foreignObject>` path
// and any HTML editing surface have to agree on these, so they live here
// rather than inside the SVG emitter.

import type { getShapeTextColumns, getShapeTextDirection } from '@office-kit/pptx';

/** Mac PowerPoint uses zero spacing when `spcCol` is absent. */
const DEFAULT_COLUMN_GAP_PX = 0;

const EMU_PER_PX = 9525;

export interface VerticalTextStyle {
  /** Declarations to add to the text container, without a trailing `;`. */
  readonly declarations: string;
  /** Transform function to compose with whatever the caller already applies,
   *  or `''`. Kept apart so a caller that rotates the box for shape rotation
   *  can append it instead of losing one of the two. */
  readonly transform: string;
}

const NONE: VerticalTextStyle = { declarations: '', transform: '' };

/**
 * `<a:bodyPr vert="…"/>` as CSS. `writing-mode` is the right primitive for
 * the East-Asian and Mongolian cases, and `text-orientation` for the WordArt
 * ones, which stack upright glyphs. `vert270` reads bottom-to-top, which no
 * writing-mode expresses on its own — hence the half turn.
 */
export function verticalTextStyle(
  vert: ReturnType<typeof getShapeTextDirection>,
): VerticalTextStyle {
  switch (vert) {
    case 'vert':
    case 'eaVert':
      return { declarations: 'writing-mode:vertical-rl', transform: '' };
    case 'mongolianVert':
      return { declarations: 'writing-mode:vertical-lr', transform: '' };
    case 'vert270':
      return { declarations: 'writing-mode:vertical-lr', transform: 'rotate(180deg)' };
    case 'wordArtVert':
      return { declarations: 'writing-mode:vertical-rl;text-orientation:upright', transform: '' };
    case 'wordArtVertRtl':
      return {
        declarations: 'writing-mode:vertical-rl;text-orientation:upright;direction:rtl',
        transform: '',
      };
    case null:
      return NONE;
  }
}

/**
 * `<a:bodyPr numCol="N" spcCol="EMU"/>` as CSS declarations, without a
 * trailing `;`. Empty below two columns, which is the same as not splitting
 * the body.
 */
export function textColumnsStyle(cols: ReturnType<typeof getShapeTextColumns>): string {
  if (!cols || cols.count < 2) return '';
  const gapPx =
    cols.gapEmu !== undefined ? (cols.gapEmu / EMU_PER_PX).toFixed(2) : DEFAULT_COLUMN_GAP_PX;
  return `column-count:${cols.count};column-gap:${gapPx}px`;
}

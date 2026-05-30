// Node-only fontkit-backed TextMeasurer for the fidelity harness. NEVER
// imported by render-slide.ts (which must stay browser-safe) — only by
// render-ours.ts. It measures advance widths and vertical metrics from the
// same vendored TTFs that resvg rasterizes with and that LibreOffice renders
// ground truth with, so the engine's wrap/positioning math agrees with the
// painted pixels.

import * as fontkit from 'fontkit';
import { existsSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import {
  ARIAL,
  MONO,
  SANS,
  SERIF,
  TIMES,
  type FontSpec,
  type TextMeasurer,
} from '../src/lib/playground/text-layout.ts';

export const FONT_DIR = fileURLToPath(new URL('./fonts/', import.meta.url));

// Internal family name (what the emitter writes + resvg matches) → file prefix.
const FAMILY_TO_PREFIX: Record<string, string> = {
  [SANS]: 'Carlito',
  [SERIF]: 'Caladea',
  [ARIAL]: 'LiberationSans',
  [TIMES]: 'LiberationSerif',
  [MONO]: 'LiberationMono',
};

const STYLES = ['Regular', 'Bold', 'Italic', 'BoldItalic'] as const;

const facePath = (prefix: string, style: string): string => `${FONT_DIR}${prefix}-${style}.ttf`;

/** Absolute paths of every bundled face — passed to resvg's `fontFiles`. */
export const FONT_FILES: string[] = Object.values(FAMILY_TO_PREFIX).flatMap((prefix) =>
  STYLES.map((s) => facePath(prefix, s)),
);

// fontkit.create returns Font | FontCollection; our TTFs are single fonts.
const openFont = (path: string): fontkit.Font => {
  const font = fontkit.create(readFileSync(path));
  if (!('layout' in font)) {
    throw new Error(`Expected a single font in ${path}, got a collection`);
  }
  return font;
};

const styleSuffix = (spec: FontSpec): (typeof STYLES)[number] => {
  if (spec.bold && spec.italic) return 'BoldItalic';
  if (spec.bold) return 'Bold';
  if (spec.italic) return 'Italic';
  return 'Regular';
};

export const buildFontkitMeasurer = (): TextMeasurer => {
  const cache = new Map<string, fontkit.Font>();
  // Load + verify every face up front. The familyName assertion guarantees the
  // emit-side family === resvg match === the face we measure, which is the load-
  // bearing invariant for "x= is where glyphs land".
  for (const [family, prefix] of Object.entries(FAMILY_TO_PREFIX)) {
    for (const style of STYLES) {
      const path = facePath(prefix, style);
      if (!existsSync(path)) throw new Error(`Missing bundled font: ${path}`);
      const font = openFont(path);
      if (style === 'Regular' && font.familyName !== family) {
        throw new Error(
          `Font family mismatch: ${path} reports "${font.familyName}", expected "${family}". ` +
            `The substitution map and resvg matching rely on this name.`,
        );
      }
      cache.set(`${prefix}-${style}`, font);
    }
  }

  const pick = (spec: FontSpec): fontkit.Font => {
    const prefix = FAMILY_TO_PREFIX[spec.family] ?? 'Carlito';
    const style = styleSuffix(spec);
    return cache.get(`${prefix}-${style}`) ?? cache.get(`${prefix}-Regular`) ?? cache.get('Carlito-Regular')!;
  };

  return (text, spec) => {
    const font = pick(spec);
    const scale = spec.sizePx / font.unitsPerEm;
    const run = font.layout(text); // applies the font's kerning / GPOS
    const glyphCount = [...text].length;
    const tracking = glyphCount > 1 ? spec.letterSpacingPx * (glyphCount - 1) : 0;
    return {
      widthPx: run.advanceWidth * scale + tracking,
      ascentPx: font.ascent * scale,
      descentPx: Math.abs(font.descent) * scale,
      lineGapPx: font.lineGap * scale,
    };
  };
};

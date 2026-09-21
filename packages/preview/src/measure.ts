// Node-only fontkit-backed TextMeasurer. NEVER imported by render-slide.ts /
// text-layout.ts (which must stay browser-safe) — only by the node entry. It
// measures advance widths and vertical metrics from the same bundled TTFs that
// resvg rasterizes with (and that LibreOffice renders ground truth with), so
// the engine's wrap/positioning math agrees with the painted pixels.

import * as fontkit from 'fontkit';
import { existsSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import {
  ARIAL,
  AVG_GLYPH_W_RATIO,
  isCjk,
  MONO,
  SANS,
  SERIF,
  substituteFamily,
  TIMES,
  type FontSpec,
  type MeasureResult,
  type TextMeasurer,
} from './text-layout.ts';

export const FONT_DIR = fileURLToPath(new URL('../fonts/', import.meta.url));

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

const styleSuffix = (spec: { bold: boolean; italic: boolean }): (typeof STYLES)[number] => {
  if (spec.bold && spec.italic) return 'BoldItalic';
  if (spec.bold) return 'Bold';
  if (spec.italic) return 'Italic';
  return 'Regular';
};

/** A font the caller supplies for measurement, keyed by the AUTHORED family
 *  name (`<a:latin typeface>` / `<a:ea typeface>` — e.g. "游ゴシック",
 *  "Montserrat"), which routinely differs from the file's internal name. */
export interface RegisteredFont {
  readonly family: string;
  /** Path to a ttf/otf file, or its bytes. */
  readonly source: string | Uint8Array;
  readonly bold?: boolean;
  readonly italic?: boolean;
}

export interface FontkitMeasurerOptions {
  /** Fonts the deck actually uses. A run whose family matches a registered
   *  font measures with it; any registered font also serves as a fallback
   *  for glyphs the resolved font lacks (the bundled faces are Latin-only,
   *  so this is how CJK text gets real metrics). */
  readonly fonts?: ReadonlyArray<RegisteredFont>;
}

const openFontSource = (source: string | Uint8Array): fontkit.Font => {
  if (typeof source === 'string') return openFont(source);
  const font = fontkit.create(Buffer.from(source));
  if (!('layout' in font)) {
    throw new Error('Expected a single font, got a collection');
  }
  return font;
};

// Per-character width estimate for glyphs no available font covers — the same
// ratios as the heuristic measurer, so estimates and defaultMeasurer agree.
const estimateCharWidth = (ch: string, sizePx: number): number =>
  sizePx * (isCjk(ch.codePointAt(0) ?? 0) ? 1 : AVG_GLYPH_W_RATIO);

export const buildFontkitMeasurer = (options: FontkitMeasurerOptions = {}): TextMeasurer => {
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

  // Registered fonts, keyed by lowercased authored family + style. No
  // familyName assertion here: the registration key is the DECK's name for
  // the font, which legitimately differs from the file's internal name.
  const registered = new Map<string, fontkit.Font>();
  const fallbackFonts: fontkit.Font[] = [];
  for (const reg of options.fonts ?? []) {
    const font = openFontSource(reg.source);
    const style = styleSuffix({ bold: reg.bold ?? false, italic: reg.italic ?? false });
    registered.set(`${reg.family.toLowerCase()}-${style}`, font);
    fallbackFonts.push(font);
  }

  const pick = (spec: FontSpec): fontkit.Font => {
    const style = styleSuffix(spec);
    const regKey = spec.family.toLowerCase();
    const reg = registered.get(`${regKey}-${style}`) ?? registered.get(`${regKey}-Regular`);
    if (reg) return reg;
    // The render path pre-substitutes families (spec.family is a bundled
    // internal name); the audit path passes the authored name through, so
    // resolve it with the same map before keying the bundled faces.
    const prefix =
      FAMILY_TO_PREFIX[spec.family] ?? FAMILY_TO_PREFIX[substituteFamily(spec.family)]!;
    return (
      cache.get(`${prefix}-${style}`) ??
      cache.get(`${prefix}-Regular`) ??
      cache.get('Carlito-Regular')!
    );
  };

  // Splits `text` into maximal segments coverable by one font: the resolved
  // font first, then each registered font in registration order. Characters
  // no font covers fall back to the per-character estimate (approximate).
  const measureWithFallback = (
    text: string,
    spec: FontSpec,
    primary: fontkit.Font,
  ): MeasureResult => {
    const chain = [primary, ...fallbackFonts.filter((f) => f !== primary)];
    let widthPx = 0;
    let approximate = false;
    let metricsFont: fontkit.Font | null = null;
    let segment = '';
    let segmentFont: fontkit.Font | null = null;
    const flush = (): void => {
      if (segment === '' || segmentFont === null) return;
      const scale = spec.sizePx / segmentFont.unitsPerEm;
      widthPx += segmentFont.layout(segment).advanceWidth * scale;
      metricsFont ??= segmentFont;
      segment = '';
    };
    for (const ch of text) {
      const cp = ch.codePointAt(0) ?? 0;
      const font = chain.find((f) => f.hasGlyphForCodePoint(cp)) ?? null;
      if (font !== segmentFont) {
        flush();
        segmentFont = font;
      }
      if (font === null) {
        widthPx += estimateCharWidth(ch, spec.sizePx);
        approximate = true;
      } else {
        segment += ch;
      }
    }
    flush();
    const vm = verticalMetrics(metricsFont ?? primary);
    const vmScale = spec.sizePx / (metricsFont ?? primary).unitsPerEm;
    const glyphCount = [...text].length;
    const tracking = glyphCount > 1 ? spec.letterSpacingPx * (glyphCount - 1) : 0;
    return {
      widthPx: widthPx + tracking,
      ascentPx: vm.ascent * vmScale,
      descentPx: vm.descent * vmScale,
      lineGapPx: vm.lineGap * vmScale,
      ...(approximate ? { approximate } : {}),
    };
  };

  return (text, spec) => measureWithFallback(text, spec, pick(spec));
};

// Which vertical-metric set a renderer uses depends on the OS/2 fsSelection
// USE_TYPO_METRICS bit. When it's clear (the common case, incl. Carlito /
// Liberation), GDI / LibreOffice place the baseline at usWinAscent and size the
// line box as usWinAscent + usWinDescent (no extra gap). When set, they use the
// sTypo* metrics plus typoLineGap. fontkit's `.ascent`/`.descent` expose the
// hhea values, which for these fonts differ from usWinAscent by ~12px at 44pt —
// exactly the baseline offset that otherwise mismatches ground truth. Returns
// font-unit values; the caller scales to px.
interface VMetrics {
  readonly ascent: number;
  readonly descent: number;
  readonly lineGap: number;
}
const verticalMetrics = (font: fontkit.Font): VMetrics => {
  const os2: unknown = (font as { 'OS/2'?: unknown })['OS/2'];
  if (os2 && typeof os2 === 'object') {
    const o = os2 as {
      fsSelection?: { useTypoMetrics?: boolean };
      typoAscender?: number;
      typoDescender?: number;
      typoLineGap?: number;
      winAscent?: number;
      winDescent?: number;
    };
    if (o.fsSelection?.useTypoMetrics && o.typoAscender !== undefined) {
      return {
        ascent: o.typoAscender,
        descent: Math.abs(o.typoDescender ?? 0),
        lineGap: o.typoLineGap ?? 0,
      };
    }
    if (o.winAscent !== undefined && o.winDescent !== undefined) {
      return { ascent: o.winAscent, descent: Math.abs(o.winDescent), lineGap: 0 };
    }
  }
  // Fall back to hhea metrics if OS/2 is absent.
  return { ascent: font.ascent, descent: Math.abs(font.descent), lineGap: font.lineGap };
};

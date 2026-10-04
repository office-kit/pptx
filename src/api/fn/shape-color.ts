// Color transforms and rPr-like element parsing.

import { NAME_A_RPR, requireRun } from './shape-runs.ts';
import {
  PATTERN_PRESETS,
  type ReadGradientStop,
  type ReadTextFormat,
  type TextFormat,
} from '../../internal/drawingml/index.ts';
// Type-only: erased at compile time, so this does not make the modules cyclic.
import type { ShapeEffectAny } from './shape-effects.ts';
import {
  NS,
  type XmlElement,
  firstChildElement,
  getAttrValue,
  qname,
} from '../../internal/xml/index.ts';
import { type SlideShapeData } from '../_internal-symbols.ts';
import { type PresentationTheme } from './theme.ts';
import { readDrawingmlPercentage } from './_drawingml-percentage.ts';
import { resolveDrawingMLPresetColor } from '../../internal/drawingml/preset-colors.ts';
import {
  colorTransformBrightness,
  colorTransformOpacity,
  readColorTransforms,
} from '../../internal/drawingml/color-transforms.ts';
// -- Color transforms (ECMA-376 §20.1.2.3.x) --------------------------------
//
// DrawingML color elements (`<a:srgbClr>`, `<a:schemeClr>`, `<a:sysClr>`,
// `<a:prstClr>`) may carry one or more transform children — `lumMod`,
// `lumOff`, `shade`, `tint`, `satMod`, `hueMod`, `alpha`, `gray`, `inv`,
// `comp`, etc. — that adjust the base color before it's painted. Real
// templates use them heavily for "tinted accent" backgrounds and "shaded
// hover" states, so any visual-fidelity story has to apply them.
//
// Percentages in the spec use the `ST_Percentage` style — `100000`
// represents 100% — though some third-party tools emit bare floats; we
// accept both forms.

type ColorTransformOp =
  | {
      readonly kind:
        | 'lumMod'
        | 'lumOff'
        | 'shade'
        | 'tint'
        | 'satMod'
        | 'satOff'
        | 'hueMod'
        | 'hueOff'
        | 'alpha'
        | 'alphaMod'
        | 'alphaOff';
      readonly val: number;
    }
  | { readonly kind: 'gray' | 'inv' | 'comp' };

const COLOR_TRANSFORM_LOCALS: ReadonlySet<string> = new Set([
  'lumMod',
  'lumOff',
  'shade',
  'tint',
  'satMod',
  'satOff',
  'hueMod',
  'hueOff',
  'alpha',
  'alphaMod',
  'alphaOff',
  'gray',
  'inv',
  'comp',
]);

const readColorPercentage = (raw: string): number => {
  // DrawingML's canonical lexical forms are fixed-point integers and a
  // percent-suffixed value. Preserve the historical bare-float tolerance for
  // third-party files while delegating canonical forms to the shared reader.
  const value = raw.trim();
  if (!value.endsWith('%') && /[.eE]/.test(value)) {
    const number = Number(value);
    // Keep the historical compatibility form for bare fractions while
    // retaining fixed-point semantics for decimal spellings of large values.
    if (Number.isFinite(number) && Math.abs(number) <= 1) return number;
  }
  return readDrawingmlPercentage(value, Number.NaN);
};

const readColorAngleDegrees = (raw: string): number => {
  const trimmed = raw.trim();
  if (!trimmed) return Number.NaN;
  const value = Number(trimmed);
  return Number.isFinite(value) ? value / 60000 : Number.NaN;
};

const parseColorTransforms = (colorEl: XmlElement): readonly ColorTransformOp[] => {
  const out: ColorTransformOp[] = [];
  for (const child of colorEl.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.dml) continue;
    const local = child.name.localName;
    if (!COLOR_TRANSFORM_LOCALS.has(local)) continue;
    if (local === 'gray' || local === 'inv' || local === 'comp') {
      out.push({ kind: local });
      continue;
    }
    const raw = getAttrValue(child, qname('', 'val', ''));
    if (raw === null) continue;
    const n = local === 'hueOff' ? readColorAngleDegrees(raw) : readColorPercentage(raw);
    if (!Number.isFinite(n)) continue;
    out.push({ kind: local as Exclude<ColorTransformOp['kind'], 'gray' | 'inv' | 'comp'>, val: n });
  }
  return out;
};

const hexToRgb01 = (hex: string): [number, number, number] => {
  const h = hex.startsWith('#') ? hex.slice(1) : hex;
  return [
    Number.parseInt(h.slice(0, 2), 16) / 255,
    Number.parseInt(h.slice(2, 4), 16) / 255,
    Number.parseInt(h.slice(4, 6), 16) / 255,
  ];
};

const rgb01ToHex = (r: number, g: number, b: number): string => {
  const clamp = (v: number): number => Math.max(0, Math.min(255, Math.round(v * 255)));
  const part = (n: number): string => n.toString(16).padStart(2, '0').toUpperCase();
  return `#${part(clamp(r))}${part(clamp(g))}${part(clamp(b))}`;
};

const rgbToHsl = (r: number, g: number, b: number): [number, number, number] => {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, s, l];
};

const hueToRgb = (p: number, q: number, t: number): number => {
  if (t < 0) t += 1;
  if (t > 1) t -= 1;
  if (t < 1 / 6) return p + (q - p) * 6 * t;
  if (t < 1 / 2) return q;
  if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
  return p;
};

const hslToRgb = (h: number, s: number, l: number): [number, number, number] => {
  if (s === 0) return [l, l, l];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return [hueToRgb(p, q, h + 1 / 3), hueToRgb(p, q, h), hueToRgb(p, q, h - 1 / 3)];
};

// PowerPoint applies <a:tint> / <a:shade> in LINEAR-LIGHT RGB, not in sRGB —
// this contradicts the literal ECMA-376 "N% of input + (100-N)% white/black"
// wording, but it is what PowerPoint computes and what LibreOffice renders
// (a 75% tint of black is mid-gray ~#8B8B8B, not the sRGB-lerp #404040). The
// next reader will expect the sRGB formula, hence this note.
const srgbToLinear = (c: number): number =>
  c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
const linearToSrgb = (c: number): number =>
  c <= 0.0031308 ? c * 12.92 : 1.055 * c ** (1 / 2.4) - 0.055;

const applyColorTransforms = (hex: string, transforms: readonly ColorTransformOp[]): string => {
  if (transforms.length === 0) return hex;
  let [r, g, b] = hexToRgb01(hex);
  for (const t of transforms) {
    switch (t.kind) {
      case 'inv':
        r = 1 - r;
        g = 1 - g;
        b = 1 - b;
        break;
      case 'gray': {
        const y = 0.3 * r + 0.59 * g + 0.11 * b;
        r = g = b = y;
        break;
      }
      case 'comp': {
        const [h, s, l] = rgbToHsl(r, g, b);
        [r, g, b] = hslToRgb((h + 0.5) % 1, s, l);
        break;
      }
      case 'shade':
        // Mix toward black in linear light: out = srgb(linear(base) * val)
        r = linearToSrgb(srgbToLinear(r) * t.val);
        g = linearToSrgb(srgbToLinear(g) * t.val);
        b = linearToSrgb(srgbToLinear(b) * t.val);
        break;
      case 'tint':
        // Mix toward white in linear light: out = srgb(linear(base)*val + (1-val))
        r = linearToSrgb(srgbToLinear(r) * t.val + (1 - t.val));
        g = linearToSrgb(srgbToLinear(g) * t.val + (1 - t.val));
        b = linearToSrgb(srgbToLinear(b) * t.val + (1 - t.val));
        break;
      case 'lumMod':
      case 'lumOff': {
        const [h, s, l] = rgbToHsl(r, g, b);
        const newL = Math.max(0, Math.min(1, t.kind === 'lumMod' ? l * t.val : l + t.val));
        [r, g, b] = hslToRgb(h, s, newL);
        break;
      }
      case 'satMod':
      case 'satOff': {
        const [h, s, l] = rgbToHsl(r, g, b);
        const newS = Math.max(0, Math.min(1, t.kind === 'satMod' ? s * t.val : s + t.val));
        [r, g, b] = hslToRgb(h, newS, l);
        break;
      }
      case 'hueMod':
      case 'hueOff': {
        const [h, s, l] = rgbToHsl(r, g, b);
        const newH = (((t.kind === 'hueMod' ? h * t.val : h + t.val / 360) % 1) + 1) % 1;
        [r, g, b] = hslToRgb(newH, s, l);
        break;
      }
      // alpha / alphaMod / alphaOff intentionally don't touch RGB —
      // `resolveDrawingColorOpacity` surfaces them as an opacity instead.
    }
  }
  return rgb01ToHex(r, g, b);
};

const SCHEME_TOKEN_TO_THEME_KEY: Record<string, keyof Omit<PresentationTheme, 'name'>> = {
  tx1: 'dark1',
  dk1: 'dark1',
  bg1: 'light1',
  lt1: 'light1',
  tx2: 'dark2',
  dk2: 'dark2',
  bg2: 'light2',
  lt2: 'light2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hlink: 'hyperlink',
  folHlink: 'followedHyperlink',
};

/**
 * Resolves a scheme token (`tx1`, `bg1`, `accent1`, …) to its `#RRGGBB`.
 *
 * When `clrMap` is supplied, the slide token is first remapped through it
 * (`<p:clrMap>` / `<a:overrideClrMapping>`) — `tx1` may point at `dk1` or
 * `lt1` depending on the deck — and only then indexed into the theme. Without
 * a map the token is indexed directly, preserving the historical behavior
 * (correct for the standard map, the overwhelming common case).
 *
 * @internal
 */
export const resolveSchemeToken = (
  token: string,
  theme: PresentationTheme | null,
  clrMap?: Readonly<Record<string, string>> | null,
): string | null => {
  if (!theme) return null;
  const mapped = clrMap?.[token] ?? token;
  const key = SCHEME_TOKEN_TO_THEME_KEY[mapped] ?? SCHEME_TOKEN_TO_THEME_KEY[token];
  if (!key) return null;
  const hex = theme[key];
  if (typeof hex !== 'string') return null;
  const normalized = hex.startsWith('#') ? hex : `#${hex}`;
  return /^#[0-9A-Fa-f]{6}$/.test(normalized) ? normalized.toUpperCase() : null;
};

/**
 * Resolves a DrawingML color element (`<a:srgbClr>` / `<a:schemeClr>` /
 * `<a:sysClr>` / `<a:prstClr>`) with all its `<a:lumMod>` / `<a:tint>` /
 * `<a:shade>` / `<a:satMod>` etc. transform children applied. Returns
 * `null` when the color is a scheme token and no theme is supplied to
 * resolve it.
 *
 * Exposed because both run-format and fill-format code paths need to
 * apply the same transform pipeline; keeping a single implementation
 * means future spec-coverage additions only have to land in one place.
 */
export const resolveDrawingColor = (
  colorEl: XmlElement,
  theme: PresentationTheme | null,
  clrMap?: Readonly<Record<string, string>> | null,
): string | null => {
  if (colorEl.name.namespaceURI !== NS.dml) return null;
  const local = colorEl.name.localName;
  let baseHex: string | null = null;
  if (local === 'srgbClr') {
    const v = getAttrValue(colorEl, qname('', 'val', ''));
    if (v) baseHex = `#${v.toUpperCase()}`;
  } else if (local === 'scrgbClr') {
    // ECMA-376 §20.1.2.3.30 defines these channels as linear-light
    // percentages; convert them to the sRGB encoding used by this API.
    const r = readColorPercentage(getAttrValue(colorEl, qname('', 'r', '')) ?? '');
    const g = readColorPercentage(getAttrValue(colorEl, qname('', 'g', '')) ?? '');
    const b = readColorPercentage(getAttrValue(colorEl, qname('', 'b', '')) ?? '');
    if ([r, g, b].every(Number.isFinite)) {
      baseHex = rgb01ToHex(linearToSrgb(r), linearToSrgb(g), linearToSrgb(b));
    }
  } else if (local === 'hslClr') {
    // ECMA-376 §20.1.2.3.13 uses a positive fixed angle (1/60000 degree)
    // plus percentage saturation and luminance attributes.
    const hue = readColorAngleDegrees(getAttrValue(colorEl, qname('', 'hue', '')) ?? '');
    const sat = readColorPercentage(getAttrValue(colorEl, qname('', 'sat', '')) ?? '');
    const lum = readColorPercentage(getAttrValue(colorEl, qname('', 'lum', '')) ?? '');
    if ([hue, sat, lum].every(Number.isFinite)) {
      const [r, g, b] = hslToRgb((((hue / 360) % 1) + 1) % 1, sat, lum);
      baseHex = rgb01ToHex(r, g, b);
    }
  } else if (local === 'schemeClr') {
    const v = getAttrValue(colorEl, qname('', 'val', ''));
    if (v) baseHex = resolveSchemeToken(v, theme, clrMap);
  } else if (local === 'sysClr') {
    const last = getAttrValue(colorEl, qname('', 'lastClr', ''));
    if (last) baseHex = `#${last.toUpperCase()}`;
  } else if (local === 'prstClr') {
    const v = getAttrValue(colorEl, qname('', 'val', ''));
    if (v) baseHex = resolveDrawingMLPresetColor(v);
  }
  if (!baseHex) return null;
  return applyColorTransforms(baseHex, parseColorTransforms(colorEl));
};

/**
 * Resolves the opacity a DrawingML color element carries through its
 * `<a:alpha>` / `<a:alphaMod>` / `<a:alphaOff>` children (ECMA-376
 * §20.1.2.3.1–3) to a `0`–`1` fraction, applied in document order from a
 * fully opaque base. Returns `null` when the element has no alpha
 * transform — PowerPoint paints that opaque, but callers can still tell
 * "unspecified" apart from an explicit `1`.
 *
 * Kept separate from `resolveDrawingColor` because alpha never changes the
 * `#RRGGBB`; OOXML encodes the two independently and renderers emit the
 * opacity next to the color (`fill-opacity` / `stroke-opacity`).
 */
export const resolveDrawingColorOpacity = (colorEl: XmlElement): number | null => {
  if (colorEl.name.namespaceURI !== NS.dml) return null;
  let opacity: number | null = null;
  for (const t of parseColorTransforms(colorEl)) {
    if (t.kind === 'alpha') opacity = t.val;
    else if (t.kind === 'alphaMod') opacity = (opacity ?? 1) * t.val;
    else if (t.kind === 'alphaOff') opacity = (opacity ?? 1) + t.val;
  }
  return opacity === null ? null : Math.max(0, Math.min(1, opacity));
};

/**
 * Parses an `<a:effectLst>` into the typed effect union. It lives here rather
 * than beside the shape effect API because a run's `<a:rPr>` carries the same
 * element, and this module is the one both sides can depend on.
 */
export const parseEffectList = (
  effectLst: XmlElement,
  theme: PresentationTheme | null,
  colorMap?: Readonly<Record<string, string>> | null,
): ShapeEffectAny[] => {
  const readEffectColor = (host: XmlElement): { color: string; opacity?: number } => {
    let inner: XmlElement | null = null;
    for (const c of host.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      if (
        c.name.localName === 'srgbClr' ||
        c.name.localName === 'scrgbClr' ||
        c.name.localName === 'hslClr' ||
        c.name.localName === 'schemeClr' ||
        c.name.localName === 'sysClr' ||
        c.name.localName === 'prstClr'
      ) {
        inner = c;
        break;
      }
    }
    if (!inner) return { color: '' };
    const opacity = resolveDrawingColorOpacity(inner);
    const hex = resolveDrawingColor(inner, theme, colorMap);
    return { color: hex ?? '', ...(opacity !== null ? { opacity } : {}) };
  };

  const out: ShapeEffectAny[] = [];
  for (const child of effectLst.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.dml) continue;
    const local = child.name.localName;
    if (local === 'outerShdw' || local === 'innerShdw') {
      const blur = Number.parseInt(getAttrValue(child, qname('', 'blurRad', '')) ?? '0', 10) || 0;
      const dist = Number.parseInt(getAttrValue(child, qname('', 'dist', '')) ?? '0', 10) || 0;
      const dir = Number.parseInt(getAttrValue(child, qname('', 'dir', '')) ?? '0', 10) || 0;
      const c = readEffectColor(child);
      const alignmentRaw = getAttrValue(child, qname('', 'algn', '')) ?? 'b';
      const alignment = ['tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br'].includes(alignmentRaw)
        ? (alignmentRaw as 'tl' | 't' | 'tr' | 'l' | 'ctr' | 'r' | 'bl' | 'b' | 'br')
        : undefined;
      const rotationRaw = getAttrValue(child, qname('', 'rotWithShape', '')) ?? 'true';
      const shadow = {
        color: c.color,
        blurEmu: blur,
        distEmu: dist,
        angleDeg: dir / 60000,
        ...(c.opacity !== undefined ? { opacity: c.opacity } : {}),
      };
      out.push(
        local === 'outerShdw'
          ? {
              kind: 'outerShdw',
              ...shadow,
              ...(alignment !== undefined ? { alignment } : {}),
              rotateWithShape: rotationRaw !== '0' && rotationRaw !== 'false',
            }
          : { kind: 'innerShdw', ...shadow },
      );
    } else if (local === 'glow') {
      const rad = Number.parseInt(getAttrValue(child, qname('', 'rad', '')) ?? '0', 10) || 0;
      const c = readEffectColor(child);
      out.push({
        kind: 'glow',
        color: c.color,
        radiusEmu: rad,
        ...(c.opacity !== undefined ? { opacity: c.opacity } : {}),
      });
    } else if (local === 'reflection') {
      const blur = Number.parseInt(getAttrValue(child, qname('', 'blurRad', '')) ?? '0', 10) || 0;
      const dist = Number.parseInt(getAttrValue(child, qname('', 'dist', '')) ?? '0', 10) || 0;
      const dir = Number.parseInt(getAttrValue(child, qname('', 'dir', '')) ?? '0', 10) || 0;
      // `stA`/`endA` are ST_PositiveFixedPercentage (0..100000); `sy` is
      // ST_Percentage and may be negative to encode the mirror flip.
      const pctFraction = (name: string): number | undefined => {
        const raw = getAttrValue(child, qname('', name, ''));
        if (raw === null) return undefined;
        const n = readColorPercentage(raw);
        if (!Number.isFinite(n)) return undefined;
        return n;
      };
      const opacity = pctFraction('endA');
      const startOpacity = pctFraction('stA');
      const startPosition = pctFraction('stPos');
      const endPosition = pctFraction('endPos');
      const fadeDirectionRaw = getAttrValue(child, qname('', 'fadeDir', ''));
      const fadeDirection =
        fadeDirectionRaw === null ? undefined : readColorAngleDegrees(fadeDirectionRaw);
      const scaleX = pctFraction('sx');
      const scaleY = pctFraction('sy');
      const skewXRaw = getAttrValue(child, qname('', 'kx', ''));
      const skewYRaw = getAttrValue(child, qname('', 'ky', ''));
      const skewX = skewXRaw === null ? undefined : readColorAngleDegrees(skewXRaw);
      const skewY = skewYRaw === null ? undefined : readColorAngleDegrees(skewYRaw);
      const alignmentRaw = getAttrValue(child, qname('', 'algn', ''));
      const alignment =
        alignmentRaw !== null &&
        ['tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br'].includes(alignmentRaw)
          ? (alignmentRaw as 'tl' | 't' | 'tr' | 'l' | 'ctr' | 'r' | 'bl' | 'b' | 'br')
          : undefined;
      const rotateWithShape = getAttrValue(child, qname('', 'rotWithShape', ''));
      out.push({
        kind: 'reflection',
        blurEmu: blur,
        distEmu: dist,
        angleDeg: dir / 60000,
        ...(opacity !== undefined ? { opacity } : {}),
        ...(startOpacity !== undefined ? { startOpacity } : {}),
        ...(startPosition !== undefined ? { startPosition } : {}),
        ...(endPosition !== undefined ? { endPosition } : {}),
        ...(fadeDirection !== undefined ? { fadeDirection } : {}),
        ...(scaleX !== undefined ? { scaleX } : {}),
        ...(scaleY !== undefined ? { scaleY } : {}),
        ...(skewX !== undefined ? { skewX } : {}),
        ...(skewY !== undefined ? { skewY } : {}),
        ...(alignment !== undefined ? { alignment } : {}),
        ...(rotateWithShape !== null
          ? { rotateWithShape: rotateWithShape !== '0' && rotateWithShape !== 'false' }
          : {}),
      });
    } else if (local === 'softEdge') {
      const rad = Number.parseInt(getAttrValue(child, qname('', 'rad', '')) ?? '0', 10) || 0;
      out.push({ kind: 'softEdge', radiusEmu: rad });
    } else if (local === 'blur') {
      const rad = Number.parseInt(getAttrValue(child, qname('', 'rad', '')) ?? '0', 10) || 0;
      out.push({ kind: 'blur', radiusEmu: rad });
    }
  }
  return out;
};

// The color of an `<a:solidFill>`, as a run reports it. With a theme, scheme
// tokens resolve to `#RRGGBB` and `<a:lumMod>` and friends are applied; without
// one, tokens pass through verbatim, which is the legacy `getShapeRunFormat`
// behavior callers round-trip against.
const colorOfFill = (
  solidFill: XmlElement,
  ctx?: {
    readonly theme: PresentationTheme | null;
    readonly colorMap?: Readonly<Record<string, string>> | null;
  },
): string | null => {
  // CT_SolidColorFillProperties holds exactly one EG_ColorChoice child
  // (srgbClr / schemeClr / sysClr / prstClr).
  let colorChild: XmlElement | null = null;
  for (const c of solidFill.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    colorChild = c;
    break;
  }
  if (colorChild === null) return null;
  const token = getAttrValue(colorChild, qname('', 'val', ''));
  if (ctx) {
    const hex = resolveDrawingColor(colorChild, ctx.theme, ctx.colorMap);
    if (hex !== null) return hex;
    // Theme not provided / token not in scheme — surface the raw token.
    return colorChild.name.localName === 'schemeClr' ? token : null;
  }
  if (colorChild.name.localName === 'srgbClr')
    return token === null ? null : `#${token.toUpperCase()}`;
  if (colorChild.name.localName === 'schemeClr') return token;
  return null;
};

const readTextGradientFill = (
  gradFill: XmlElement,
  ctx?: {
    readonly theme: PresentationTheme | null;
    readonly colorMap?: Readonly<Record<string, string>> | null;
  },
): ReadTextFormat['textFill'] => {
  const gsList = firstChildElement(gradFill, qname('a', 'gsLst', NS.dml));
  if (gsList === null) return undefined;
  const stops: ReadGradientStop[] = [];
  for (const gs of gsList.children) {
    if (gs.kind !== 'element' || gs.name.localName !== 'gs') continue;
    const pos = readDrawingmlPercentage(getAttrValue(gs, qname('', 'pos', '')) ?? '', Number.NaN);
    const colorEl = gs.children.find(
      (child) => child.kind === 'element' && child.name.namespaceURI === NS.dml,
    );
    if (!Number.isFinite(pos) || colorEl?.kind !== 'element') continue;
    const base = colorOfFill({ ...gs, children: [colorEl] }, ctx);
    if (base === null) continue;
    const transforms = readColorTransforms(colorEl);
    const opacity = colorTransformOpacity(transforms);
    const brightness = colorTransformBrightness(transforms);
    const resolvedColor = ctx ? resolveDrawingColor(colorEl, ctx.theme, ctx.colorMap) : null;
    stops.push({
      offset: pos,
      color: base,
      ...(transforms.length ? { colorTransforms: transforms } : {}),
      ...(opacity === undefined ? {} : { opacity }),
      ...(brightness === undefined ? {} : { brightness }),
      ...(resolvedColor === null ? {} : { resolvedColor }),
    });
  }
  if (stops.length === 0) return undefined;
  const lin = firstChildElement(gradFill, qname('a', 'lin', NS.dml));
  const path = firstChildElement(gradFill, qname('a', 'path', NS.dml));
  const rect = (el: XmlElement | null) => {
    if (el === null) return undefined;
    const read = (name: string) => {
      const raw = getAttrValue(el, qname('', name, ''));
      return raw === null ? 0 : Number.parseFloat(raw) / (raw.endsWith('%') ? 100 : 100000);
    };
    return { left: read('l'), top: read('t'), right: read('r'), bottom: read('b') };
  };
  const angle =
    lin === null ? undefined : Number.parseInt(getAttrValue(lin, qname('', 'ang', '')) ?? '', 10);
  const pathValue = getAttrValue(path ?? gradFill, qname('', 'path', ''));
  const kind =
    pathValue === 'circle' || pathValue === 'rect' || pathValue === 'shape' ? pathValue : undefined;
  const tile = firstChildElement(gradFill, qname('a', 'tileRect', NS.dml));
  return {
    kind: 'gradient',
    stops,
    ...(angle !== undefined && Number.isFinite(angle) ? { angleDeg: angle / 60000 } : {}),
    ...(lin !== null
      ? {
          scaled: ['1', 'true'].includes(getAttrValue(lin, qname('', 'scaled', '')) ?? ''),
        }
      : {}),
    ...(getAttrValue(gradFill, qname('', 'rotWithShape', '')) !== null
      ? {
          rotateWithShape: !['0', 'false'].includes(
            getAttrValue(gradFill, qname('', 'rotWithShape', '')) ?? '',
          ),
        }
      : {}),
    ...(kind ? { path: kind } : {}),
    ...(path
      ? (() => {
          const f = rect(firstChildElement(path, qname('a', 'fillToRect', NS.dml)));
          return f ? { focus: f } : {};
        })()
      : {}),
    ...(tile
      ? (() => {
          const t = rect(tile);
          return t ? { tileRect: t } : {};
        })()
      : {}),
  };
};

const readTextPatternFill = (
  pattFill: XmlElement,
  ctx?: {
    readonly theme: PresentationTheme | null;
    readonly colorMap?: Readonly<Record<string, string>> | null;
  },
): ReadTextFormat['textFill'] => {
  const preset = getAttrValue(pattFill, qname('', 'prst', ''));
  if (preset === null) return undefined;
  const read = (local: 'fgClr' | 'bgClr') => {
    const holder = firstChildElement(pattFill, qname('a', local, NS.dml));
    return holder === null ? null : colorOfFill(holder, ctx);
  };
  const foreground = read('fgClr');
  const background = read('bgClr');
  if (foreground === null || background === null) return undefined;
  if (!PATTERN_PRESETS.includes(preset as (typeof PATTERN_PRESETS)[number])) return undefined;
  const transforms = (local: 'fgClr' | 'bgClr') => {
    const holder = firstChildElement(pattFill, qname('a', local, NS.dml));
    const color = holder?.children.find(
      (child) => child.kind === 'element' && child.name.namespaceURI === NS.dml,
    );
    return color?.kind === 'element' ? readColorTransforms(color) : [];
  };
  const foregroundTransforms = transforms('fgClr');
  const backgroundTransforms = transforms('bgClr');
  const validatedPreset = preset as (typeof PATTERN_PRESETS)[number];
  return {
    kind: 'pattern',
    preset: validatedPreset,
    foreground,
    background,
    ...(foregroundTransforms.length ? { foregroundTransforms } : {}),
    ...(backgroundTransforms.length ? { backgroundTransforms } : {}),
  };
};

// Reads any element shaped like `CT_TextCharacterProperties` (the schema
// shared by `<a:rPr>`, `<a:defRPr>`, and `<a:endParaRPr>`) into a partial
// TextFormat. Used by both the literal-only `getShapeRunFormat` and the
// inheritance-aware `getShapeRunFormatEffective`.
//
// When `ctx.theme` is provided, scheme tokens are resolved to concrete
// `#RRGGBB` and color transforms (`<a:lumMod>` etc.) are applied. Without
// a theme, transforms are not applied and theme tokens are passed through
// verbatim — this preserves the legacy `getShapeRunFormat` behavior.
export const parseRPrLikeElement = (
  rPr: XmlElement,
  ctx?: {
    readonly theme: PresentationTheme | null;
    readonly colorMap?: Readonly<Record<string, string>> | null;
  },
): Partial<ReadTextFormat> => {
  const out: Partial<ReadTextFormat> = {};
  const sz = getAttrValue(rPr, qname('', 'sz', ''));
  if (sz !== null) {
    const n = Number.parseInt(sz, 10);
    if (Number.isFinite(n)) out.size = n / 100;
  }
  const b = getAttrValue(rPr, qname('', 'b', ''));
  if (b !== null) out.bold = b === '1' || b === 'true';
  const i = getAttrValue(rPr, qname('', 'i', ''));
  if (i !== null) out.italic = i === '1' || i === 'true';
  const u = getAttrValue(rPr, qname('', 'u', ''));
  if (u !== null) {
    if (u === 'none') out.underline = false;
    else if (u === 'sng') out.underline = true;
    else out.underline = u;
  }
  const strike = getAttrValue(rPr, qname('', 'strike', ''));
  if (strike !== null) {
    if (strike === 'noStrike') out.strike = false;
    else if (strike === 'sngStrike') out.strike = true;
    else out.strike = strike;
  }
  const spc = getAttrValue(rPr, qname('', 'spc', ''));
  if (spc !== null) {
    const n = Number.parseInt(spc, 10);
    if (Number.isFinite(n)) out.spc = n;
  }
  const kern = getAttrValue(rPr, qname('', 'kern', ''));
  if (kern !== null) {
    const n = Number.parseInt(kern, 10);
    if (Number.isFinite(n)) out.kern = n;
  }
  const baselineAttr = getAttrValue(rPr, qname('', 'baseline', ''));
  const normalizeHeight = getAttrValue(rPr, qname('', 'normalizeH', ''));
  if (normalizeHeight !== null)
    out.normalizeHeight = normalizeHeight === '1' || normalizeHeight === 'true';
  if (baselineAttr !== null) {
    const baseline = readDrawingmlPercentage(baselineAttr, Number.NaN);
    if (Number.isFinite(baseline)) out.baseline = baseline;
  }
  const cap = getAttrValue(rPr, qname('', 'cap', ''));
  if (cap === 'none' || cap === 'small' || cap === 'all') {
    out.cap = cap;
  }
  // <a:highlight><a:srgbClr val="…"/></a:highlight>
  const highlight = firstChildElement(rPr, qname('a', 'highlight', NS.dml));
  if (highlight !== null) {
    let hlChild: XmlElement | null = null;
    for (const c of highlight.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      hlChild = c;
      break;
    }
    if (hlChild) {
      if (ctx) {
        const hex = resolveDrawingColor(hlChild, ctx.theme, ctx.colorMap);
        if (hex !== null) out.highlight = hex;
      } else if (hlChild.name.localName === 'srgbClr') {
        const v = getAttrValue(hlChild, qname('', 'val', ''));
        if (v !== null) out.highlight = `#${v.toUpperCase()}`;
      } else if (hlChild.name.localName === 'schemeClr') {
        const v = getAttrValue(hlChild, qname('', 'val', ''));
        if (v !== null) out.highlight = v;
      }
    }
  }
  const solidFill = firstChildElement(rPr, qname('a', 'solidFill', NS.dml));
  if (solidFill !== null) {
    const color = colorOfFill(solidFill, ctx);
    if (color !== null) out.color = color;
  }
  const gradientFill = firstChildElement(rPr, qname('a', 'gradFill', NS.dml));
  if (gradientFill !== null) {
    const fill = readTextGradientFill(gradientFill, ctx);
    if (fill) out.textFill = fill;
  }
  const patternFill = firstChildElement(rPr, qname('a', 'pattFill', NS.dml));
  if (patternFill !== null) {
    const fill = readTextPatternFill(patternFill, ctx);
    if (fill) out.textFill = fill;
  }
  // Underline fill is a separate DrawingML choice from the run's text fill.
  // `uFillTx` is meaningful even without a color child: it explicitly follows
  // the text color and must therefore remain distinct from an omitted value.
  const underlineFillText = firstChildElement(rPr, qname('a', 'uFillTx', NS.dml));
  if (underlineFillText !== null) {
    out.underlineColor = null;
  } else {
    const underlineFill = firstChildElement(rPr, qname('a', 'uFill', NS.dml));
    const underlineSolidFill =
      underlineFill === null
        ? null
        : firstChildElement(underlineFill, qname('a', 'solidFill', NS.dml));
    if (underlineSolidFill !== null) {
      const color = colorOfFill(underlineSolidFill, ctx);
      if (color !== null) out.underlineColor = color;
    }
  }
  const latin = firstChildElement(rPr, qname('a', 'latin', NS.dml));
  if (latin !== null) {
    const t = getAttrValue(latin, qname('', 'typeface', ''));
    if (t !== null) out.font = t;
  }
  const ea = firstChildElement(rPr, qname('a', 'ea', NS.dml));
  if (ea !== null) {
    const t = getAttrValue(ea, qname('', 'typeface', ''));
    if (t !== null) out.fontEastAsian = t;
  }
  const cs = firstChildElement(rPr, qname('a', 'cs', NS.dml));
  if (cs !== null) {
    const t = getAttrValue(cs, qname('', 'typeface', ''));
    if (t !== null) out.fontComplexScript = t;
  }
  const ln = firstChildElement(rPr, qname('a', 'ln', NS.dml));
  if (ln !== null) {
    const outline: { color?: string; widthEmu?: number } = {};
    const w = getAttrValue(ln, qname('', 'w', ''));
    if (w !== null) {
      const n = Number.parseInt(w, 10);
      if (Number.isFinite(n)) outline.widthEmu = n;
    }
    const lnFill = firstChildElement(ln, qname('a', 'solidFill', NS.dml));
    const color = lnFill === null ? null : colorOfFill(lnFill, ctx);
    if (color !== null) outline.color = color;
    out.outline = outline;
  }
  // `<a:effectLst>` on a run holds the same effects as on a shape; a run that
  // states one states it for its own glyphs. Only effects the library writes
  // are surfaced here — the rest stay readable through `getShapeEffects`'
  // union, which is not what a character format is.
  const effects = firstChildElement(rPr, qname('a', 'effectLst', NS.dml));
  if (effects !== null) {
    for (const effect of parseEffectList(effects, ctx?.theme ?? null)) {
      if (effect.kind === 'outerShdw') {
        out.shadow = {
          color: effect.color,
          ...(effect.alignment !== undefined ? { alignment: effect.alignment } : {}),
          ...(effect.rotateWithShape !== undefined
            ? { rotateWithShape: effect.rotateWithShape }
            : {}),
          blurEmu: effect.blurEmu,
          offsetEmu: effect.distEmu,
          angleDeg: effect.angleDeg,
          ...(effect.opacity !== undefined ? { opacity: effect.opacity } : {}),
        };
      } else if (effect.kind === 'innerShdw') {
        out.innerShadow = {
          color: effect.color,
          blurEmu: effect.blurEmu,
          offsetEmu: effect.distEmu,
          angleDeg: effect.angleDeg,
          ...(effect.opacity !== undefined ? { opacity: effect.opacity } : {}),
        };
      } else if (effect.kind === 'glow') {
        out.glow = {
          color: effect.color,
          radiusEmu: effect.radiusEmu,
          ...(effect.opacity !== undefined ? { opacity: effect.opacity } : {}),
        };
      } else if (effect.kind === 'reflection') {
        out.reflection = {
          blurEmu: effect.blurEmu,
          offsetEmu: effect.distEmu,
          angleDeg: effect.angleDeg,
          ...(effect.opacity !== undefined ? { opacity: effect.opacity } : {}),
          ...(effect.startOpacity !== undefined ? { startOpacity: effect.startOpacity } : {}),
          ...(effect.startPosition !== undefined ? { startPosition: effect.startPosition } : {}),
          ...(effect.endPosition !== undefined ? { endPosition: effect.endPosition } : {}),
          ...(effect.fadeDirection !== undefined ? { fadeDirection: effect.fadeDirection } : {}),
          ...(effect.scaleX !== undefined ? { scaleX: effect.scaleX } : {}),
          ...(effect.scaleY !== undefined ? { scaleY: effect.scaleY } : {}),
          ...(effect.skewX !== undefined ? { skewX: effect.skewX } : {}),
          ...(effect.skewY !== undefined ? { skewY: effect.skewY } : {}),
          ...(effect.alignment !== undefined ? { alignment: effect.alignment } : {}),
          ...(effect.rotateWithShape !== undefined
            ? { rotateWithShape: effect.rotateWithShape }
            : {}),
        };
      }
    }
  }
  return out;
};

/**
 * Reads back the format of a single run. Returns `null` when the run
 * has no `<a:rPr>` (it inherits its format from the paragraph /
 * layout / master). Boolean attributes that are explicitly `"0"`
 * decode to `false`.
 *
 * Use `getShapeRunFormatEffective` if you want the resolved format
 * after walking the placeholder / lstStyle / master inheritance chain.
 */
export const getShapeRunFormat = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
): ReadTextFormat | null => {
  const run = requireRun(shape, paragraphIndex, runIndex);
  const rPr = firstChildElement(run, NAME_A_RPR);
  if (rPr === null) return null;
  return parseRPrLikeElement(rPr) as TextFormat;
};

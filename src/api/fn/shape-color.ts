// Color transforms and rPr-like element parsing.

import { NAME_A_RPR, requireRun } from './shape-runs.ts';
import { type TextFormat } from '../../internal/drawingml/index.ts';
import {
  NS,
  type XmlElement,
  firstChildElement,
  getAttrValue,
  qname,
} from '../../internal/xml/index.ts';
import { type SlideShapeData } from '../_internal-symbols.ts';
import { type PresentationTheme } from './theme.ts';
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
    let n = Number.parseFloat(raw);
    if (!Number.isFinite(n)) continue;
    // PowerPoint emits ST_Percentage (`100000` = 100%); tolerate the
    // bare-float form some third-party tools emit.
    if (Math.abs(n) > 1) n = n / 100000;
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
        // Mix toward black: out = base * val
        r *= t.val;
        g *= t.val;
        b *= t.val;
        break;
      case 'tint':
        // Mix toward white: out = base * val + (1 - val)
        r = r * t.val + (1 - t.val);
        g = g * t.val + (1 - t.val);
        b = b * t.val + (1 - t.val);
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
      // alpha / alphaMod / alphaOff intentionally don't touch RGB — they
      // surface as `fill-opacity`, not as a tinted color.
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

const resolveSchemeToken = (token: string, theme: PresentationTheme | null): string | null => {
  if (!theme) return null;
  const key = SCHEME_TOKEN_TO_THEME_KEY[token];
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
): string | null => {
  if (colorEl.name.namespaceURI !== NS.dml) return null;
  const local = colorEl.name.localName;
  let baseHex: string | null = null;
  if (local === 'srgbClr') {
    const v = getAttrValue(colorEl, qname('', 'val', ''));
    if (v) baseHex = `#${v.toUpperCase()}`;
  } else if (local === 'schemeClr') {
    const v = getAttrValue(colorEl, qname('', 'val', ''));
    if (v) baseHex = resolveSchemeToken(v, theme);
  } else if (local === 'sysClr') {
    const last = getAttrValue(colorEl, qname('', 'lastClr', ''));
    if (last) baseHex = `#${last.toUpperCase()}`;
  } else if (local === 'prstClr') {
    // Preset colors aren't worth a full lookup table in this pass —
    // black / white cover most cases anyone reaches for in PresentationML.
    const v = getAttrValue(colorEl, qname('', 'val', ''));
    if (v === 'black') baseHex = '#000000';
    else if (v === 'white') baseHex = '#FFFFFF';
  }
  if (!baseHex) return null;
  return applyColorTransforms(baseHex, parseColorTransforms(colorEl));
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
  ctx?: { readonly theme: PresentationTheme | null },
): Partial<TextFormat> => {
  const out: Partial<TextFormat> = {};
  const sz = getAttrValue(rPr, qname('', 'sz', ''));
  if (sz !== null) {
    const n = Number.parseInt(sz, 10);
    if (Number.isFinite(n)) out.size = n / 100;
  }
  const b = getAttrValue(rPr, qname('', 'b', ''));
  if (b !== null) out.bold = b !== '0';
  const i = getAttrValue(rPr, qname('', 'i', ''));
  if (i !== null) out.italic = i !== '0';
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
  if (baselineAttr !== null) {
    // ST_Percentage: 100000 = 100%; tolerate bare floats.
    let n = Number.parseFloat(baselineAttr);
    if (Number.isFinite(n)) {
      if (Math.abs(n) > 1) n = n / 100000;
      out.baseline = n;
    }
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
        const hex = resolveDrawingColor(hlChild, ctx.theme);
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
    // Find the inner color element (srgbClr / schemeClr / sysClr / prstClr).
    // CT_SolidColorFillProperties holds exactly one EG_ColorChoice child.
    let colorChild: XmlElement | null = null;
    for (const c of solidFill.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      colorChild = c;
      break;
    }
    if (colorChild) {
      if (ctx) {
        // Apply transforms + resolve scheme tokens to hex.
        const hex = resolveDrawingColor(colorChild, ctx.theme);
        if (hex !== null) out.color = hex;
        else if (colorChild.name.localName === 'schemeClr') {
          // Theme not provided / token not in scheme — surface the raw token.
          const v = getAttrValue(colorChild, qname('', 'val', ''));
          if (v !== null) out.color = v;
        }
      } else {
        // Legacy `getShapeRunFormat` path: no transforms, scheme tokens
        // emitted as bare strings to match prior public behavior.
        if (colorChild.name.localName === 'srgbClr') {
          const v = getAttrValue(colorChild, qname('', 'val', ''));
          if (v !== null) out.color = `#${v.toUpperCase()}`;
        } else if (colorChild.name.localName === 'schemeClr') {
          const v = getAttrValue(colorChild, qname('', 'val', ''));
          if (v !== null) out.color = v;
        }
      }
    }
  }
  const latin = firstChildElement(rPr, qname('a', 'latin', NS.dml));
  if (latin !== null) {
    const t = getAttrValue(latin, qname('', 'typeface', ''));
    if (t !== null) out.font = t;
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
): TextFormat | null => {
  const run = requireRun(shape, paragraphIndex, runIndex);
  const rPr = firstChildElement(run, NAME_A_RPR);
  if (rPr === null) return null;
  return parseRPrLikeElement(rPr) as TextFormat;
};

// Lays out an OMML equation (the `math` paragraph element of
// `@office-kit/pptx`) as one inline SVG box for the text engine.
//
// The rules follow TeX's math layout in simplified form: a display equation
// (`<m:oMathPara>`) keeps fractions, matrices and limits at full size, while an
// inline one sets them at script size with n-ary limits beside the operator.
// The desktop app's own fallback image of an inline equation shows the same
// reductions. Metrics are fixed em fractions instead of the font's MATH table:
// the bundled substitutes have none, and a preview needs a readable equation,
// not glyph-exact placement.

import type { MathNode, MathNodes, ReadTextFormat } from '@office-kit/pptx';
import type { InlineBox, TextMeasurer } from './text-layout.ts';

export interface MathFont {
  /** Font family as passed to the measurer. */
  readonly family: string;
  /** `font-family` attribute value, when it differs from `family`. */
  readonly cssFamily?: string;
  readonly sizePx: number;
  readonly fillHex: string;
  /** Overrides from a math run's own `<a:rPr>`. */
  readonly runStyle?: (format: ReadTextFormat) => { sizePx?: number; fillHex?: string };
}

// Em fractions of the current size.
const AXIS = 0.25; // math axis above the baseline
const RULE = 0.05; // fraction bar / radical overline thickness
const STRUT_ASCENT = 0.7;
const STRUT_DESCENT = 0.2;
const SCRIPT_RATIO = 0.7;
const SCRIPT_SCRIPT_RATIO = 0.5;
const THIN_SPACE = 1 / 6;
const MEDIUM_SPACE = 2 / 9;
const THICK_SPACE = 5 / 18;
const DISPLAY_OPERATOR_SCALE = 1.4;
const UNSUPPORTED_STROKE = '#D00000';

interface Ctx {
  readonly sizePx: number;
  readonly baseSizePx: number;
  readonly fillHex: string;
  readonly display: boolean;
  readonly level: number;
  readonly upright: boolean;
}

interface Box {
  readonly w: number;
  readonly a: number;
  readonly d: number;
  readonly draw: (x: number, y: number, out: string[]) => void;
}

const fmt = (n: number): string => {
  const r = Math.round(n * 100) / 100;
  return Object.is(r, -0) ? '0' : String(r);
};

const escapeXml = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const emptyBox = (): Box => ({ w: 0, a: 0, d: 0, draw: () => {} });

const shifted = (box: Box, dy: number): Box => ({
  w: box.w,
  a: box.a - dy,
  d: box.d + dy,
  draw: (x, y, out) => box.draw(x, y + dy, out),
});

const padded = (box: Box, left: number, right: number): Box => ({
  w: left + box.w + right,
  a: box.a,
  d: box.d,
  draw: (x, y, out) => box.draw(x + left, y, out),
});

const hlist = (boxes: ReadonlyArray<Box>): Box => {
  let w = 0;
  let a = 0;
  let d = 0;
  for (const box of boxes) {
    w += box.w;
    a = Math.max(a, box.a);
    d = Math.max(d, box.d);
  }
  return {
    w,
    a,
    d,
    draw: (x, y, out) => {
      let cursor = x;
      for (const box of boxes) {
        box.draw(cursor, y, out);
        cursor += box.w;
      }
    },
  };
};

// A function name's upright style covers its own letters, not a script or
// limit attached to them (the `n` of `lim_{n→∞}` stays italic).
const scriptCtx = (ctx: Ctx): Ctx => ({
  ...ctx,
  sizePx: ctx.baseSizePx * (ctx.level === 0 ? SCRIPT_RATIO : SCRIPT_SCRIPT_RATIO),
  display: false,
  level: ctx.level + 1,
  upright: false,
});

// Fraction parts and matrix cells: one style smaller, so full size in a display
// equation and script size inline.
const smallerCtx = (ctx: Ctx): Ctx => (ctx.display ? { ...ctx, display: false } : scriptCtx(ctx));

// ---------------------------------------------------------------------------
// Characters.

// Mathematical Alphanumeric Symbols (U+1D400 block): the desktop app stores
// variables as these italic code points. The bundled fonts lack the block, so
// fold them back to the base letter with the matching style.
interface Glyph {
  readonly ch: string;
  readonly bold: boolean;
  readonly italic: boolean | null; // null: the math default for the character
}

const LATIN_RANGES: ReadonlyArray<{ start: number; bold: boolean; italic: boolean }> = [
  { start: 0x1d400, bold: true, italic: false },
  { start: 0x1d434, bold: false, italic: true },
  { start: 0x1d468, bold: true, italic: true },
];
const GREEK_RANGES: ReadonlyArray<{ start: number; bold: boolean; italic: boolean }> = [
  { start: 0x1d6a8, bold: true, italic: false },
  { start: 0x1d6e2, bold: false, italic: true },
  { start: 0x1d71c, bold: true, italic: true },
];
// Each Greek range: 25 capitals (Α..Ω with ϴ at the 18th slot), ∇, then
// 25 small letters (α..ω with ς), ∂, ϵ ϑ ϰ ϕ ϱ ϖ.
const GREEK_SEQUENCE = [...'ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡϴΣΤΥΦΧΨΩ∇', ...'αβγδεζηθικλμνξοπρςστυφχψω∂ϵϑϰϕϱϖ'];
const ITALIC_SMALL_H = 0x210e;

const glyphOf = (ch: string): Glyph => {
  const cp = ch.codePointAt(0)!;
  if (cp === ITALIC_SMALL_H) return { ch: 'h', bold: false, italic: true };
  for (const range of LATIN_RANGES) {
    const offset = cp - range.start;
    if (offset >= 0 && offset < 52) {
      const letter = offset < 26 ? 0x41 + offset : 0x61 + offset - 26;
      return { ch: String.fromCodePoint(letter), bold: range.bold, italic: range.italic };
    }
  }
  for (const range of GREEK_RANGES) {
    const offset = cp - range.start;
    if (offset >= 0 && offset < GREEK_SEQUENCE.length) {
      return { ch: GREEK_SEQUENCE[offset]!, bold: range.bold, italic: range.italic };
    }
  }
  if (cp >= 0x1d7ce && cp <= 0x1d7d7) {
    return { ch: String.fromCodePoint(0x30 + cp - 0x1d7ce), bold: true, italic: false };
  }
  // In a math zone a hyphen-minus is the minus sign.
  if (ch === '-') return { ch: '−', bold: false, italic: null };
  return { ch, bold: false, italic: null };
};

const RELATIONS = new Set('=<>≤≥≠≈≡∼≅∝→←↔⇒⇐⇔∈∉∋⊂⊃⊆⊇≪≫∣');
const BINARIES = new Set('+−±∓×÷·∗∘∙∪∩⊕⊗∧∨');
const PUNCTUATION = new Set(',;');
const LETTER = /\p{L}/u;

type AtomClass = 'ord' | 'rel' | 'bin' | 'punct';

const classOf = (ch: string): AtomClass =>
  RELATIONS.has(ch) ? 'rel' : BINARIES.has(ch) ? 'bin' : PUNCTUATION.has(ch) ? 'punct' : 'ord';

// ---------------------------------------------------------------------------
// Layout.

interface Env {
  readonly font: MathFont;
  readonly measure: TextMeasurer;
  readonly unsupported: string[];
}

const textBox = (
  env: Env,
  text: string,
  bold: boolean,
  italic: boolean,
  fillHex: string,
  sizePx: number,
): Box => {
  const w = env.measure(text, {
    family: env.font.family,
    sizePx,
    bold,
    italic,
    letterSpacingPx: 0,
  }).widthPx;
  const family = escapeXml(env.font.cssFamily ?? env.font.family);
  return {
    w,
    a: sizePx * STRUT_ASCENT,
    d: sizePx * STRUT_DESCENT,
    draw: (x, y, out) => {
      out.push(
        `<text x="${fmt(x)}" y="${fmt(y)}" font-family="${family}" font-size="${fmt(sizePx)}" fill="${fillHex}"${bold ? ' font-weight="700"' : ''}${italic ? ' font-style="italic"' : ''} xml:space="preserve">${escapeXml(text)}</text>`,
      );
    },
  };
};

// A math run becomes one box per character class change, with TeX's spacing
// around relations and binary operators. Consecutive characters of one style
// share a box so the SVG stays compact.
const runBoxes = (
  env: Env,
  ctx: Ctx,
  node: Extract<MathNode, { kind: 'text' }>,
  previous: AtomClass | null,
): { boxes: Box[]; last: AtomClass | null } => {
  const style = node.format ? env.font.runStyle?.(node.format) : undefined;
  const sizePx =
    style?.sizePx !== undefined ? style.sizePx * (ctx.sizePx / ctx.baseSizePx) : ctx.sizePx;
  const fillHex = style?.fillHex ?? ctx.fillHex;
  const boxes: Box[] = [];
  let chunk = '';
  let chunkBold = false;
  let chunkItalic = false;
  const flush = (): void => {
    if (chunk === '') return;
    boxes.push(textBox(env, chunk, chunkBold, chunkItalic, fillHex, sizePx));
    chunk = '';
  };
  let last = previous;
  for (const raw of node.text) {
    const glyph = glyphOf(raw);
    const cls = classOf(glyph.ch);
    const defaultItalic = !ctx.upright && LETTER.test(glyph.ch);
    const bold = glyph.bold || node.style === 'bold' || node.style === 'boldItalic';
    const italic =
      node.style === 'plain' || node.style === 'bold'
        ? false
        : node.style === 'italic' || node.style === 'boldItalic'
          ? true
          : (glyph.italic ?? defaultItalic);
    // Scripts drop operator spacing, as TeX does.
    const spaced = ctx.level === 0 && cls !== 'ord';
    // A binary operator with nothing to its left is a sign (−x), not an operator.
    const binary = cls === 'bin' && last !== null && last !== 'rel' && last !== 'bin';
    const before = spaced ? (cls === 'rel' ? THICK_SPACE : binary ? MEDIUM_SPACE : 0) : 0;
    const after = spaced
      ? cls === 'rel'
        ? THICK_SPACE
        : binary
          ? MEDIUM_SPACE
          : cls === 'punct'
            ? THIN_SPACE
            : 0
      : 0;
    if (before > 0 || after > 0) {
      flush();
      const box = textBox(env, glyph.ch, bold, italic, fillHex, sizePx);
      boxes.push(padded(box, before * sizePx, after * sizePx));
    } else {
      if (chunk !== '' && (bold !== chunkBold || italic !== chunkItalic)) flush();
      chunk += glyph.ch;
      chunkBold = bold;
      chunkItalic = italic;
    }
    last = cls;
  }
  flush();
  return { boxes, last };
};

const list = (env: Env, ctx: Ctx, nodes: MathNodes): Box => {
  const boxes: Box[] = [];
  let last: AtomClass | null = null;
  for (const node of nodes) {
    if (node.kind === 'text') {
      const run = runBoxes(env, ctx, node, last);
      boxes.push(...run.boxes);
      last = run.last;
    } else {
      boxes.push(layoutNode(env, ctx, node));
      last = 'ord';
    }
  }
  return boxes.length === 0 ? emptyBox() : hlist(boxes);
};

const line = (x1: number, y1: number, x2: number, y2: number, width: number, stroke: string) =>
  `<line x1="${fmt(x1)}" y1="${fmt(y1)}" x2="${fmt(x2)}" y2="${fmt(y2)}" stroke="${stroke}" stroke-width="${fmt(width)}"/>`;

const fraction = (env: Env, ctx: Ctx, node: Extract<MathNode, { kind: 'fraction' }>): Box => {
  const inner = smallerCtx(ctx);
  const num = list(env, inner, node.numerator);
  const den = list(env, inner, node.denominator);
  const s = ctx.sizePx;
  if (node.type === 'linear' || node.type === 'skewed') {
    const slash = textBox(env, '/', false, false, ctx.fillHex, s);
    return node.type === 'linear'
      ? hlist([list(env, ctx, node.numerator), slash, list(env, ctx, node.denominator)])
      : hlist([shifted(num, -0.3 * s), slash, shifted(den, 0.15 * s)]);
  }
  const axis = AXIS * s;
  const rule = node.type === 'noBar' ? 0 : RULE * s;
  const gap = 0.12 * s;
  const pad = 0.1 * s;
  const w = Math.max(num.w, den.w) + 2 * pad;
  const up = axis + rule / 2 + gap + num.d;
  const down = -axis + rule / 2 + gap + den.a;
  return {
    w,
    a: up + num.a,
    d: down + den.d,
    draw: (x, y, out) => {
      num.draw(x + (w - num.w) / 2, y - up, out);
      den.draw(x + (w - den.w) / 2, y + down, out);
      if (rule > 0)
        out.push(line(x + pad / 2, y - axis, x + w - pad / 2, y - axis, rule, ctx.fillHex));
    },
  };
};

const scripts = (
  ctx: Ctx,
  base: Box,
  sub: Box | null,
  sup: Box | null,
  kern = 0.05 * ctx.sizePx,
): Box => {
  const s = ctx.sizePx;
  let up = sup ? Math.max(0.38 * s, base.a - 0.3 * s, sup.d + 0.1 * s) : 0;
  let down = sub ? Math.max(0.2 * s, base.d + 0.05 * s, sub.a - 0.45 * s) : 0;
  if (sub && sup) {
    // Keep a visible gap between the two scripts.
    const clash = 0.1 * s - (up - sup.d - (sub.a - down));
    if (clash > 0) {
      down += clash / 2;
      up += clash / 2;
    }
  }
  const scriptW = Math.max(sub?.w ?? 0, sup?.w ?? 0);
  return {
    w: base.w + kern + scriptW,
    a: Math.max(base.a, sup ? up + sup.a : 0),
    d: Math.max(base.d, sub ? down + sub.d : 0),
    draw: (x, y, out) => {
      base.draw(x, y, out);
      sup?.draw(x + base.w + kern, y - up, out);
      sub?.draw(x + base.w + kern, y + down, out);
    },
  };
};

const radical = (env: Env, ctx: Ctx, node: Extract<MathNode, { kind: 'radical' }>): Box => {
  const s = ctx.sizePx;
  const body = list(env, ctx, node.body);
  const rule = RULE * s;
  const clearance = 0.12 * s;
  const top = body.a + clearance + rule;
  const bottom = Math.max(body.d, 0.1 * s);
  const height = top + bottom;
  const signW = 0.45 * s + 0.08 * height;
  const degree =
    node.degree && node.degree.length > 0
      ? list(env, scriptCtx(scriptCtx(ctx)), node.degree)
      : null;
  const lead = degree ? Math.max(0, degree.w - signW * 0.55) : 0;
  return {
    w: lead + signW + body.w + 0.1 * s,
    a: Math.max(top, degree ? height * 0.6 - bottom + degree.a : 0),
    d: bottom,
    draw: (x, y, out) => {
      const x0 = x + lead;
      const yb = y + bottom;
      const yt = y - top + rule / 2;
      degree?.draw(x0 + signW * 0.55 - degree.w, yb - height * 0.6, out);
      out.push(
        `<path d="M${fmt(x0)} ${fmt(yb - height * 0.42)}L${fmt(x0 + signW * 0.22)} ${fmt(yb - height * 0.5)}L${fmt(x0 + signW * 0.5)} ${fmt(yb)}L${fmt(x0 + signW)} ${fmt(yt)}L${fmt(x0 + signW + body.w + 0.05 * s)} ${fmt(yt)}" fill="none" stroke="${ctx.fillHex}" stroke-width="${fmt(rule)}" stroke-linejoin="round"/>`,
      );
      body.draw(x0 + signW, y, out);
    },
  };
};

const INTEGRALS = new Set('∫∬∭∮∯∰∱∲∳');

const nary = (env: Env, ctx: Ctx, node: Extract<MathNode, { kind: 'nary' }>): Box => {
  const s = ctx.sizePx;
  const opSize = ctx.display ? s * DISPLAY_OPERATOR_SCALE : s;
  const integral = INTEGRALS.has(node.operator);
  const glyph = textBox(env, node.operator, false, false, ctx.fillHex, opSize);
  // Centre the operator's ink (about 0.05..0.75 em above the baseline) on the axis.
  const dy = 0.35 * opSize - AXIS * s;
  const op: Box = {
    w: glyph.w,
    a: 0.8 * opSize - dy,
    d: 0.1 * opSize + dy,
    draw: (x, y, out) => glyph.draw(x, y + dy, out),
  };
  const limitCtx = scriptCtx(ctx);
  const sub = node.sub && node.sub.length > 0 ? list(env, limitCtx, node.sub) : null;
  const sup = node.sup && node.sup.length > 0 ? list(env, limitCtx, node.sup) : null;
  const underOver =
    node.limitLocation === 'underOver' ||
    (node.limitLocation === undefined && ctx.display && !integral);
  const operator = underOver
    ? stacked(ctx, op, sub, sup)
    : scripts(ctx, op, sub, sup, integral ? 0.1 * s : 0.05 * s);
  const body = list(env, ctx, node.body);
  return hlist([operator, padded(body, THIN_SPACE * s, 0)]);
};

// `base` with `under` / `over` centred below / above it.
const stacked = (ctx: Ctx, base: Box, under: Box | null, over: Box | null): Box => {
  const gap = 0.1 * ctx.sizePx;
  const w = Math.max(base.w, under?.w ?? 0, over?.w ?? 0);
  const up = over ? base.a + gap + over.d : 0;
  const down = under ? base.d + gap + under.a : 0;
  return {
    w,
    a: over ? up + over.a : base.a,
    d: under ? down + under.d : base.d,
    draw: (x, y, out) => {
      base.draw(x + (w - base.w) / 2, y, out);
      over?.draw(x + (w - over.w) / 2, y - up, out);
      under?.draw(x + (w - under.w) / 2, y + down, out);
    },
  };
};

// Width of a drawn delimiter, in em of the current size.
const DELIMITER_WIDTH: Readonly<Record<string, number>> = {
  '(': 0.35,
  ')': 0.35,
  '[': 0.3,
  ']': 0.3,
  '{': 0.4,
  '}': 0.4,
  '|': 0.25,
  '‖': 0.35,
  '⟨': 0.35,
  '⟩': 0.35,
  '〈': 0.35,
  '〉': 0.35,
  '⌊': 0.3,
  '⌋': 0.3,
  '⌈': 0.3,
  '⌉': 0.3,
};

// A delimiter stretched to span `top`..`bottom` (offsets from the baseline),
// drawn as strokes so it scales without a font's extensible glyphs.
const delimiterBox = (env: Env, ctx: Ctx, ch: string, top: number, bottom: number): Box => {
  const s = ctx.sizePx;
  if (ch === '') return emptyBox();
  const em = DELIMITER_WIDTH[ch];
  if (em === undefined) {
    return textBox(env, ch, false, false, ctx.fillHex, s);
  }
  const w = em * s;
  const stroke = 0.06 * s;
  const left =
    ch === ')' || ch === ']' || ch === '}' || ch === '⟩' || ch === '〉' || ch === '⌋' || ch === '⌉';
  return {
    w,
    a: top,
    d: bottom,
    draw: (x, y, out) => {
      const yt = y - top + stroke;
      const yb = y + bottom - stroke;
      const ym = (yt + yb) / 2;
      // Outer and inner x for an opening delimiter; mirrored for a closing one.
      const outer = left ? x + w * 0.2 : x + w * 0.8;
      const inner = left ? x + w * 0.8 : x + w * 0.2;
      const mid = (outer + inner) / 2;
      const path = (d: string) =>
        out.push(
          `<path d="${d}" fill="none" stroke="${ctx.fillHex}" stroke-width="${fmt(stroke)}" stroke-linecap="round" stroke-linejoin="round"/>`,
        );
      switch (ch) {
        case '(':
        case ')':
          path(
            `M${fmt(outer)} ${fmt(yt)}Q${fmt(inner - (outer - inner) * 0.3)} ${fmt(ym)} ${fmt(outer)} ${fmt(yb)}`,
          );
          break;
        case '[':
        case ']':
          path(`M${fmt(outer)} ${fmt(yt)}H${fmt(mid)}V${fmt(yb)}H${fmt(outer)}`);
          break;
        case '⌊':
        case '⌋':
          path(`M${fmt(mid)} ${fmt(yt)}V${fmt(yb)}H${fmt(outer)}`);
          break;
        case '⌈':
        case '⌉':
          path(`M${fmt(outer)} ${fmt(yt)}H${fmt(mid)}V${fmt(yb)}`);
          break;
        case '{':
        case '}':
          path(
            `M${fmt(outer)} ${fmt(yt)}Q${fmt(mid)} ${fmt(yt)} ${fmt(mid)} ${fmt(yt + (ym - yt) * 0.3)}V${fmt(ym - (ym - yt) * 0.3)}Q${fmt(mid)} ${fmt(ym)} ${fmt(inner)} ${fmt(ym)}Q${fmt(mid)} ${fmt(ym)} ${fmt(mid)} ${fmt(ym + (yb - ym) * 0.3)}V${fmt(yb - (yb - ym) * 0.3)}Q${fmt(mid)} ${fmt(yb)} ${fmt(outer)} ${fmt(yb)}`,
          );
          break;
        case '|':
          path(`M${fmt(x + w / 2)} ${fmt(yt)}V${fmt(yb)}`);
          break;
        case '‖':
          path(
            `M${fmt(x + w * 0.35)} ${fmt(yt)}V${fmt(yb)}M${fmt(x + w * 0.65)} ${fmt(yt)}V${fmt(yb)}`,
          );
          break;
        default:
          path(`M${fmt(outer)} ${fmt(yt)}L${fmt(inner)} ${fmt(ym)}L${fmt(outer)} ${fmt(yb)}`);
      }
    },
  };
};

const delimiter = (env: Env, ctx: Ctx, node: Extract<MathNode, { kind: 'delimiter' }>): Box => {
  const s = ctx.sizePx;
  const parts: Box[] = [];
  const items = node.items.map((item) => list(env, ctx, item));
  const axis = AXIS * s;
  // Symmetric about the axis, at least one strut tall.
  let half = (STRUT_ASCENT + STRUT_DESCENT) * s * 0.5;
  for (const item of items) half = Math.max(half, item.a - axis, item.d + axis);
  half += 0.05 * s;
  const top = axis + half;
  const bottom = half - axis;
  parts.push(delimiterBox(env, ctx, node.open, top, bottom));
  items.forEach((item, index) => {
    if (index > 0) parts.push(delimiterBox(env, ctx, node.separator, top, bottom));
    parts.push(padded(item, 0.05 * s, 0.05 * s));
  });
  parts.push(delimiterBox(env, ctx, node.close, top, bottom));
  return hlist(parts);
};

const matrix = (env: Env, ctx: Ctx, node: Extract<MathNode, { kind: 'matrix' }>): Box => {
  const inner = smallerCtx(ctx);
  const s = inner.sizePx;
  const cells = node.rows.map((row) => row.map((cell) => list(env, inner, cell)));
  const columns = Math.max(0, ...cells.map((row) => row.length));
  const colW = Array.from({ length: columns }, (_, c) =>
    Math.max(0, ...cells.map((row) => row[c]?.w ?? 0)),
  );
  const rowA = cells.map((row) => Math.max(STRUT_ASCENT * s, ...row.map((cell) => cell.a)));
  const rowD = cells.map((row) => Math.max(STRUT_DESCENT * s, ...row.map((cell) => cell.d)));
  const rowGap = 0.2 * s;
  const colGap = 0.8 * s;
  const height =
    rowA.reduce((sum, a) => sum + a, 0) +
    rowD.reduce((sum, d) => sum + d, 0) +
    rowGap * Math.max(0, cells.length - 1);
  const width = colW.reduce((sum, w) => sum + w, 0) + colGap * Math.max(0, columns - 1);
  const axis = AXIS * ctx.sizePx;
  const top = axis + height / 2;
  return {
    w: width,
    a: top,
    d: height - top,
    draw: (x, y, out) => {
      let rowTop = y - top;
      cells.forEach((row, r) => {
        const baseline = rowTop + rowA[r]!;
        let cellX = x;
        row.forEach((cell, c) => {
          cell.draw(cellX + (colW[c]! - cell.w) / 2, baseline, out);
          cellX += colW[c]! + colGap;
        });
        rowTop = baseline + rowD[r]! + rowGap;
      });
    },
  };
};

// Spacing forms of the combining accents the desktop app offers.
const SPACING_ACCENT: Readonly<Record<string, string>> = {
  '\u0300': '`',
  '\u0301': '´',
  '\u0302': 'ˆ',
  '\u0303': '˜',
  '\u0306': '˘',
  '\u0307': '˙',
  '\u0308': '¨',
  '\u030C': 'ˇ',
};
const BAR_ACCENTS = new Set(['\u0304', '\u0305', '¯']);
const ARROW_ACCENTS = new Set(['\u20D7', '\u20D1', '\u20D6']);

const accent = (env: Env, ctx: Ctx, node: Extract<MathNode, { kind: 'accent' }>): Box => {
  const s = ctx.sizePx;
  const body = list(env, ctx, node.body);
  const top = body.a;
  if (BAR_ACCENTS.has(node.accent) || ARROW_ACCENTS.has(node.accent)) {
    const rise = 0.08 * s;
    const arrow = ARROW_ACCENTS.has(node.accent);
    return {
      w: body.w,
      a: top + rise + (arrow ? 0.12 * s : RULE * s),
      d: body.d,
      draw: (x, y, out) => {
        body.draw(x, y, out);
        const ly = y - top - rise;
        out.push(line(x + 0.05 * s, ly, x + body.w - 0.02 * s, ly, RULE * s, ctx.fillHex));
        if (arrow) {
          const tip = x + body.w - 0.02 * s;
          out.push(
            `<path d="M${fmt(tip - 0.12 * s)} ${fmt(ly - 0.08 * s)}L${fmt(tip)} ${fmt(ly)}L${fmt(tip - 0.12 * s)} ${fmt(ly + 0.08 * s)}" fill="none" stroke="${ctx.fillHex}" stroke-width="${fmt(RULE * s)}"/>`,
          );
        }
      },
    };
  }
  const mark = textBox(
    env,
    SPACING_ACCENT[node.accent] ?? node.accent,
    false,
    false,
    ctx.fillHex,
    s,
  );
  // The spacing accents sit about half an em above their baseline; lift the
  // mark so it clears the body's top.
  const lift = top - 0.42 * s;
  return {
    w: Math.max(body.w, mark.w),
    a: top + 0.25 * s,
    d: body.d,
    draw: (x, y, out) => {
      const w = Math.max(body.w, mark.w);
      body.draw(x + (w - body.w) / 2, y, out);
      mark.draw(x + (w - mark.w) / 2 + 0.05 * s, y - lift, out);
    },
  };
};

const unsupported = (env: Env, ctx: Ctx, node: Extract<MathNode, { kind: 'unsupported' }>): Box => {
  env.unsupported.push(node.element);
  const s = ctx.sizePx;
  const text = textBox(env, node.text === '' ? '?' : node.text, false, false, ctx.fillHex, s);
  const pad = 0.1 * s;
  return {
    w: text.w + 2 * pad,
    a: text.a + pad,
    d: text.d + pad,
    draw: (x, y, out) => {
      out.push(
        `<rect x="${fmt(x + pad / 2)}" y="${fmt(y - text.a - pad / 2)}" width="${fmt(text.w + pad)}" height="${fmt(text.a + text.d + pad)}" fill="none" stroke="${UNSUPPORTED_STROKE}" stroke-width="${fmt(Math.max(1, 0.04 * s))}" stroke-dasharray="${fmt(0.15 * s)} ${fmt(0.1 * s)}"><title>Unsupported equation element: ${escapeXml(node.element)}</title></rect>`,
      );
      text.draw(x + pad, y, out);
    },
  };
};

const layoutNode = (env: Env, ctx: Ctx, node: Exclude<MathNode, { kind: 'text' }>): Box => {
  switch (node.kind) {
    case 'fraction':
      return fraction(env, ctx, node);
    case 'subscript':
      return scripts(ctx, list(env, ctx, node.base), list(env, scriptCtx(ctx), node.sub), null);
    case 'superscript':
      return scripts(ctx, list(env, ctx, node.base), null, list(env, scriptCtx(ctx), node.sup));
    case 'subSuperscript':
      return scripts(
        ctx,
        list(env, ctx, node.base),
        list(env, scriptCtx(ctx), node.sub),
        list(env, scriptCtx(ctx), node.sup),
      );
    case 'radical':
      return radical(env, ctx, node);
    case 'nary':
      return nary(env, ctx, node);
    case 'delimiter':
      return delimiter(env, ctx, node);
    case 'matrix':
      return matrix(env, ctx, node);
    case 'accent':
      return accent(env, ctx, node);
    case 'function':
      return hlist([
        list(env, { ...ctx, upright: true }, node.name),
        padded(list(env, ctx, node.body), THIN_SPACE * ctx.sizePx, 0),
      ]);
    case 'lowerLimit':
    case 'upperLimit': {
      const limit = list(env, scriptCtx(ctx), node.limit);
      const base = list(env, ctx, node.base);
      return node.kind === 'lowerLimit'
        ? stacked(ctx, base, limit, null)
        : stacked(ctx, base, null, limit);
    }
    case 'unsupported':
      return unsupported(env, ctx, node);
  }
};

/** Lays out one equation. `display` selects display style (`<m:oMathPara>`). */
export const layoutMath = (
  nodes: MathNodes,
  display: boolean,
  font: MathFont,
  measure: TextMeasurer,
): InlineBox => {
  const env: Env = { font, measure, unsupported: [] };
  const ctx: Ctx = {
    sizePx: font.sizePx,
    baseSizePx: font.sizePx,
    fillHex: font.fillHex,
    display,
    level: 0,
    upright: false,
  };
  const box = list(env, ctx, nodes);
  return {
    width: box.w,
    ascent: box.a,
    descent: box.d,
    svg: (x, baselineY) => {
      const out: string[] = [];
      box.draw(x, baselineY, out);
      return `<g data-pptx-math="">${out.join('')}</g>`;
    },
    // Local names of OMML elements drawn as plain text because they are not modeled.
    unsupported: env.unsupported,
  };
};

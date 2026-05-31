// Unit coverage for the pure-SVG text layout engine. It is dependency-free
// (no pptx-kit, no Node, no browser), so the wrap / position / emit logic and
// the font substitution map are testable in isolation. See
// site/src/lib/playground/text-layout.ts.

import { describe, expect, it } from 'vitest';
import {
  defaultMeasurer,
  layoutTextSvg,
  substituteFamily,
  type FontSpec,
  type ParaInput,
  type PieceInput,
  type TextBodyInput,
  type TextMeasurer,
} from '../site/src/lib/playground/text-layout.ts';

describe('substituteFamily', () => {
  it('maps PowerPoint fonts to bundled internal family names', () => {
    expect(substituteFamily('Calibri')).toBe('Carlito');
    expect(substituteFamily('Calibri Light')).toBe('Carlito');
    expect(substituteFamily('Aptos')).toBe('Carlito');
    expect(substituteFamily('Cambria')).toBe('Caladea');
    expect(substituteFamily('Arial')).toBe('Liberation Sans');
    expect(substituteFamily('Times New Roman')).toBe('Liberation Serif');
    expect(substituteFamily('Courier New')).toBe('Liberation Mono');
  });

  it('falls back to the sans substitute for unknown or missing families', () => {
    expect(substituteFamily('Some Brand Font')).toBe('Carlito');
    expect(substituteFamily(null)).toBe('Carlito');
    expect(substituteFamily(undefined)).toBe('Carlito');
  });
});

describe('defaultMeasurer', () => {
  const spec = (over: Partial<FontSpec> = {}): FontSpec => ({
    family: 'Carlito',
    sizePx: 10,
    bold: false,
    italic: false,
    letterSpacingPx: 0,
    ...over,
  });

  it('estimates width from the glyph-width ratio', () => {
    // 0.55 em per Latin glyph: "AB" at 10px ⇒ 11.
    expect(defaultMeasurer('AB', spec()).widthPx).toBeCloseTo(11, 5);
  });

  it('adds letter-spacing between glyphs only (n-1 gaps)', () => {
    expect(defaultMeasurer('AB', spec({ letterSpacingPx: 2 })).widthPx).toBeCloseTo(13, 5);
    expect(defaultMeasurer('A', spec({ letterSpacingPx: 2 })).widthPx).toBeCloseTo(5.5, 5);
  });

  it('widens CJK glyphs to ~1em', () => {
    expect(defaultMeasurer('日', spec()).widthPx).toBeCloseTo(10, 5);
  });
});

// A deterministic measurer: every glyph is `sizePx` wide; fixed vertical metrics.
const stubMeasurer: TextMeasurer = (text, s) => ({
  widthPx: [...text].length * s.sizePx,
  ascentPx: s.sizePx * 0.8,
  descentPx: s.sizePx * 0.2,
  lineGapPx: 0,
});

const piece = (text: string, over: Partial<PieceInput> = {}): PieceInput => ({
  text,
  family: 'Carlito',
  sizePx: 10,
  bold: false,
  italic: false,
  letterSpacingPx: 0,
  fillHex: '#000000',
  underline: false,
  strike: false,
  superSub: 0,
  href: null,
  isBreak: false,
  ...over,
});

const para = (pieces: PieceInput[], over: Partial<ParaInput> = {}): ParaInput => ({
  align: 'left',
  marLpx: 0,
  marRpx: 0,
  firstIndentPx: 0,
  spcBefPx: 0,
  spcAftPx: 0,
  lineSpacing: null,
  lineAdvanceScale: 1,
  bullet: null,
  pieces,
  fallbackSizePx: 10,
  ...over,
});

const body = (paragraphs: ParaInput[], over: Partial<TextBodyInput> = {}): TextBodyInput => ({
  boxXpx: 0,
  boxYpx: 0,
  boxWpx: 1000,
  boxHpx: 200,
  anchor: 'top',
  wrap: true,
  paragraphs,
  ...over,
});

const countText = (svg: string): number => (svg.match(/<text /g) ?? []).length;

describe('layoutTextSvg', () => {
  it('emits one <text> for a single line, left-anchored at the box edge', () => {
    const svg = layoutTextSvg(body([para([piece('Hello')])]), stubMeasurer);
    expect(countText(svg)).toBe(1);
    expect(svg).toContain('text-anchor="start"');
    expect(svg).toContain('>Hello</tspan>');
    // baseline = boxY(0) + ascent(8) for a top-anchored first line.
    expect(svg).toContain('y="8"');
  });

  it('uses text-anchor=middle / end for centered / right alignment', () => {
    expect(layoutTextSvg(body([para([piece('x')], { align: 'center' })]), stubMeasurer)).toContain(
      'text-anchor="middle"',
    );
    expect(layoutTextSvg(body([para([piece('x')], { align: 'right' })]), stubMeasurer)).toContain(
      'text-anchor="end"',
    );
  });

  it('wraps to the inner width', () => {
    // 4 words × 10px/glyph; narrow box forces multiple lines.
    const svg = layoutTextSvg(body([para([piece('aa bb cc dd')])], { boxWpx: 40 }), stubMeasurer);
    expect(countText(svg)).toBeGreaterThan(1);
  });

  it('honors explicit breaks as new lines', () => {
    const svg = layoutTextSvg(
      body([para([piece('a'), piece('', { isBreak: true }), piece('b')])]),
      stubMeasurer,
    );
    expect(countText(svg)).toBe(2);
  });

  it('does not wrap when wrap is disabled', () => {
    const svg = layoutTextSvg(
      body([para([piece('aa bb cc dd')])], { boxWpx: 40, wrap: false }),
      stubMeasurer,
    );
    expect(countText(svg)).toBe(1);
  });

  it('emits run styling as tspan attributes', () => {
    const svg = layoutTextSvg(
      body([para([piece('B', { bold: true, italic: true, underline: true, fillHex: '#FF0000' })])]),
      stubMeasurer,
    );
    expect(svg).toContain('font-weight="700"');
    expect(svg).toContain('font-style="italic"');
    expect(svg).toContain('text-decoration="underline"');
    expect(svg).toContain('fill="#FF0000"');
  });

  it('renders a bullet glyph ahead of the first line', () => {
    const svg = layoutTextSvg(
      body([
        para([piece('item')], {
          bullet: { text: '•', family: 'Carlito', sizePx: 10, fillHex: '#000000' },
        }),
      ]),
      stubMeasurer,
    );
    expect(svg).toContain('>•</text>');
    expect(countText(svg)).toBe(2); // bullet + line
  });

  it('centers the block for bottom / center vertical anchor', () => {
    const top = layoutTextSvg(body([para([piece('x')])], { anchor: 'top' }), stubMeasurer);
    const bottom = layoutTextSvg(body([para([piece('x')])], { anchor: 'bottom' }), stubMeasurer);
    const yOf = (s: string): number => Number(/y="([\d.]+)"/.exec(s)?.[1] ?? '0');
    expect(yOf(bottom)).toBeGreaterThan(yOf(top));
  });
});

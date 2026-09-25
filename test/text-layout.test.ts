// Unit coverage for the pure-SVG text layout engine. It is dependency-free
// (no @office-kit/pptx, no Node, no browser), so the wrap / position / emit logic and
// the font substitution map are testable in isolation. See
// packages/preview/src/text-layout.ts.

import { describe, expect, it } from 'vitest';
import {
  defaultMeasurer,
  layoutTextSvg,
  layoutCore,
  substituteFamily,
  type FontSpec,
  type ParaInput,
  type PieceInput,
  type TextBodyInput,
  type TextMeasurer,
} from '../packages/preview/src/text-layout.ts';

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
  underline: 'none',
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
  it.each(['left', 'center', 'right'] as const)(
    'positions run highlights behind %s-aligned text',
    (align) => {
      const svg = layoutTextSvg(
        body([para([piece('A'), piece('B', { highlightHex: '#ffff00' }), piece('C')], { align })], {
          boxWpx: 100,
        }),
        stubMeasurer,
      );
      const rect =
        /<rect x="([^"]+)" y="([^"]+)" width="([^"]+)" height="([^"]+)" fill="#ffff00"\/>/.exec(
          svg,
        )!;
      expect(rect).not.toBeNull();
      expect(Number(rect[1])).toBeCloseTo(
        (align === 'left' ? 10 : align === 'center' ? 45 : 80) - 0.75,
        1,
      );
      expect(Number(rect[3])).toBe(10);
      expect(Number(rect[4])).toBe(10);
      expect(svg.indexOf('<rect')).toBeLessThan(svg.indexOf('<text'));
      expect(svg).toContain('>A</tspan>');
      expect(svg).toContain('>B</tspan>');
    },
  );
  it.each([1, -1] as const)('measures highlighted script %s at its rendered size', (superSub) => {
    const svg = layoutTextSvg(
      body(
        [
          para([piece('A'), piece('B', { highlightHex: '#ffff00', superSub }), piece('C')], {
            align: 'center',
          }),
        ],
        { boxWpx: 100 },
      ),
      stubMeasurer,
    );
    const rect =
      /<rect x="([^"]+)" y="([^"]+)" width="([^"]+)" height="([^"]+)" fill="#ffff00"/.exec(svg)!;
    expect(Number(rect[1])).toBe(46);
    expect(Number(rect[3])).toBe(6.5);
    expect(Number(rect[4])).toBe(6.5);
  });
  it('splits highlights across wrapped lines and leaves plain text without backgrounds', () => {
    const svg = layoutTextSvg(
      body([para([piece('ABCD', { highlightHex: '#00ff00' })])], { boxWpx: 20 }),
      stubMeasurer,
    );
    expect((svg.match(/fill="#00ff00"/g) ?? []).length).toBe(2);
    expect(layoutTextSvg(body([para([piece('plain')])]), stubMeasurer)).not.toContain('<rect');
  });

  it('emits one <text> for a single line, left-anchored at the box edge', () => {
    const svg = layoutTextSvg(body([para([piece('Hello')])]), stubMeasurer);
    expect(countText(svg)).toBe(1);
    expect(svg).toContain('text-anchor="start"');
    expect(svg).toContain('>Hello</tspan>');
    // baseline = boxY(0) + ascent(8) + the first-line leading drop
    // (0.036·(ascent+descent) = 0.036·10 = 0.36), applied for all anchors.
    expect(svg).toContain('y="8.36"');
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
      body([
        para([piece('B', { bold: true, italic: true, underline: 'single', fillHex: '#FF0000' })]),
      ]),
      stubMeasurer,
    );
    expect(svg).toContain('font-weight="700"');
    expect(svg).toContain('font-style="italic"');
    expect(svg).toContain('text-decoration="underline"');
    expect(svg).toContain('fill="#FF0000"');
  });

  it('draws wavy underline as a path, not text-decoration (resvg has no text-decoration-style)', () => {
    const svg = layoutTextSvg(
      body([para([piece('wavy', { underline: 'wavy', fillHex: '#0000FF' })])]),
      stubMeasurer,
    );
    expect(svg).not.toContain('text-decoration');
    expect(svg).toContain('<path');
    expect(svg).toContain('stroke="#0000FF"');
  });

  it('scales the wavy-underline path down for a superscript run, matching its shrunk glyph size', () => {
    // tspan() renders a super/subscript run's glyphs at 0.65x sizePx — the
    // wave must scale down the same way, or it reads oversized under the
    // smaller glyphs it's supposed to underline.
    const strokeWidthOf = (svg: string): number => {
      const m = /<path d="[^"]*" stroke="[^"]*" stroke-width="([\d.]+)"/.exec(svg);
      if (!m) throw new Error('no wavy-underline path found');
      return Number(m[1]);
    };
    const normal = layoutTextSvg(
      body([para([piece('wavy', { underline: 'wavy', sizePx: 20 })])]),
      stubMeasurer,
    );
    const superscript = layoutTextSvg(
      body([para([piece('wavy', { underline: 'wavy', sizePx: 20, superSub: 1 })])]),
      stubMeasurer,
    );
    expect(strokeWidthOf(superscript)).toBeLessThan(strokeWidthOf(normal));
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

// X/Y of every emitted <text> in document order (bullets included).
const coords = (svg: string): { x: number; y: number }[] =>
  [...svg.matchAll(/<text x="(-?[\d.]+)" y="(-?[\d.]+)"/g)].map((m) => ({
    x: Number(m[1]),
    y: Number(m[2]),
  }));

const br = (): PieceInput => piece('', { isBreak: true });

describe('layoutTextSvg vertical text', () => {
  // Box 100×200 ⇒ centre (50,100); the rotation pivots there so the swapped
  // 200×100 layout frame lands back on the shape after rotation.
  const vbox = { boxXpx: 0, boxYpx: 0, boxWpx: 100, boxHpx: 200 } as const;

  it('cw90 wraps output in a rotate(90) transform about the box centre', () => {
    const svg = layoutTextSvg(body([para([piece('A')])], { ...vbox, vert: 'cw90' }), stubMeasurer);
    expect(svg).toMatch(/^<g transform="rotate\(90 50 100\)">/);
    expect(svg.endsWith('</g>')).toBe(true);
  });

  it('cw270 rotates the opposite way', () => {
    const svg = layoutTextSvg(body([para([piece('A')])], { ...vbox, vert: 'cw270' }), stubMeasurer);
    expect(svg).toMatch(/^<g transform="rotate\(270 50 100\)">/);
  });

  it('stacks successive lines along the local axis (increasing y before rotation)', () => {
    const svg = layoutTextSvg(
      body([para([piece('A'), br(), piece('B')])], { ...vbox, vert: 'cw90' }),
      stubMeasurer,
    );
    const ys = coords(svg).map((c) => c.y);
    expect(ys).toHaveLength(2);
    // Lines stack in local +y; the rotate transform turns that into the
    // right-to-left column stacking that vert text shows visually.
    expect(ys[1]!).toBeGreaterThan(ys[0]!);
  });

  it('ignores numCol for vertical text (rotation wins over columns)', () => {
    const svg = layoutTextSvg(
      body([para([piece('A'), br(), piece('B')])], {
        ...vbox,
        vert: 'cw90',
        columns: { count: 2, gapPx: 10 },
      }),
      stubMeasurer,
    );
    // A single x-offset ⇒ no column split was applied.
    expect(new Set(coords(svg).map((c) => c.x)).size).toBe(1);
  });
});

describe('layoutTextSvg multi-column', () => {
  it('sequentially fills column 1, then spills overflow into column 2 at the gap offset', () => {
    // 5 lines, each 10px tall (stub metrics); box height 30 fits 3 per column.
    const fiveLines = para([
      piece('A'),
      br(),
      piece('B'),
      br(),
      piece('C'),
      br(),
      piece('D'),
      br(),
      piece('E'),
    ]);
    const svg = layoutTextSvg(
      body([fiveLines], { boxWpx: 200, boxHpx: 30, columns: { count: 2, gapPx: 10 } }),
      stubMeasurer,
    );
    const xs = coords(svg).map((c) => c.x);
    // colW = (200 − 1·10) / 2 = 95; column 2 sits at 95 + 10 (gap) = 105.
    // GRID_NUDGE_X (−0.75) shifts every emitted x — see text-layout.ts.
    expect(new Set(xs)).toEqual(new Set([-0.75, 104.25]));
    expect(xs.filter((x) => x === 104.25)).toHaveLength(2); // D, E overflowed
  });

  it('wraps a new row of columns once every column in the current row is full', () => {
    // 7 lines, 10px each; box height 30 fits 3 per column — filling BOTH
    // columns of the first row, so the 7th line must start a fresh row
    // rather than pile up past column 2's capacity.
    const pieces: PieceInput[] = [];
    for (const ch of ['A', 'B', 'C', 'D', 'E', 'F', 'G']) {
      if (pieces.length > 0) pieces.push(br());
      pieces.push(piece(ch));
    }
    const svg = layoutTextSvg(
      body([para(pieces)], { boxWpx: 200, boxHpx: 30, columns: { count: 2, gapPx: 10 } }),
      stubMeasurer,
    );
    const cs = coords(svg);
    expect(cs).toHaveLength(7);
    // Row 1: A–C in column 1 (x=-0.75), D–F in column 2 (x=104.25) — same as
    // the single-row case above, y unaffected by the second column.
    expect(cs.slice(0, 3).map((c) => c.x)).toEqual([-0.75, -0.75, -0.75]);
    expect(cs.slice(3, 6).map((c) => c.x)).toEqual([104.25, 104.25, 104.25]);
    // Row 2: G wraps back to column 1, offset down by the box height (30) —
    // NOT stacked into column 2 past its 3-line capacity.
    expect(cs[6]!.x).toBe(-0.75);
    expect(cs[6]!.y).toBe(cs[0]!.y + 30);
  });

  it('keeps few lines entirely in the first column', () => {
    const svg = layoutTextSvg(
      body([para([piece('A'), br(), piece('B')])], {
        boxWpx: 200,
        boxHpx: 100,
        columns: { count: 2, gapPx: 10 },
      }),
      stubMeasurer,
    );
    expect(new Set(coords(svg).map((c) => c.x))).toEqual(new Set([-0.75]));
  });
});

describe('layoutTextSvg horizontal parity', () => {
  it('omitting vert/columns is identical to passing none/null', () => {
    const paras = [para([piece('hello world foo bar baz')])];
    const plain = layoutTextSvg(body(paras, { boxWpx: 60 }), stubMeasurer);
    const explicit = layoutTextSvg(
      body(paras, { boxWpx: 60, vert: 'none', columns: null }),
      stubMeasurer,
    );
    expect(plain).toBe(explicit);
  });

  it('emits the calibrated horizontal layout unchanged (snapshot guard)', () => {
    const svg = layoutTextSvg(
      body([para([piece('aa bb cc dd')]), para([piece('ee ff')], { align: 'center' })], {
        boxWpx: 40,
        anchor: 'center',
      }),
      stubMeasurer,
    );
    // Snapshot reflects: the GRID_NUDGE_X horizontal calibration (every x
    // shifts −0.75px), the first-line leading drop (+0.36 on every baseline),
    // and space-inclusive line breaking — "aa bb" (50px incl. space) no longer
    // fits the 40px box, so each word wraps to its own line. See the fidelity
    // calibration notes in site/fidelity/README.md.
    // The `<g data-pptx-paragraph>` wrappers group each paragraph's lines for
    // a build player; they carry no transform, so every coordinate above is
    // the one this guard was calibrated against.
    expect(svg).toMatchInlineSnapshot(
      `"<g data-pptx-paragraph="0"><text x="-0.75" y="78.36" text-anchor="start" xml:space="preserve"><tspan font-family="Carlito" font-size="10" fill="#000000">aa</tspan></text><text x="-0.75" y="88.36" text-anchor="start" xml:space="preserve"><tspan font-family="Carlito" font-size="10" fill="#000000">bb</tspan></text><text x="-0.75" y="98.36" text-anchor="start" xml:space="preserve"><tspan font-family="Carlito" font-size="10" fill="#000000">cc</tspan></text><text x="-0.75" y="108.36" text-anchor="start" xml:space="preserve"><tspan font-family="Carlito" font-size="10" fill="#000000">dd</tspan></text></g><g data-pptx-paragraph="1"><text x="19.25" y="118.36" text-anchor="middle" xml:space="preserve"><tspan font-family="Carlito" font-size="10" fill="#000000">ee</tspan></text><text x="19.25" y="128.36" text-anchor="middle" xml:space="preserve"><tspan font-family="Carlito" font-size="10" fill="#000000">ff</tspan></text></g>"`,
    );
  });
});

describe('centered text bounds', () => {
  it('preserves the full paragraph frame with opposing alignment', () => {
    const paragraphs = [para([piece('ABC')]), para([piece('A')], { align: 'right' })];
    const normal = layoutCore(body(paragraphs, { boxWpx: 100 }), stubMeasurer);
    const centered = layoutCore(
      body(paragraphs, { boxWpx: 100, anchorCentered: true }),
      stubMeasurer,
    );
    expect(centered.anchorShift).toBe(0);
    expect(centered.placements).toEqual(normal.placements);
  });
  it.each(['top', 'center', 'bottom'] as const)(
    'centers a left-aligned block independently of %s anchoring',
    (anchor) => {
      const paragraphs = [para([piece('ABCD')]), para([piece('A')])];
      const normal = layoutCore(
        body(paragraphs, { boxXpx: 20, boxWpx: 100, anchor }),
        stubMeasurer,
      );
      const centered = layoutCore(
        body(paragraphs, { boxXpx: 20, boxWpx: 100, anchor, anchorCentered: true }),
        stubMeasurer,
      );
      expect(centered.placements.map((p) => p.line.anchorX + p.dx)).toEqual([50, 50]);
      expect(centered.placements.map((p) => p.baselineY)).toEqual(
        normal.placements.map((p) => p.baselineY),
      );
      expect(centered.placements.map((p) => p.line.textAnchor)).toEqual(['start', 'start']);
    },
  );
  it('includes a hanging bullet and ignores trailing spaces', () => {
    const line = para([piece('AB  ')], {
      firstIndentPx: -10,
      bullet: { text: '•', family: 'Carlito', sizePx: 10, fillHex: '#000000' },
    });
    const {
      placements: [p],
    } = layoutCore(body([line], { boxWpx: 100, anchorCentered: true }), stubMeasurer);
    // Bullet at 30, a ten-pixel gap, then two ten-pixel glyphs to 70.
    expect(p!.line.bullet!.x + p!.dx).toBe(30);
    expect(p!.line.anchorX + p!.dx).toBe(50);
  });
  it('centers the reading dimension before rotating vertical text', () => {
    const {
      placements: [p],
    } = layoutCore(
      body([para([piece('AB')])], { boxWpx: 100, boxHpx: 200, vert: 'cw90', anchorCentered: true }),
      stubMeasurer,
    );
    expect(p!.line.anchorX + p!.dx).toBe(40);
  });
});

describe('paragraph tab stops', () => {
  it.each([
    ['left', 80],
    ['center', 55],
    ['right', 30],
    ['decimal', 50],
  ] as const)('aligns a mixed-format tab field using %s alignment', (alignment, gap) => {
    const input = body([
      para([piece('AB\t12'), piece('3.4', { bold: true })], {
        tabStops: [{ positionPx: 100, alignment }],
      }),
    ]);
    const lines = layoutCore(input, stubMeasurer).placements.map((placement) => placement.line);
    expect(lines[0]!.tokens.find((token) => token.isTab)?.width).toBe(gap);
    expect(layoutTextSvg(input, stubMeasurer)).toContain(`dx="${gap}"`);
  });

  it('uses default stops relative to the paragraph margin, accounting for first-line indent', () => {
    const input = body([
      para([piece('A\tB\tC')], { marLpx: 30, firstIndentPx: 10, defaultTabSizePx: 50 }),
    ]);
    const lines = layoutCore(input, stubMeasurer).placements.map((placement) => placement.line);
    expect(lines[0]!.tokens.filter((token) => token.isTab).map((token) => token.width)).toEqual([
      30, 40,
    ]);
  });

  it('resets tab positions after a line break and right-aligns decimal fields without a point', () => {
    const input = body([
      para([piece('A\t12'), piece('', { isBreak: true }), piece('B\t123')], {
        tabStops: [{ positionPx: 100, alignment: 'decimal' }],
      }),
    ]);
    const lines = layoutCore(input, stubMeasurer).placements.map((placement) => placement.line);
    expect(lines.map((line) => line.tokens.find((token) => token.isTab)?.width)).toEqual([70, 60]);
  });
});

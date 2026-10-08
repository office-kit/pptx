import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  inches,
  setShapeTextFormat,
} from '../src/api/index.ts';
import { renderSlideToRgba } from '../packages/preview/src/node.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import type { Color } from '../src/internal/drawingml/color.ts';
import {
  defaultMeasurer,
  layoutTextSvg,
  type PieceInput,
  type TextBodyInput,
} from '../packages/preview/src/text-layout.ts';

const piece = (text: string, innerShadow?: PieceInput['innerShadow']): PieceInput => ({
  text,
  family: 'Carlito',
  sizePx: 32,
  bold: false,
  italic: false,
  letterSpacingPx: 0,
  fillHex: '#336699',
  underline: 'none',
  strike: false,
  baseline: 0,
  href: null,
  isBreak: false,
  ...(innerShadow ? { innerShadow } : {}),
});

const body = (pieces: PieceInput[]): TextBodyInput => ({
  boxXpx: 0,
  boxYpx: 0,
  boxWpx: 240,
  boxHpx: 100,
  anchor: 'top',
  wrap: false,
  paragraphs: [
    {
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
      fallbackSizePx: 32,
    },
  ],
});

describe('character inner shadow in preview text', () => {
  it('uses a glyph-alpha mask and isolates mixed formatted runs', () => {
    const svg = layoutTextSvg(
      body([
        piece('shadow', {
          color: '#000000',
          blurEmu: 63500,
          offsetEmu: 50800,
          // Matches test/fixtures/native/text-art-background2-inner-shadow-shape.xml
          // (dir="13500000", in DrawingML's 1/60000-degree units = 225°).
          angleDeg: 225,
          opacity: 0.7,
        }),
        piece(' plain'),
      ]),
      defaultMeasurer,
    );
    expect((svg.match(/<filter id="text-inner-shadow-/g) ?? []).length).toBe(1);
    expect(svg).toContain('in="SourceAlpha"');
    expect(svg).toContain('in2="SourceAlpha" operator="arithmetic"');
    expect(svg).toContain('flood-color="#000000"');
    expect(svg).toContain('plain</tspan>');
    expect(svg).not.toContain('<feDropShadow');
    expect(svg).toMatch(
      /<text[^>]*filter="url\(#text-inner-shadow-[^"]+"[^>]*aria-hidden="true"[^>]*pointer-events="none"/,
    );
  });

  it('renders character inner shadow through foreignObject and raster output', () => {
    const make = (withShadow: boolean, shadowColor: Color = '#000000') => {
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      const shape = addSlideTextBox(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(1.5),
        text: 'Native inner shadow',
      });
      if (withShadow) {
        setShapeTextFormat(
          shape,
          {
            innerShadow: {
              color: shadowColor,
              blurEmu: 63500,
              offsetEmu: 50800,
              angleDeg: 225,
              opacity: 0.7,
            },
          },
          { range: { start: 0, end: 6 } },
        );
      }
      return { pres, slide };
    };
    const shadowDeck = make(true, 'scheme:accent1');
    const svg = renderSlideToSvg(shadowDeck.pres, shadowDeck.slide, {
      textLayout: 'foreignObject',
    });
    expect(svg).toContain('<foreignObject');
    expect(svg).toContain('text-inner-shadow-');
    expect(svg).not.toContain('flood-color="scheme:accent1"');
    const filterStart = svg.indexOf('<filter id="text-inner-shadow-');
    const filterEnd = svg.indexOf('</filter>', filterStart);
    expect(svg.indexOf('<foreignObject')).toBeLessThan(filterStart);
    expect(filterStart).toBeGreaterThanOrEqual(0);
    expect(svg.slice(filterStart, filterEnd)).not.toContain('SourceGraphic');
    const plainDeck = make(false);
    const plain = renderSlideToRgba(plainDeck.pres, plainDeck.slide, { width: 640 }).image.data;
    const shadow = renderSlideToRgba(shadowDeck.pres, shadowDeck.slide, { width: 640 }).image.data;
    let changed = 0;
    let changedOutsideInk = 0;
    for (let i = 0; i < plain.length; i += 4) {
      if (
        plain[i]! !== shadow[i]! ||
        plain[i + 1]! !== shadow[i + 1]! ||
        plain[i + 2]! !== shadow[i + 2]!
      ) {
        changed += 1;
        if (plain[i]! > 248 && plain[i + 1]! > 248 && plain[i + 2]! > 248) changedOutsideInk += 1;
      }
    }
    expect(changed).toBeGreaterThan(100);
    expect(changedOutsideInk).toBe(0);
  });

  it('keeps an explicitly empty shadow distinct from an omitted effect', () => {
    const svg = layoutTextSvg(body([piece('shadow', {}), piece(' plain')]), defaultMeasurer);
    expect((svg.match(/<filter id="text-inner-shadow-/g) ?? []).length).toBe(1);
  });
});

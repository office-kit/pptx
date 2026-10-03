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
import {
  defaultMeasurer,
  layoutTextSvg,
  type PieceInput,
  type TextBodyInput,
} from '../packages/preview/src/text-layout.ts';

const piece = (text: string, effects: Pick<PieceInput, 'shadow' | 'glow'> = {}): PieceInput => ({
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
  ...effects,
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

describe('character shadow and glow in SVG text', () => {
  it('emits glyph-only filters for mixed runs with resolved defaults', () => {
    const svg = layoutTextSvg(
      body([
        piece('shadow', {
          shadow: { color: '#000000', offsetEmu: 38100, blurEmu: 50800, angleDeg: 45 },
        }),
        piece(' glow', { glow: { color: '#00FF00', radiusEmu: 63500 } }),
        piece(' plain'),
      ]),
      defaultMeasurer,
    );
    expect((svg.match(/<filter id="text-outer-shadow-/g) ?? []).length).toBe(1);
    expect((svg.match(/<filter id="text-glow-/g) ?? []).length).toBe(1);
    expect(svg).toContain('feGaussianBlur in="SourceAlpha"');
    expect(svg).toContain('feOffset in="shadowBlur"');
    expect(svg).not.toContain('feDropShadow');
    expect(svg).not.toContain('in="SourceGraphic"');
    expect(svg).toContain('feMorphology');
    expect(svg).toContain('flood-color="#00FF00"');
    expect(svg).toMatch(
      /<text[^>]*filter="url\(#text-outer-shadow-[^"]+"[^>]*aria-hidden="true"[^>]*pointer-events="none"/,
    );
    expect(svg).toContain('plain</tspan>');
  });

  it('renders outer effects in the SVG path and changes raster pixels', () => {
    const make = (kind: 'none' | 'shadow' | 'glow' | 'both') => {
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      const shape = addSlideTextBox(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(1.5),
        text: 'Outer effects',
      });
      if (kind !== 'none') {
        setShapeTextFormat(shape, {
          ...(kind === 'shadow' || kind === 'both'
            ? { shadow: { color: '#000000', blurEmu: 50800, offsetEmu: 38100, angleDeg: 45 } }
            : {}),
          ...(kind === 'glow' || kind === 'both'
            ? { glow: { color: 'scheme:accent1', radiusEmu: 63500 } }
            : {}),
        });
      }
      return { pres, slide };
    };
    const effected = make('both');
    const svg = renderSlideToSvg(effected.pres, effected.slide, { textLayout: 'svg' });
    expect(svg).toContain('text-outer-shadow-');
    expect(svg).toContain('text-glow-');
    expect(svg).not.toContain('flood-color="scheme:accent1"');
    const plainDeck = make('none');
    const plain = renderSlideToRgba(plainDeck.pres, plainDeck.slide, { width: 640 }).image.data;
    const rendered = renderSlideToRgba(effected.pres, effected.slide, { width: 640 }).image.data;
    let changed = 0;
    for (let i = 0; i < plain.length; i += 4) {
      if (
        plain[i] !== rendered[i] ||
        plain[i + 1] !== rendered[i + 1] ||
        plain[i + 2] !== rendered[i + 2]
      ) {
        changed += 1;
      }
    }
    expect(changed).toBeGreaterThan(100);
    for (const kind of ['shadow', 'glow'] as const) {
      const single = make(kind);
      const raster = renderSlideToRgba(single.pres, single.slide, { width: 640 }).image.data;
      let singleChanged = 0;
      for (let i = 0; i < plain.length; i += 4) {
        if (
          plain[i] !== raster[i] ||
          plain[i + 1] !== raster[i + 1] ||
          plain[i + 2] !== raster[i + 2]
        ) {
          singleChanged += 1;
        }
      }
      expect(singleChanged).toBeGreaterThan(20);
    }
  });

  it('keeps zero-opacity shadows inert and large single-glyph shadows inside the filter region', () => {
    const zero = layoutTextSvg(
      body([piece('A', { shadow: { color: '#000000', opacity: 0 } })]),
      defaultMeasurer,
    );
    const plain = layoutTextSvg(body([piece('A')]), defaultMeasurer);
    expect(zero.replace(/text-outer-shadow-[^"]+/g, 'shadow')).not.toBe(plain);
    expect(zero).toContain('flood-opacity="0.000"');

    const makeRaster = (opacity: number | null) => {
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      const shape = addSlideTextBox(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(2),
        h: inches(1),
        text: 'A',
      });
      if (opacity !== null) {
        setShapeTextFormat(shape, {
          shadow: { color: '#000000', opacity, offsetEmu: 38100, blurEmu: 50800 },
        });
      }
      return renderSlideToRgba(pres, slide, { width: 320 }).image.data;
    };
    expect(makeRaster(0)).toEqual(makeRaster(null));

    const largeDeck = createPresentation();
    const largeSlide = addBlankSlide(largeDeck);
    const largeShape = addSlideTextBox(largeSlide, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: 'A',
    });
    setShapeTextFormat(largeShape, {
      shadow: { color: '#FF0000', offsetEmu: 914400, blurEmu: 0, angleDeg: 0 },
    });
    const largeSvg = renderSlideToSvg(largeDeck, largeSlide, { textLayout: 'svg' });
    expect(largeSvg).toContain('filterUnits="userSpaceOnUse"');
    expect(largeSvg).toMatch(/dx="96(?:\.00)?"/);
    const largeRaster = renderSlideToRgba(largeDeck, largeSlide, { width: 960 }).image;
    const ink = { count: 0, minX: Infinity, maxX: -Infinity };
    const shadow = { count: 0, minX: Infinity, maxX: -Infinity };
    for (let y = 0; y < largeRaster.height; y += 1) {
      for (let x = 0; x < largeRaster.width; x += 1) {
        const i = (y * largeRaster.width + x) * 4;
        const [r, g, b] = largeRaster.data.subarray(i, i + 3);
        const region =
          r! < 100 && g! < 100 && b! < 100
            ? ink
            : r! > 180 && g! < 100 && b! < 100
              ? shadow
              : undefined;
        if (region) {
          region.count++;
          region.minX = Math.min(region.minX, x);
          region.maxX = Math.max(region.maxX, x);
        }
      }
    }
    expect(ink.count).toBeGreaterThan(10);
    expect(shadow.count / ink.count).toBeGreaterThan(0.95);
    expect(shadow.count / ink.count).toBeLessThan(1.05);
    // The 1280px slide viewBox scales the authored 96px offset to 72px.
    expect(shadow.minX - ink.minX).toBeCloseTo(72, 0);
    expect(shadow.maxX - ink.maxX).toBeCloseTo(72, 0);
  });
});

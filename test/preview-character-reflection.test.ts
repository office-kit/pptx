import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  inches,
  setShapeTextFormat,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import { renderSlideToRgba } from '../packages/preview/src/node.ts';
import {
  defaultMeasurer,
  layoutTextSvg,
  type PieceInput,
  type TextBodyInput,
} from '../packages/preview/src/text-layout.ts';

const piece = (text: string, reflection?: PieceInput['reflection']): PieceInput => ({
  text,
  family: 'Carlito',
  sizePx: 32,
  bold: false,
  italic: false,
  letterSpacingPx: 0,
  fillHex: '#000000',
  underline: 'none',
  strike: false,
  baseline: 0,
  href: null,
  isBreak: false,
  ...(reflection ? { reflection } : {}),
});

const body = (pieces: PieceInput[]): TextBodyInput => ({
  boxXpx: 0,
  boxYpx: 0,
  boxWpx: 160,
  boxHpx: 160,
  anchor: 'top',
  wrap: true,
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

describe('character reflection in the SVG text layout', () => {
  it('does not merge an explicit default reflection into an adjacent plain run', () => {
    const svg = layoutTextSvg(
      { ...body([piece('plain'), piece('effect', {})]), wrap: false, reflectionsOnly: true },
      defaultMeasurer,
    );
    expect(svg).toContain('effect</tspan>');
    expect(svg).not.toContain('plain');
  });
  it('does not duplicate paragraph bullets in the reflection-only overlay', () => {
    const input = body([piece('Reflected', { scaleY: -1 })]);
    const svg = layoutTextSvg(
      {
        ...input,
        reflectionsOnly: true,
        paragraphs: input.paragraphs.map((paragraph) => ({
          ...paragraph,
          bullet: { text: '•', family: 'Carlito', sizePx: 32, fillHex: '#000000' },
        })),
      },
      defaultMeasurer,
    );
    expect(svg).toContain('data-pptx-reflection');
    expect(svg).not.toContain('•');
  });

  it('reflects each formatted run after wrapping, without reflecting an adjacent run', () => {
    const reflection = {
      blurEmu: 6350,
      offsetEmu: 0,
      startOpacity: 0.53,
      opacity: 0.003,
      endPosition: 0.355,
      scaleY: -0.9,
    };
    const svg = layoutTextSvg(
      body([piece('Reflected text that wraps', reflection), piece(' plain')]),
      defaultMeasurer,
    );

    expect((svg.match(/data-pptx-reflection="text"/g) ?? []).length).toBeGreaterThan(1);
    expect(svg).toContain('stop-opacity="0.530"');
    expect(svg).toContain('offset="0.355"');
    expect(svg).toContain('stdDeviation="0.33"');
    expect(svg).toContain('pointer-events="none"');
    expect(svg).toContain('aria-hidden="true"');
    expect(svg).toContain('plain</tspan>');
  });

  it('keeps reflection IDs unique across independently laid out text bodies', () => {
    const reflection = { scaleY: -1, startOpacity: 1, opacity: 0 };
    const first = layoutTextSvg(body([piece('one', reflection)]), defaultMeasurer);
    const second = layoutTextSvg(body([piece('two', reflection)]), defaultMeasurer);
    const firstId = first.match(/id="(text-reflection-[^"]+)-mask"/)?.[1];
    const secondId = second.match(/id="(text-reflection-[^"]+)-mask"/)?.[1];
    expect(firstId).toBeDefined();
    expect(secondId).toBeDefined();
    expect(secondId).not.toBe(firstId);
  });

  it('keeps the near fade at the visual contact edge after the vertical flip', () => {
    const common = { startOpacity: 0.8, opacity: 0.1, endPosition: 0.4 };
    const half = layoutTextSvg(body([piece('text', { ...common, scaleY: -0.5 })]), defaultMeasurer);
    const full = layoutTextSvg(body([piece('text', { ...common, scaleY: -1 })]), defaultMeasurer);
    expect(half).toContain('x1="0" y1="1" x2="0" y2="0"');
    expect(full).toContain('x1="0" y1="1" x2="0" y2="0"');
    expect(half).not.toMatch(/scale\(1 -1\)/);
    expect(full).toMatch(/scale\(1 -1\)/);
  });

  it('honors schema defaults and positive scale for a non-flipping reflection', () => {
    const svg = layoutTextSvg(
      body([piece('offset', { scaleY: 1, offsetEmu: 9525, opacity: 0.4 })]),
      defaultMeasurer,
    );
    expect(svg).toContain('data-pptx-reflection="text"');
    // dir defaults to zero degrees: a one-pixel offset points along +X.
    const offset = svg.match(/translate\(([-\d.]+) [-\d.]+\).*translate\(([-\d.]+)/);
    expect(offset).toBeDefined();
    expect(Number(offset?.[1]) + Number(offset?.[2])).toBeCloseTo(1, 2);
  });

  it('uses the center anchor for ctr alignment instead of the right edge', () => {
    const svg = layoutTextSvg(
      body([piece('center', { scaleY: -1, alignment: 'ctr' })]),
      defaultMeasurer,
    );
    const transform = svg.match(/<g transform="([^"]+)" mask="url\(#text-reflection/)?.[1];
    expect(transform).toBeDefined();
    // The anchor is the run midpoint, so the first translate is not the run's
    // right edge (which would be roughly 160px in this test body).
    expect(transform).not.toMatch(/translate\(1(?:6|7)\d/);
  });

  it('keeps foreignObject text editable while adding a sibling reflection layer', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'Editable reflection',
    });
    setShapeTextFormat(
      box,
      { reflection: { scaleY: -0.9, startOpacity: 0.53, opacity: 0.003, endPosition: 0.355 } },
      { range: { start: 0, end: 8 } },
    );
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'foreignObject' });
    expect(svg).toContain('<foreignObject');
    expect(svg).toContain('data-pptx-reflection="text"');
    expect(svg).toContain('pointer-events="none"');
  });

  it('survives the resvg raster path as visible reflected pixels', () => {
    const makeDeck = (withReflection: boolean) => {
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      const box = addSlideTextBox(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(1.5),
        text: 'Raster reflection',
      });
      if (withReflection) {
        setShapeTextFormat(
          box,
          { reflection: { scaleY: -1, startOpacity: 0.8, opacity: 0.2, offsetEmu: 9525 } },
          { range: { start: 0, end: 6 } },
        );
      }
      return { pres, slide };
    };
    const plainDeck = makeDeck(false);
    const reflectedDeck = makeDeck(true);
    const plain = renderSlideToRgba(plainDeck.pres, plainDeck.slide, { width: 640 }).image.data;
    const reflected = renderSlideToRgba(reflectedDeck.pres, reflectedDeck.slide, { width: 640 })
      .image.data;
    let changed = 0;
    for (let i = 0; i < plain.length; i += 1) {
      if (plain[i] !== reflected[i]) changed += 1;
    }
    expect(changed).toBeGreaterThan(100);
  });
});

// Gradient outlines, sketched outlines and the text fills that need the SVG
// glyph layer (no fill, picture, gradient text outline).

import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideLine,
  addSlideShape,
  addSlideTextBox,
  createPresentation,
  inches,
  setShapeStroke,
  setShapeStrokeSketch,
  setShapeTextFormat,
  type LineFill,
  type LineSketch,
  type TextFormat,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

const GRADIENT: LineFill = {
  kind: 'gradient',
  path: 'linear',
  angleDeg: 0,
  scaled: true,
  stops: [
    { offset: 0, color: '#FF0000' },
    { offset: 1, color: '#0000FF' },
  ],
};

const slideWithRect = () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shape = addSlideShape(slide, {
    preset: 'rect',
    x: inches(1),
    y: inches(1),
    w: inches(3),
    h: inches(2),
  });
  return { pres, slide, shape };
};

const textDeck = (format: TextFormat) => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shape = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(6),
    h: inches(2),
    text: 'Styled',
  });
  setShapeTextFormat(shape, { size: 40, ...format });
  return { pres, slide };
};

const gradientId = (svg: string): string => {
  const id =
    /<linearGradient id="([^"]+)"[^>]*>(?:(?!<\/linearGradient>).)*stop-color="#FF0000"/.exec(
      svg,
    )?.[1];
  if (!id) throw new Error('no red-to-blue <linearGradient>');
  return id;
};

describe('renderSlideToSvg: gradient lines', () => {
  it('strokes a shape outline with the gradient paint', () => {
    const { pres, slide, shape } = slideWithRect();
    setShapeStroke(shape, { fill: GRADIENT, widthEmu: 38_100 });
    const svg = renderSlideToSvg(pres, slide);
    const id = gradientId(svg);
    expect(svg).toMatch(new RegExp(`<rect [^>]*stroke="url\\(#${id}\\)" stroke-width="4\\.00"`));
    expect(svg).toContain('stop-color="#0000FF"');
  });

  it('spans a straight connector with a one-line-wide user-space box', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const line = addSlideLine(slide, {
      from: { x: inches(1), y: inches(2) },
      to: { x: inches(5), y: inches(2) },
    });
    setShapeStroke(line, { fill: GRADIENT, widthEmu: 38_100 });
    const svg = renderSlideToSvg(pres, slide);
    const id = gradientId(svg);
    const gradient = new RegExp(`<linearGradient id="${id}"[^>]*>`).exec(svg)![0];
    // An objectBoundingBox paint cannot span a zero-height line.
    expect(gradient).toContain('gradientUnits="userSpaceOnUse"');
    expect(gradient).toContain('matrix(384.0000 0.0000 0.0000 4.0000 96.0000 190.0000)');
    expect(svg).toMatch(new RegExp(`<line [^>]*stroke="url\\(#${id}\\)"`));
  });
});

describe('renderSlideToSvg: sketched outlines', () => {
  const pathOf = (sketch: LineSketch | null): string => {
    const { pres, slide, shape } = slideWithRect();
    setShapeStroke(shape, { color: '#000000', widthEmu: 12_700 });
    setShapeStrokeSketch(shape, sketch);
    return renderSlideToSvg(pres, slide);
  };

  it('keeps an unsketched rectangle crisp', () => {
    expect(pathOf(null)).toMatch(/<rect x="96\.00" y="96\.00" width="288\.00" height="192\.00"/);
  });

  it.each(['curved', 'freehand', 'scribble'] as const)(
    'draws %s as a seeded hand-drawn path',
    (sketch) => {
      const svg = pathOf(sketch);
      expect(svg).not.toMatch(/<rect x="96\.00" y="96\.00" width="288\.00"/);
      const d = /<path d="(M[^"]+)"[^>]*stroke="#000000"/.exec(svg)![1]!;
      expect(d).toMatch(/Q/);
      // Same shape, same wobble.
      expect(pathOf(sketch)).toBe(svg);
    },
  );

  it('wanders further for scribble than for curved', () => {
    const spread = (sketch: LineSketch): number => {
      const d = /<path d="(M[^"]+)"/.exec(pathOf(sketch))![1]!;
      // Points along the top edge, away from the corners.
      const ys = [...d.matchAll(/([\d.]+),([\d.]+)/g)]
        .filter((m) => Number(m[1]) > 120 && Number(m[1]) < 360 && Number(m[2]) < 150)
        .map((m) => Number(m[2]));
      return Math.max(...ys.map((y) => Math.abs(y - 96)));
    };
    expect(spread('scribble')).toBeGreaterThan(spread('curved'));
  });
});

describe('renderSlideToSvg: text fills and outlines', () => {
  it('paints no glyph fill for a noFill run', () => {
    const { pres, slide } = textDeck({ textFill: { kind: 'none' }, outline: { color: '#00FF00' } });
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    expect(svg).toMatch(/<tspan [^>]*fill="none"[^>]*stroke="#00FF00"[^>]*>Styled<\/tspan>/);
  });

  it('stretches a picture fill over the text block', () => {
    const { pres, slide } = textDeck({ textFill: { kind: 'image', bytes: PNG } });
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    const pattern =
      /<pattern id="([^"]+)" patternUnits="userSpaceOnUse"[^>]*><image href="data:image\/png;base64,[^"]+"[^>]*preserveAspectRatio="none"\/><\/pattern>/.exec(
        svg,
      );
    expect(pattern).not.toBeNull();
    expect(svg).toMatch(
      new RegExp(`<tspan [^>]*fill="url\\(#${pattern![1]}\\)"[^>]*>Styled</tspan>`),
    );
  });

  it('strokes a gradient text outline in both text layouts', () => {
    const { pres, slide } = textDeck({
      color: '#FFFFFF',
      outline: { fill: GRADIENT, widthEmu: 19_050 },
    });
    for (const textLayout of ['svg', 'foreignObject'] as const) {
      const svg = renderSlideToSvg(pres, slide, { textLayout });
      const id = gradientId(svg);
      expect(svg).toMatch(
        new RegExp(`<tspan [^>]*fill="#FFFFFF"[^>]*stroke="url\\(#${id}\\)"[^>]*>Styled</tspan>`),
      );
    }
    // The editable HTML glyphs stay transparent under the painted layer.
    expect(renderSlideToSvg(pres, slide)).toMatch(
      /<span style="[^"]*color:transparent[^"]*">Styled<\/span>/,
    );
  });
});

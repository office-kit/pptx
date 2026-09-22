// A run's own outline, shadow and glow, as the preview paints them.

import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  inches,
  setShapeTextFormat,
} from '@office-kit/pptx';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

const deck = (format: Parameters<typeof setShapeTextFormat>[1]) => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shape = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(1),
    text: 'WordArt 文字',
  });
  setShapeTextFormat(shape, { size: 40, ...format });
  return { pres, slide };
};

describe('renderSlideToSvg: character-level effects', () => {
  it('strokes the glyphs behind their fill', () => {
    const { pres, slide } = deck({ outline: { color: '#FF0000', widthEmu: 19050 } });
    const svg = renderSlideToSvg(pres, slide);

    // 19050 EMU = 2px at 96 DPI.
    expect(svg).toContain('-webkit-text-stroke:2.00px #FF0000');
    expect(svg).toContain('paint-order:stroke fill');
  });

  it('places the shadow at the stated angle and distance', () => {
    const { pres, slide } = deck({
      shadow: { color: '#000000', blurEmu: 0, offsetEmu: 95250, angleDeg: 0 },
    });
    const svg = renderSlideToSvg(pres, slide);

    // 95250 EMU = 10px, straight to the right (0° is 3 o'clock).
    expect(svg).toContain('text-shadow:10.00px 0.00px 0.00px #000000');
  });

  it('folds an effect’s alpha into the CSS color, which has no opacity of its own', () => {
    const { pres, slide } = deck({
      shadow: { color: '#112233', blurEmu: 0, offsetEmu: 0, angleDeg: 0, opacity: 0.25 },
    });

    expect(renderSlideToSvg(pres, slide)).toContain('rgba(17,34,51,0.250)');
  });

  it('draws a glow as a centred halo, ahead of the shadow layer', () => {
    const { pres, slide } = deck({
      glow: { color: '#00FF00', radiusEmu: 47625 },
      shadow: { color: '#000000', blurEmu: 0, offsetEmu: 0, angleDeg: 0 },
    });
    const svg = renderSlideToSvg(pres, slide);

    expect(svg).toContain('text-shadow:0 0 5.00px #00FF00,0.00px 0.00px 0.00px #000000');
  });

  it('paints nothing extra for a run that states no effects', () => {
    const { pres, slide } = deck({});
    const svg = renderSlideToSvg(pres, slide);

    expect(svg).not.toContain('text-shadow');
    expect(svg).not.toContain('text-stroke');
  });

  it('strokes the glyphs in the SVG text path too, where CSS does not reach', () => {
    const { pres, slide } = deck({ outline: { color: '#FF0000', widthEmu: 19050 } });
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });

    expect(svg).toContain('stroke="#FF0000"');
    expect(svg).toContain('paint-order="stroke fill"');
  });
});

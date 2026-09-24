import { describe, it, expect } from 'vitest';
import {
  createPresentation,
  addBlankSlide,
  setSlideBackground,
  getSlideBackground,
  savePresentation,
  loadPresentation,
  getSlides,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

describe('solid background opacity', () => {
  it('retains opacity through save/reload and paints it over the white slide base', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    setSlideBackground(slide, '#123456', 0.6);
    expect(getSlideBackground(slide)).toEqual({ kind: 'solid', color: '#123456', opacity: 0.6 });
    const loaded = await loadPresentation(await savePresentation(pres));
    expect(getSlideBackground(getSlides(loaded)[0]!)).toEqual(getSlideBackground(slide));
    expect(renderSlideToSvg(pres, slide, { textLayout: 'svg' })).toContain(
      'fill="#123456" fill-opacity="0.6"',
    );
    setSlideBackground(slide, 'accent1', 0);
    expect(getSlideBackground(slide)).toEqual({
      kind: 'solid',
      color: 'scheme:accent1',
      opacity: 0,
    });
    setSlideBackground(slide, '#FFFFFF');
    expect(getSlideBackground(slide)).toEqual({ kind: 'solid', color: '#FFFFFF' });
  });

  it('rejects invalid opacity without replacing the existing fill', () => {
    const slide = addBlankSlide(createPresentation());
    setSlideBackground(slide, '#123456', 0.6);
    for (const opacity of [-0.1, 1.1, NaN, Infinity]) {
      expect(() => setSlideBackground(slide, '#ABCDEF', opacity)).toThrow(RangeError);
      expect(getSlideBackground(slide)).toEqual({ kind: 'solid', color: '#123456', opacity: 0.6 });
    }
  });
});

import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import {
  getSlideBackground,
  getSlideBackgroundGradientFill,
  getSlideXmlString,
  getSlides,
  loadPresentation,
  savePresentation,
  setSlideBackground,
  setSlideBackgroundGradientFill,
  clearSlideBackground,
} from '../src/api/index.ts';

async function fixture() {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url)),
  );
  return { pres, slide: getSlides(pres)[0]! };
}

const stops = [
  { offset: 0, color: '#FF0000' as const, opacity: 0.4, brightness: 0.3 },
  { offset: 1, color: 'scheme:accent2' as const },
];

describe('background gradient editing', () => {
  it('replaces solid fill with an editable theme gradient and round-trips direction and stop settings', async () => {
    const { pres, slide } = await fixture();
    setSlideBackground(slide, '#123456');
    setSlideBackgroundGradientFill(slide, {
      stops,
      path: 'rect',
      focus: { left: 1, top: 1, right: 0, bottom: 0 },
      tileRect: { left: 0, top: 0, right: 0, bottom: 0 },
    });
    expect(getSlideBackground(slide)).toEqual({ kind: 'gradient' });
    const gradient = getSlideBackgroundGradientFill(slide);
    expect(gradient).toMatchObject({
      stops,
      path: 'rect',
      focus: { left: 1, top: 1, right: 0, bottom: 0 },
    });
    expect(getSlideXmlString(slide)).not.toContain('<a:solidFill><a:srgbClr val="123456"');
    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getSlideBackgroundGradientFill(getSlides(reloaded)[0]!)).toEqual(gradient);
    clearSlideBackground(slide);
    expect(getSlideBackground(slide)).toEqual({ kind: 'inherit' });
  });

  it('writes linear angle and scaling without changing other slides', async () => {
    const { pres, slide } = await fixture();
    const other = getSlides(pres)[1]!;
    const original = getSlideXmlString(other);
    setSlideBackgroundGradientFill(slide, { stops, angleDeg: 45, scaled: true });
    expect(getSlideBackgroundGradientFill(slide)).toMatchObject({ angleDeg: 45, scaled: true });
    expect(getSlideXmlString(other)).toBe(original);
  });

  it('rejects invalid gradients without modifying the previous background', async () => {
    const { pres, slide } = await fixture();
    setSlideBackground(slide, '#123456');
    const original = getSlideXmlString(slide);
    expect(() =>
      setSlideBackgroundGradientFill(slide, {
        stops: [{ offset: 2, color: '#FFFFFF' }, stops[1]!],
      }),
    ).toThrow(RangeError);
    expect(getSlideXmlString(slide)).toBe(original);
    expect(getSlideBackground(slide)).toEqual({ kind: 'solid', color: '#123456' });
    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getSlideBackground(getSlides(reloaded)[0]!)).toEqual({
      kind: 'solid',
      color: '#123456',
    });
  });
});

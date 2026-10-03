import { describe, expect, it } from 'vitest';
import { buildFontkitMeasurer } from '../packages/preview/src/measure.ts';

describe('fontkit ink metrics', () => {
  const measure = buildFontkitMeasurer();
  const spec = {
    family: 'Carlito',
    sizePx: 40,
    bold: false,
    italic: false,
    letterSpacingPx: 0,
  } as const;

  it('reports glyph ink bounds separately from line metrics', () => {
    const cap = measure('A', spec);
    const descender = measure('g', spec);
    expect(cap.ascentPx).toBeGreaterThan(0);
    expect(cap.descentPx).toBeGreaterThan(0);
    expect(cap.inkAscentPx).toBeGreaterThan(0);
    expect(cap.inkDescentPx).toBe(0);
    expect(descender.inkDescentPx).toBeGreaterThan(cap.inkDescentPx!);
    expect(descender.ascentPx).toBe(cap.ascentPx);
    expect(descender.descentPx).toBe(cap.descentPx);
  });

  it('keeps whitespace ink-free while retaining advance and line metrics', () => {
    const space = measure(' ', spec);
    expect(space.widthPx).toBeGreaterThan(0);
    expect(space.ascentPx).toBeGreaterThan(0);
    expect(space.descentPx).toBeGreaterThan(0);
    expect(space.inkAscentPx).toBe(0);
    expect(space.inkDescentPx).toBe(0);
  });

  it('matches the positioned fontkit bbox for supported combining marks', () => {
    const text = 'A\u0300';
    const source = 'packages/preview/fonts/LiberationSerif-Regular.ttf';
    // Bundled Liberation Serif: the grave accent's top is 1825 font units,
    // positioned 10 units above its outline, on a 2048-unit em. Ignoring the
    // shaped offset gives a different ascent even though advance is unchanged.
    const expectedTop = ((1825 + 10) * spec.sizePx) / 2048;
    const expectedBottom = 0;
    const measureShaped = buildFontkitMeasurer({ fonts: [{ family: 'Ink metrics', source }] });
    const actual = measureShaped(text, { ...spec, family: 'Ink metrics' });
    expect(actual.approximate).toBeUndefined();
    expect(actual.inkAscentPx).toBeCloseTo(expectedTop, 6);
    expect(actual.inkDescentPx).toBeCloseTo(expectedBottom, 6);
    expect(actual.widthPx).toBeCloseTo((1479 * spec.sizePx) / 2048, 6);
  });

  it('keeps ink bounds signed for punctuation above and below the baseline', () => {
    const above = measure('\u00B7', spec);
    const below = measure(',', spec);
    expect(above.inkAscentPx).toBeGreaterThan(0);
    expect(above.inkDescentPx).toBeLessThan(0);
    expect(measure(' \u00B7 ', spec).inkDescentPx).toBe(above.inkDescentPx);
    expect(below.inkDescentPx).toBeGreaterThan(0);
  });

  it('does not report partial ink bounds when a glyph is estimated', () => {
    const missing = measure('A\u{10FFFF}', spec);
    expect(missing.approximate).toBe(true);
    expect(missing.inkAscentPx).toBeUndefined();
    expect(missing.inkDescentPx).toBeUndefined();
  });
});

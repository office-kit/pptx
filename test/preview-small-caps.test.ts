import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  inches,
  loadPresentation,
  setShapeTextFormat,
} from '@office-kit/pptx';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

describe('preview SVG small caps', () => {
  it('uppercases lowercase source letters while keeping their glyphs smaller', async () => {
    const pres = await loadPresentation(await readFile(fixturePath));
    const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
    const shape = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(5),
      h: inches(1),
      text: 'aBCd',
    });
    setShapeTextFormat(shape, { size: 20, cap: 'small' });

    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    const tspans = [...svg.matchAll(/<tspan[^>]*font-size="([\d.]+)"[^>]*>([^<]*)<\/tspan>/g)].map(
      (match) => ({ size: Number(match[1]), text: match[2] }),
    );
    const text = tspans.map((span) => span.text).join('');
    const lower = tspans.filter((span) => span.text === 'A' || span.text === 'D');
    const authored = tspans.find((span) => span.text === 'BC');

    expect(text).toContain('ABCD');
    expect(lower).toHaveLength(2);
    expect(authored).toBeDefined();
    expect(lower[0]!.size).toBeCloseTo(lower[1]!.size, 5);
    expect(lower[0]!.size).toBeLessThan(authored!.size);
    // SVG attributes are rounded for compact output, so allow the formatter's
    // sub-per-mille rounding around the renderer's documented ratio.
    expect(lower[0]!.size / authored!.size).toBeCloseTo(0.8, 3);
  });

  it('keeps a combining accent with its lowercase small-cap grapheme', async () => {
    const pres = await loadPresentation(await readFile(fixturePath));
    const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
    const shape = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(5),
      h: inches(1),
      text: 'a\u0301BC',
    });
    setShapeTextFormat(shape, { size: 20, cap: 'small' });

    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    const tspans = [...svg.matchAll(/<tspan[^>]*font-size="([\d.]+)"[^>]*>([^<]*)<\/tspan>/g)].map(
      (match) => ({ size: Number(match[1]), text: match[2] }),
    );
    const accented = tspans.find((span) => span.text === 'A\u0301');
    const authored = tspans.find((span) => span.text === 'BC');

    expect(accented).toBeDefined();
    expect(authored).toBeDefined();
    expect(accented!.size).toBeLessThan(authored!.size);
  });
});

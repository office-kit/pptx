import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  inches,
  loadPresentation,
  setShapeParagraphs,
} from '../src/api/index.ts';
import { renderSlideToRgba } from '../packages/preview/src/node.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

const blankSlide = async () => {
  const pres = await loadPresentation(
    await (await import('node:fs/promises')).readFile(
      new URL('./fixtures/minimal/blank.pptx', import.meta.url),
    ),
  );
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('Blank layout not found');
  return { pres, slide: addSlide(pres, { layout }) };
};

describe('underline color preview', () => {
  it('keeps an explicit SVG underline color separate from a solid strike', async () => {
    const { pres, slide } = await blankSlide();
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(1),
      text: 'colored underline',
    });
    setShapeParagraphs(box, [
      {
        runs: [
          {
            text: 'colored ',
            format: {
              color: '#154687',
              underline: 'sng',
              underlineColor: '#FF0000',
              strike: true,
              size: 32,
            },
          },
          {
            text: 'underline',
            format: { color: '#008000', underline: 'sng', underlineColor: '#0000FF', size: 32 },
          },
        ],
      },
    ]);
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    expect(svg).toContain('text-decoration="line-through"');
    expect(svg).not.toContain('text-decoration="underline"');
    expect(svg).not.toContain('text-decoration-color="#FF0000"');
    expect(svg).not.toContain('text-decoration-color="#0000FF"');
    expect(svg).toContain('stroke="#FF0000"');
    expect(svg).toContain('stroke="#0000FF"');
    expect(svg).not.toContain('text-decoration="underline line-through"');
  });

  it('rasterizes the underline red while glyphs and strike remain blue', async () => {
    const { pres, slide } = await blankSlide();
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(1),
      text: 'colored underline',
    });
    setShapeParagraphs(box, [
      {
        runs: [
          {
            text: 'colored underline',
            format: {
              color: '#154687',
              underline: 'sng',
              underlineColor: '#FF0000',
              strike: true,
              size: 32,
            },
          },
        ],
      },
    ]);
    const { image } = renderSlideToRgba(pres, slide, { width: 960 });
    let red = 0;
    let blue = 0;
    for (let i = 0; i < image.data.length; i += 4) {
      const r = image.data[i]!;
      const g = image.data[i + 1]!;
      const b = image.data[i + 2]!;
      if (r > 180 && g < 100 && b < 100) red++;
      if (b > 60 && r < 100 && g < 180) blue++;
    }
    expect(red).toBeGreaterThan(2);
    expect(blue).toBeGreaterThan(2);
  });
});

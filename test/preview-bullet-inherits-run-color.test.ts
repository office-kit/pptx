// A bullet with no `<a:buClr>` takes the colour of its paragraph's first run,
// not the body default — otherwise a dark deck's numbered agenda draws its
// markers in the default black and they vanish against the background.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  inches,
  loadPresentation,
  setShapeText,
  setShapeTextFormat,
} from '@office-kit/pptx';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

const RUN_COLOR = '#ADADAD';

const numberedDeck = async () => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
  const box = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(8),
    h: inches(3),
    text: 'one\ntwo',
  });
  setShapeText(box, 'one\ntwo', { bullets: 'number' });
  setShapeTextFormat(box, { color: RUN_COLOR });
  return { pres, slide };
};

describe('renderSlideToSvg: bullet colour', () => {
  it('paints an unstyled number in the first run’s colour', async () => {
    const { pres, slide } = await numberedDeck();
    const svg = renderSlideToSvg(pres, slide);

    const markers = [...svg.matchAll(/<span style="([^"]*)">1\.<\/span>/g)].map((m) => m[1]!);
    expect(markers).toHaveLength(1);
    expect(markers[0]).toContain(`color:${RUN_COLOR}`);
  });

  it('does the same on the pure-SVG path', async () => {
    const { pres, slide } = await numberedDeck();
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });

    const marker = svg.match(/<text[^>]*fill="([^"]*)"[^>]*>1\.<\/text>/);
    expect(marker?.[1]).toBe(RUN_COLOR);
  });
});

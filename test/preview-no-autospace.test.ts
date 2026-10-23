import { expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTable,
  addSlideTextBox,
  createPresentation,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

// Browser-laid text must not gain CJK/Latin spacing the deck never had.
it('turns off CSS autospace in text boxes and table cells', async () => {
  const source = createPresentation();
  const slide = addBlankSlide(source);
  addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(1),
    text: '9 月に 15,640 件',
  });
  addSlideTable(slide, {
    x: inches(1),
    y: inches(3),
    w: inches(4),
    h: inches(1),
    rows: [['AI の回答']],
  });
  const pres = await loadPresentation(await savePresentation(source));
  const svg = renderSlideToSvg(pres, getSlides(pres)[0]!);

  const containers = svg.match(/<foreignObject[^>]*><div [^>]*style="[^"]*"/g) ?? [];
  expect(containers).toHaveLength(2);
  for (const div of containers) expect(div).toContain('text-autospace:no-autospace');
});

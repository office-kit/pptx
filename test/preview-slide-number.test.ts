import { expect, it } from 'vitest';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import {
  createPresentation,
  moveSlide,
  addBlankSlide,
  addSlideTextBox,
  inches,
  savePresentation,
  loadPresentation,
  getSlides,
  setSlideSize,
  getSlideShapes,
  setShapeTextRangeClickAction,
  SLIDE_SIZE_4_3,
} from '@office-kit/pptx';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

it('renders live slide-number fields while preserving unrelated cached fields', async () => {
  const pres = createPresentation();
  for (let i = 0; i < 2; i++) {
    const slide = addBlankSlide(pres);
    addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'CACHED_NUMBER',
    });
  }
  const parts = unzipSync(await savePresentation(pres));
  for (const name of Object.keys(parts).filter((name) =>
    /^ppt\/slides\/slide\d+\.xml$/.test(name),
  )) {
    const xml = strFromU8(parts[name]!);
    parts[name] = strToU8(
      xml.replace(
        /<a:r>([\s\S]*?)<\/a:r>/,
        '<a:fld id="{DCC01256-5B26-4D21-ACCE-781D34E01234}" type="slidenum">$1</a:fld>',
      ),
    );
  }
  const loaded = await loadPresentation(zipSync(parts));
  setSlideSize(loaded, SLIDE_SIZE_4_3, { firstSlideNumber: 10 });
  const first = getSlides(loaded)[0]!;
  for (const textLayout of ['svg', 'foreignObject'] as const) {
    const rendered = getSlides(loaded).map((slide) =>
      renderSlideToSvg(loaded, slide, { textLayout }),
    );
    expect(rendered[0]).not.toContain('CACHED_NUMBER');
    expect(rendered[0]).toMatch(/>10</);
    expect(rendered[1]).toMatch(/>11</);
  }
  setShapeTextRangeClickAction(
    getSlideShapes(first)[0]!,
    0,
    'CACHED_NUMBER'.length,
    { kind: 'nextSlide' },
    'Number link',
  );
  for (const textLayout of ['svg', 'foreignObject'] as const) {
    const svg = renderSlideToSvg(loaded, first, { textLayout });
    expect(svg).toContain('href="#slide-nextSlide"');
    expect(svg).toContain('Number link');
    expect(svg).toMatch(/>10</);
  }
  moveSlide(loaded, first, 1);
  expect(renderSlideToSvg(loaded, first)).toMatch(/>11</);
  for (const name of Object.keys(parts).filter((name) =>
    /^ppt\/slides\/slide\d+\.xml$/.test(name),
  )) {
    parts[name] = strToU8(strFromU8(parts[name]!).replace('type="slidenum"', 'type="datetime"'));
  }
  const dates = await loadPresentation(zipSync(parts));
  expect(renderSlideToSvg(dates, getSlides(dates)[0]!)).toContain('CACHED_NUMBER');
});

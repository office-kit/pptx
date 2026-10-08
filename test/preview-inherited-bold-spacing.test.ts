// WordArt presets remove `b` and `spc` from the runs, as the reference desktop app does, so
// the text falls back to the weight and spacing its list style gives it. The
// preview has to draw that inherited bold and tracking, which an explicit
// `b="0"` / `spc="0"` would have overridden.

import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
  savePresentation,
  setShapeTextFormat,
  type TextFormat,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

// 3pt of tracking in 1/100 pt, as the list style authors it.
const INHERITED_SPC = 300;
const PX_PER_PT = 4 / 3;

const deckWithStyledList = async () => {
  const pres = createPresentation();
  addSlideTextBox(addBlankSlide(pres), {
    x: inches(1),
    y: inches(1),
    w: inches(6),
    h: inches(1),
    text: 'Outline title',
  });
  const zip = readZip(await savePresentation(pres));
  const entry = zip.entries.find((part) => part.name === 'ppt/slides/slide1.xml')!;
  const xml = new TextDecoder().decode(entry.data);
  const styled = xml.replace(
    /<a:lstStyle\/>/,
    `<a:lstStyle><a:lvl1pPr><a:defRPr b="1" spc="${INHERITED_SPC}"/></a:lvl1pPr></a:lstStyle>`,
  );
  expect(styled).not.toBe(xml);
  const data = new TextEncoder().encode(styled);
  return loadPresentation(
    writeZip(zip.entries.map((part) => (part === entry ? { ...part, data } : part))),
  );
};

const render = async (format: TextFormat) => {
  const pres = await deckWithStyledList();
  const slide = getSlides(pres)[0]!;
  setShapeTextFormat(getSlideShapes(slide).at(-1)!, format);
  return renderSlideToSvg(pres, slide, { textLayout: 'foreignObject' });
};

describe('preview of bold and spacing removed from a run', () => {
  const tracking = `letter-spacing:${((INHERITED_SPC / 100) * PX_PER_PT).toFixed(3)}px`;

  it('draws the inherited bold and spacing once the run no longer sets them', async () => {
    const html = await render({ bold: null, spc: null });
    expect(html).toContain('font-weight:700');
    expect(html).toContain(tracking);
  });

  it('an explicit b="0" / spc="0" overrides them instead', async () => {
    const html = await render({ bold: false, spc: 0 });
    expect(html).not.toContain('font-weight:700');
    expect(html).not.toContain(tracking);
  });
});

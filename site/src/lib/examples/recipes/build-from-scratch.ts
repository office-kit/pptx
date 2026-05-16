// Build a deck on top of a blank template: add a title slide and a content
// slide, then save.

import { readFile, writeFile } from 'node:fs/promises';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  findSlidePlaceholder,
  inches,
  loadPresentation,
  savePresentation,
  setShapeText,
} from 'pptx-kit';

const pres = await loadPresentation(await readFile('blank.pptx'));

const titleLayout = findSlideLayout(pres, 'Title Slide');
if (titleLayout) {
  const cover = addSlide(pres, { layout: titleLayout });
  const t = findSlidePlaceholder(cover, 'ctrTitle') ?? findSlidePlaceholder(cover, 'title');
  if (t) setShapeText(t, 'pptx-kit demo');
  const sub = findSlidePlaceholder(cover, 'subTitle');
  if (sub) setShapeText(sub, 'an OOXML library for TypeScript');
}

const blank = findSlideLayout(pres, 'Blank') ?? titleLayout;
if (blank) {
  const body = addSlide(pres, { layout: blank });
  addSlideTextBox(body, {
    x: inches(1),
    y: inches(1),
    w: inches(8),
    h: inches(1),
    text: 'Free-form text box',
  });
}

await writeFile('out.pptx', await savePresentation(pres));

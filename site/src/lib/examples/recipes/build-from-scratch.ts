// Build a deck from scratch: the deck createPresentation() returns already
// has 'Title Slide', 'Title and Content' and 'Blank' layouts to add slides on.

import { writeFile } from 'node:fs/promises';
import {
  addSlide,
  addSlideTextBox,
  createPresentation,
  findSlideLayout,
  findSlidePlaceholder,
  inches,
  savePresentation,
  setShapeText,
} from '@office-kit/pptx';

const pres = createPresentation();

const titleLayout = findSlideLayout(pres, 'Title Slide');
if (titleLayout) {
  const cover = addSlide(pres, { layout: titleLayout });
  const t = findSlidePlaceholder(cover, 'ctrTitle') ?? findSlidePlaceholder(cover, 'title');
  if (t) setShapeText(t, '@office-kit/pptx demo');
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

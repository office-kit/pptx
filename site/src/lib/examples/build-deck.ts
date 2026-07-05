// Build a deck from a blank template: add slides on a layout, drop in a
// text box, an image, and a chart, then save.

import { readFile, writeFile } from 'node:fs/promises';
import {
  addSlide,
  addSlideChart,
  addSlideImage,
  addSlideTextBox,
  findSlideLayout,
  findSlidePlaceholder,
  inches,
  loadPresentation,
  savePresentation,
  setShapeText,
} from '@office-kit/pptx';

const pres = await loadPresentation(await readFile('blank.pptx'));
const titleLayout = findSlideLayout(pres, 'Title Slide');
if (!titleLayout) throw new Error('no Title Slide layout');

const cover = addSlide(pres, { layout: titleLayout });
const title = findSlidePlaceholder(cover, 'ctrTitle') ?? findSlidePlaceholder(cover, 'title');
if (title) setShapeText(title, 'Q3 review');

const blank = findSlideLayout(pres, 'Blank') ?? titleLayout;
const slide = addSlide(pres, { layout: blank });
addSlideTextBox(slide, {
  x: inches(0.7),
  y: inches(0.5),
  w: inches(9),
  h: inches(0.7),
  text: 'Numbers up and to the right',
});
addSlideChart(slide, {
  x: inches(0.7),
  y: inches(1.5),
  w: inches(8),
  h: inches(4.5),
  spec: {
    kind: 'column',
    categories: ['Q1', 'Q2', 'Q3', 'Q4'],
    series: [{ name: 'Revenue', values: [120, 180, 240, 300] }],
    title: 'FY26',
  },
});
addSlideImage(slide, await readFile('logo.png'), {
  x: inches(8),
  y: inches(0.4),
  w: inches(1.5),
  h: inches(1),
});

await writeFile('out.pptx', await savePresentation(pres));

// Build a deck from scratch: start from the built-in empty deck, add slides
// on its layouts, drop in a text box, an image, and a chart, then save.

import { readFile, writeFile } from 'node:fs/promises';
import {
  addBlankSlide,
  addSlideChart,
  addSlideImage,
  addSlideTextBox,
  addTitleSlide,
  createPresentation,
  inches,
  savePresentation,
} from '@office-kit/pptx';

const pres = createPresentation();
addTitleSlide(pres, 'Q3 review');

const slide = addBlankSlide(pres);
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

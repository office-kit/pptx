// Load a presentation, walk every slide, print titles and text length.

import { readFile } from 'node:fs/promises';
import {
  getSlideText,
  getSlideTextLength,
  getSlideTitle,
  getSlides,
  loadPresentation,
} from '@office-kit/pptx';

const pres = await loadPresentation(await readFile('input.pptx'));
for (const slide of getSlides(pres)) {
  console.log(`${getSlideTitle(slide) ?? '(no title)'} — ${getSlideTextLength(slide)} chars`);
  console.log(getSlideText(slide));
}

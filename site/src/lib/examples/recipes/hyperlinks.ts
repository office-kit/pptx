// External URL hyperlink + an in-deck click action (jump to slide 2).

import {
  addSlideTextBox,
  getSlides,
  inches,
  setShapeClickAction,
  setShapeHyperlink,
  setShapeTextFormat,
  type PresentationData,
  type SlideData,
} from 'pptx-kit';

declare const pres: PresentationData;
declare const slide: SlideData;

const link = addSlideTextBox(slide, {
  x: inches(1),
  y: inches(2),
  w: inches(6),
  h: inches(0.6),
  text: 'Open the docs',
});
setShapeHyperlink(link, 'https://github.com/baseballyama/pptx-kit');
setShapeTextFormat(link, { color: '#0563C1', underline: true });

const nav = addSlideTextBox(slide, {
  x: inches(1),
  y: inches(3),
  w: inches(6),
  h: inches(0.6),
  text: 'Jump to slide 2',
});
const slide2 = getSlides(pres)[1];
if (slide2) setShapeClickAction(nav, { kind: 'slide', slide: slide2 });

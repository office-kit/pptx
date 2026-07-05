// Drop in a few preset shapes from the 180+ DrawingML preset library.

import { addSlideShape, inches, type SlideData } from '@office-kit/pptx';

declare const slide: SlideData;

addSlideShape(slide, {
  preset: 'star5',
  x: inches(1),
  y: inches(2),
  w: inches(2),
  h: inches(2),
  text: '★',
});
addSlideShape(slide, {
  preset: 'rightArrow',
  x: inches(4),
  y: inches(2.4),
  w: inches(2.5),
  h: inches(1.2),
  text: 'next',
});
addSlideShape(slide, {
  preset: 'roundRect',
  x: inches(7),
  y: inches(2),
  w: inches(2.5),
  h: inches(2),
  text: 'callout',
});

// Slide transitions + shape entrance / exit animations.

import {
  addSlideShape,
  inches,
  setShapeAnimation,
  setShapeFill,
  setSlideTransition,
  type SlideData,
} from 'pptx-kit';

declare const slide: SlideData;

setSlideTransition(slide, { effect: 'fade', speed: 'med' });

const a = addSlideShape(slide, {
  preset: 'roundRect',
  x: inches(1),
  y: inches(2),
  w: inches(3),
  h: inches(1.5),
  text: 'fadeIn',
});
setShapeFill(a, '#2E75B6');
setShapeAnimation(a, { effect: 'fadeIn', durationMs: 700 });

const b = addSlideShape(slide, {
  preset: 'roundRect',
  x: inches(5),
  y: inches(2),
  w: inches(3),
  h: inches(1.5),
  text: 'fadeOut',
});
setShapeFill(b, '#C00000');
setShapeAnimation(b, { effect: 'fadeOut', durationMs: 700 });

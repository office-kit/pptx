// Solid / gradient / pattern fills + a drop shadow and an outer glow.

import {
  addSlideShape,
  inches,
  pt,
  setShapeFill,
  setShapeGlow,
  setShapeGradientFill,
  setShapePatternFill,
  setShapeShadow,
  type SlideData,
} from 'pptx-kit';

declare const slide: SlideData;

const solid = addSlideShape(slide, {
  preset: 'rect',
  x: inches(0.7),
  y: inches(1),
  w: inches(2.5),
  h: inches(2),
  text: 'Solid',
});
setShapeFill(solid, '#C00000');
setShapeShadow(solid, {
  blurEmu: pt(8),
  offsetEmu: pt(4),
  angleDeg: 45,
  color: '#000000',
  opacity: 0.5,
});

const grad = addSlideShape(slide, {
  preset: 'rect',
  x: inches(3.5),
  y: inches(1),
  w: inches(2.5),
  h: inches(2),
  text: 'Gradient',
});
setShapeGradientFill(grad, {
  stops: [
    { offset: 0, color: '#FFD966' },
    { offset: 1, color: '#C00000' },
  ],
  angleDeg: 45,
});

const pat = addSlideShape(slide, {
  preset: 'rect',
  x: inches(6.3),
  y: inches(1),
  w: inches(2.5),
  h: inches(2),
  text: 'Pattern',
});
setShapePatternFill(pat, { preset: 'pct50', foreground: '#1F4E79', background: '#FFFFFF' });
setShapeGlow(pat, { radiusEmu: pt(12), color: '#2E75B6' });

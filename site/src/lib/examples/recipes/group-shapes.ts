// Compose a rectangle + label into a single "KPI card" component, then
// move/resize it as one unit — the members scale with it.

import {
  addSlideShape,
  addSlideTextBox,
  groupShapes,
  inches,
  setShapeFill,
  setShapePosition,
  setShapeSize,
  type SlideData,
} from '@office-kit/pptx';

declare const slide: SlideData;

const card = addSlideShape(slide, {
  preset: 'roundRect',
  x: inches(1),
  y: inches(1),
  w: inches(2.5),
  h: inches(1.2),
});
setShapeFill(card, '#0B1F3A');

const label = addSlideTextBox(slide, {
  x: inches(1.2),
  y: inches(1.3),
  w: inches(2.1),
  h: inches(0.6),
  text: 'Revenue +12%',
});

const kpiCard = groupShapes([card, label], { name: 'KPI Card' });

// Move + resize the group; the rectangle and label scale together.
setShapePosition(kpiCard, inches(5), inches(2));
setShapeSize(kpiCard, inches(3.75), inches(1.8));

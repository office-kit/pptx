// Column chart with two series. addSlideChart writes the chart XML, the
// drawing rels, and an embedded xlsx so PowerPoint's "Edit data" works.

import { addSlideChart, inches, type SlideData } from '@office-kit/pptx';

declare const slide: SlideData;

addSlideChart(slide, {
  x: inches(0.5),
  y: inches(1.5),
  w: inches(9),
  h: inches(4.5),
  spec: {
    kind: 'column',
    categories: ['Q1', 'Q2', 'Q3', 'Q4'],
    series: [
      { name: 'Revenue', values: [120, 180, 240, 300] },
      { name: 'Cost', values: [80, 90, 130, 160] },
    ],
    title: 'FY26',
  },
});

// Add a 5×4 table with a header row and banded styling.

import { addSlideTable, inches, type SlideData } from '@office-kit/pptx';

declare const slide: SlideData;

addSlideTable(slide, {
  x: inches(0.7),
  y: inches(1.5),
  w: inches(9),
  h: inches(3),
  rows: [
    ['Quarter', 'Revenue', 'Cost', 'Margin'],
    ['Q1', '$1.2M', '$0.8M', '33%'],
    ['Q2', '$1.8M', '$0.9M', '50%'],
    ['Q3', '$2.4M', '$1.3M', '46%'],
    ['Q4', '$3.0M', '$1.6M', '47%'],
  ],
  firstRow: true,
  bandRow: true,
});

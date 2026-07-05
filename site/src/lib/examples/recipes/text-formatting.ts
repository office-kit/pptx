// Per-run formatting: bold, size, color, font on a single paragraph.
//
// Index a shape's text by (paragraphIndex, runIndex). Each \n in the
// `text:` argument starts a new paragraph; the first run of each is
// index 0.

import {
  addSlideTextBox,
  inches,
  setParagraphAlignment,
  setShapeRunFormat,
  type SlideData,
} from '@office-kit/pptx';

declare const slide: SlideData;

const box = addSlideTextBox(slide, {
  x: inches(0.7),
  y: inches(1.5),
  w: inches(9),
  h: inches(2.5),
  text: 'Default text\nBold red 24pt Calibri\nItalic underline 18pt',
});
setShapeRunFormat(box, 1, 0, {
  bold: true,
  size: 24,
  color: '#C00000',
  font: 'Calibri',
});
setShapeRunFormat(box, 2, 0, {
  italic: true,
  underline: true,
  size: 18,
  color: '#1F4E79',
});
setParagraphAlignment(box, 2, 'ctr');

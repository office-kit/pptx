// Insert an image. setShapeImage / addSlideImage detect the format
// (PNG / JPEG / GIF / SVG / BMP / TIFF) from magic bytes.

import { readFile } from 'node:fs/promises';
import { addSlideImage, inches, type SlideData } from 'pptx-kit';

declare const slide: SlideData;

const logo = await readFile('logo.png');
addSlideImage(slide, logo, {
  x: inches(8),
  y: inches(0.4),
  w: inches(1.5),
  h: inches(1),
});

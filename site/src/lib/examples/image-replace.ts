// Swap an image in a template, preserving its crop / transform / sizing.
// The format is auto-detected from the new bytes.

import { readFile, writeFile } from 'node:fs/promises';
import {
  getShapeKind,
  getShapeName,
  getSlideShapes,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeImage,
} from 'pptx-kit';

const pres = await loadPresentation(await readFile('template.pptx'));
const newLogo = await readFile('new-logo.png');

for (const slide of getSlides(pres)) {
  for (const shape of getSlideShapes(slide)) {
    if (getShapeKind(shape) === 'picture' && getShapeName(shape) === 'Logo') {
      setShapeImage(shape, newLogo);
    }
  }
}
await writeFile('out.pptx', await savePresentation(pres));

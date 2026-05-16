// Browser: fetch a .pptx and load it. loadPresentation accepts any of
// Uint8Array / ArrayBuffer / Blob — `await response.arrayBuffer()` is the
// most browser-portable form.

import { getSlides, getSlideTitle, loadPresentation } from 'pptx-kit';

const response = await fetch('/template.pptx');
const pres = await loadPresentation(await response.arrayBuffer());
for (const slide of getSlides(pres)) {
  console.log(getSlideTitle(slide));
}

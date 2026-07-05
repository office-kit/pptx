// Inspect the raw OPC parts in a package without dropping to the
// internal OpcPackage class.

import { type PresentationData, listPackageParts, readPackagePart } from '@office-kit/pptx';

declare const pres: PresentationData;

for (const part of listPackageParts(pres)) {
  console.log(`${part.name} (${part.contentType}, ${part.byteLength} bytes)`);
}

const themeBytes = readPackagePart(pres, '/ppt/theme/theme1.xml');
if (themeBytes) console.log(new TextDecoder().decode(themeBytes));

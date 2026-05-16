// One-shot read + save direct from / to disk via the pptx-kit/node
// helpers, no manual fs glue needed.

import { loadPresentationFile, savePresentationToFile } from 'pptx-kit/node';

const pres = await loadPresentationFile('input.pptx');
// ...mutate pres...
await savePresentationToFile(pres, 'output.pptx');

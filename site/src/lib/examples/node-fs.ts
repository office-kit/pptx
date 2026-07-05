// One-shot read + save direct from / to disk via the @office-kit/pptx/node
// helpers, no manual fs glue needed.

import { loadPresentationFile, savePresentationToFile } from '@office-kit/pptx/node';

const pres = await loadPresentationFile('input.pptx');
// ...mutate pres...
await savePresentationToFile(pres, 'output.pptx');

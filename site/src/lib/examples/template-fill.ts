// Open a template .pptx, replace `{{tokens}}` everywhere, save it back.
//
// This file is imported as ?raw into the docs site so the snippet shown
// to readers is exactly what svelte-check / tsc compiled — if an API
// rename breaks this import, the docs build fails before deploy.

import { readFile, writeFile } from 'node:fs/promises';
import {
  loadPresentation,
  replaceTokensInPresentation,
  savePresentation,
} from 'pptx-kit';

const pres = await loadPresentation(await readFile('template.pptx'));
replaceTokensInPresentation(pres, {
  name: 'Yamashita',
  event: 'Re:Invent',
  date: '2026-12-01',
});
await writeFile('out.pptx', await savePresentation(pres));

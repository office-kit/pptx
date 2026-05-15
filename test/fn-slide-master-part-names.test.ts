// getSlideMasterPartNames — slide-master URIs declared by presentation.xml.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getSlideMasterCount,
  getSlideMasterPartNames,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideMasterPartNames', () => {
  it('returns paths matching the master-count', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const names = getSlideMasterPartNames(pres);
    expect(names.length).toBe(getSlideMasterCount(pres));
    for (const n of names) {
      expect(n).toMatch(/^\/ppt\/slideMasters\/slideMaster\d+\.xml$/);
    }
  });

  it('returns at least one master on a real fixture', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    expect(getSlideMasterPartNames(pres).length).toBeGreaterThanOrEqual(1);
  });
});

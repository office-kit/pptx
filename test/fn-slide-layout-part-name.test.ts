// getSlideLayoutPartName / findSlideLayoutByPartName — mirror of
// getSlidePartName / findSlideByPartName, for layouts.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  findSlideLayoutByPartName,
  getSlideLayoutPartName,
  getSlideLayouts,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideLayoutPartName / findSlideLayoutByPartName', () => {
  it('every layout reports a /ppt/slideLayouts/... path', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    for (const layout of getSlideLayouts(pres)) {
      const path = getSlideLayoutPartName(layout);
      expect(path.startsWith('/ppt/slideLayouts/slideLayout')).toBe(true);
      expect(path.endsWith('.xml')).toBe(true);
    }
  });

  it('round-trips via the part name', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    for (const layout of getSlideLayouts(pres)) {
      const path = getSlideLayoutPartName(layout);
      const looked = findSlideLayoutByPartName(pres, path);
      expect(looked).not.toBeNull();
      // getSlideLayouts builds fresh SlideLayoutData on each call,
      // so identity won't hold; compare by part name instead.
      expect(getSlideLayoutPartName(looked!)).toBe(path);
    }
  });

  it('returns null on an unknown path', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    expect(findSlideLayoutByPartName(pres, '/ppt/slideLayouts/no-such.xml')).toBeNull();
  });
});

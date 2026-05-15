// findLayoutsWithPlaceholderType — every layout that can host a slot.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  findLayoutsWithPlaceholderType,
  getSlideLayoutPlaceholders,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findLayoutsWithPlaceholderType', () => {
  it('finds every layout with a title slot', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layouts = findLayoutsWithPlaceholderType(pres, 'title');
    expect(layouts.length).toBeGreaterThan(0);
    for (const layout of layouts) {
      const phs = getSlideLayoutPlaceholders(layout);
      expect(phs.some((p) => p.type === 'title')).toBe(true);
    }
  });

  it('finds every layout with a body slot', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layouts = findLayoutsWithPlaceholderType(pres, 'body');
    expect(layouts.length).toBeGreaterThan(0);
  });

  it('returns an empty list for a placeholder type no layout exposes', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    expect(findLayoutsWithPlaceholderType(pres, 'no-such-ph-type')).toEqual([]);
  });
});

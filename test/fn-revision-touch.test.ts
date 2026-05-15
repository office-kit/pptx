// incrementRevision + touchModified — sugar around setCoreProperties
// for "I'm about to save, please refresh the metadata."

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getCoreProperties,
  getPresentationModified,
  incrementRevision,
  loadPresentation,
  setCoreProperties,
  touchModified,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: incrementRevision', () => {
  it('bumps an existing revision', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    setCoreProperties(pres, { revision: '5' });
    expect(incrementRevision(pres)).toBe(6);
    expect(getCoreProperties(pres)!.revision).toBe('6');
  });

  it('treats missing/unparseable revision as 0', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    setCoreProperties(pres, { revision: null });
    expect(incrementRevision(pres)).toBe(1);

    setCoreProperties(pres, { revision: 'not a number' });
    expect(incrementRevision(pres)).toBe(1);
  });
});

describe('fn API: touchModified', () => {
  it('writes the current time by default', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const before = Date.now();
    touchModified(pres);
    const after = Date.now();
    const modified = getPresentationModified(pres)!;
    expect(modified.getTime()).toBeGreaterThanOrEqual(before - 1000);
    expect(modified.getTime()).toBeLessThanOrEqual(after + 1000);
  });

  it('accepts an explicit Date', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const t = new Date('2026-05-15T12:34:56Z');
    touchModified(pres, t);
    expect(getPresentationModified(pres)!.toISOString()).toBe(t.toISOString());
  });
});

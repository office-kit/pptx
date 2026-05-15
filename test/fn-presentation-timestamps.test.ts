// getPresentationCreated / getPresentationModified — convenience
// over getCoreProperties() that parses the W3CDTF strings into Date.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getPresentationCreated,
  getPresentationModified,
  loadPresentation,
  setCoreProperties,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: presentation timestamps', () => {
  it('parses the fixture timestamps as Date objects', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const created = getPresentationCreated(pres);
    const modified = getPresentationModified(pres);
    expect(created).toBeInstanceOf(Date);
    expect(modified).toBeInstanceOf(Date);
    expect(Number.isFinite(created!.getTime())).toBe(true);
    expect(Number.isFinite(modified!.getTime())).toBe(true);
  });

  it('round-trips through setCoreProperties', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    setCoreProperties(pres, {
      created: '2026-01-01T00:00:00Z',
      modified: '2026-05-15T12:34:56Z',
    });
    expect(getPresentationCreated(pres)!.toISOString()).toBe('2026-01-01T00:00:00.000Z');
    expect(getPresentationModified(pres)!.toISOString()).toBe('2026-05-15T12:34:56.000Z');
  });

  it('returns null when the value is unparseable', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    setCoreProperties(pres, { created: 'not a date', modified: null });
    expect(getPresentationCreated(pres)).toBeNull();
    expect(getPresentationModified(pres)).toBeNull();
  });
});

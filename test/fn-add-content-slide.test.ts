// addContentSlide — sugar combining title + body in one call.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addContentSlide,
  getSlideBody,
  getSlideTitle,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: addContentSlide', () => {
  it('writes both title and body', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const slide = addContentSlide(pres, {
      title: 'Roadmap',
      body: 'Q1: discovery\nQ2: build',
    });
    expect(getSlideTitle(slide)).toBe('Roadmap');
    expect(getSlideBody(slide)).toContain('Q1');
  });

  it('skips body / title when not provided', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const slide = addContentSlide(pres, { title: 'Only title' });
    expect(getSlideTitle(slide)).toBe('Only title');
  });
});

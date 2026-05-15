// setSlidePlaceholders — bulk-fill placeholders by type token.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  getSlideBody,
  getSlideTitle,
  loadPresentation,
  setSlidePlaceholders,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: setSlidePlaceholders', () => {
  it('fills title + body in one call', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    const slide = addSlide(pres, { layout });
    setSlidePlaceholders(slide, {
      title: 'Quarterly Status',
      body: 'On track\nAhead of plan',
    });
    expect(getSlideTitle(slide)).toBe('Quarterly Status');
    expect(getSlideBody(slide)).toContain('On track');
  });

  it('silently skips unknown placeholder types', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    const slide = addSlide(pres, { layout });
    setSlidePlaceholders(slide, {
      title: 'A',
      // No 'ftr' placeholder on this layout; should be skipped.
      ftr: 'footer text',
    });
    expect(getSlideTitle(slide)).toBe('A');
  });
});

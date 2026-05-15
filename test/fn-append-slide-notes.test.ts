// appendSlideNotes — preserve existing notes and add another line.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  appendSlideNotes,
  getSlideNotes,
  getSlides,
  loadPresentation,
  setSlideNotes,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: appendSlideNotes', () => {
  it('writes the first line when no notes exist yet', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(getSlideNotes(slide)).toBeNull();
    appendSlideNotes(slide, 'first');
    expect(getSlideNotes(slide)).toBe('first');
  });

  it('preserves existing notes and adds a newline', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    setSlideNotes(slide, 'opening');
    appendSlideNotes(slide, 'reminder');
    expect(getSlideNotes(slide)).toBe('opening\nreminder');

    // Chaining works.
    appendSlideNotes(slide, 'last');
    expect(getSlideNotes(slide)).toBe('opening\nreminder\nlast');
  });
});

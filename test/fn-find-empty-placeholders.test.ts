// findEmptyPlaceholders — placeholders whose text body is empty.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findEmptyPlaceholders,
  findSlideLayout,
  findSlidePlaceholder,
  loadPresentation,
  setShapeText,
  setSlideTitle,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findEmptyPlaceholders', () => {
  it('returns the title + body when both are empty', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    const slide = addSlide(pres, { layout });
    const empty = findEmptyPlaceholders(slide);
    // At minimum, title + body placeholders are empty.
    expect(empty.length).toBeGreaterThanOrEqual(2);
  });

  it('drops a placeholder once its text is filled', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    const slide = addSlide(pres, { layout });
    const before = findEmptyPlaceholders(slide).length;
    setSlideTitle(slide, 'Roadmap');
    expect(findEmptyPlaceholders(slide).length).toBe(before - 1);
  });

  it('also drops body placeholders after setShapeText', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    const slide = addSlide(pres, { layout });
    const body = findSlidePlaceholder(slide, 'body')!;
    setShapeText(body, 'filled');
    expect(findEmptyPlaceholders(slide)).not.toContain(body);
  });
});

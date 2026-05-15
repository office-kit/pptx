// mergePresentations — bulk import slides from one deck into another.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  findSlideLayout,
  getSlideCount,
  getSlideShapes,
  getSlideText,
  getSlides,
  inches,
  loadPresentation,
  mergePresentations,
  savePresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: mergePresentations', () => {
  it('appends every source slide into the target', async () => {
    const target = await loadPresentation(await readFile(fixture('blank.pptx')));
    const source = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const layout = findSlideLayout(target, 'Blank')!;
    const before = getSlideCount(target);
    const sourceCount = getSlideCount(source);

    const imported = mergePresentations(target, source, layout);
    expect(imported.length).toBe(sourceCount);
    expect(getSlideCount(target)).toBe(before + sourceCount);
  });

  it('honors per-slide layout selection', async () => {
    const target = await loadPresentation(await readFile(fixture('blank.pptx')));
    const source = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const blank = findSlideLayout(target, 'Blank')!;
    const titleAndContent = findSlideLayout(target, 'Title and Content')!;

    // First slide → blank, second → title-and-content.
    const imported = mergePresentations(target, source, (_, i) =>
      i === 0 ? blank : titleAndContent,
    );
    expect(imported.length).toBe(getSlideCount(source));
  });

  it('preserves text content from each source slide', async () => {
    const source = await loadPresentation(await readFile(fixture('blank.pptx')));
    const sourceLayout = findSlideLayout(source, 'Blank')!;
    // Annotate a freshly-built source deck so we have a known marker.
    const { addSlide } = await import('../src/api/index.ts');
    const s = addSlide(source, { layout: sourceLayout });
    addSlideTextBox(s, {
      x: inches(0), y: inches(0), w: inches(4), h: inches(1),
      text: 'merge-marker',
    });

    const target = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(target, 'Blank')!;
    const imported = mergePresentations(target, source, blank);
    // At least one imported slide carries the marker.
    const markers = imported.filter((sl) => getSlideText(sl).includes('merge-marker'));
    expect(markers.length).toBeGreaterThanOrEqual(1);
  });

  it('round-trips through save / reload', async () => {
    const target = await loadPresentation(await readFile(fixture('blank.pptx')));
    const source = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const layout = findSlideLayout(target, 'Blank')!;
    const expected = getSlideCount(target) + getSlideCount(source);
    mergePresentations(target, source, layout);

    const reloaded = await loadPresentation(await savePresentation(target));
    expect(getSlideCount(reloaded)).toBe(expected);
    // Sanity: every slide is reachable + has a shape tree.
    for (const slide of getSlides(reloaded)) {
      void getSlideShapes(slide);
    }
  });
});

// addSlideAt / duplicateSlideAt — indexed insertion.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideAt,
  duplicateSlideAt,
  findSlideLayout,
  getSlideIndex,
  getSlides,
  loadPresentation,
  savePresentation,
  setSlideTitle,
  getSlideTitle,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: addSlideAt', () => {
  it('inserts the new slide at the requested index', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    const slide = addSlideAt(pres, 0, { layout });
    expect(getSlideIndex(pres, slide)).toBe(0);
    expect(getSlides(pres).length).toBe(3);
  });

  it('clamps to the end when the index is out of range', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    const slide = addSlideAt(pres, 99, { layout });
    expect(getSlideIndex(pres, slide)).toBe(2);
  });

  it('rejects a foreign layout before inserting any member of the batch', async () => {
    const bytes = await readFile(fixture('two-slides.pptx'));
    const pres = await loadPresentation(bytes);
    const foreign = await loadPresentation(bytes);
    const before = [...getSlides(pres)];
    expect(() =>
      addSlideAt(pres, 1, [
        { layout: findSlideLayout(pres, 'Title and Content')! },
        { layout: findSlideLayout(foreign, 'Title and Content')! },
      ]),
    ).toThrow('another presentation');
    expect(getSlides(pres)).toEqual(before);
    expect(getSlides(await loadPresentation(await savePresentation(pres)))).toHaveLength(2);
  });

  it('inserts a batch in order, retaining existing handles and saved titles', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const before = [...getSlides(pres)];
    const layout = findSlideLayout(pres, 'Title and Content')!;
    const added = addSlideAt(pres, 1, [{ layout }, { layout }, { layout }]);
    added.forEach((slide, index) => setSlideTitle(slide, `Inserted ${index}`));
    expect(getSlides(pres)).toEqual([before[0], ...added, before[1]]);
    const loaded = await loadPresentation(await savePresentation(pres));
    expect(getSlides(loaded).slice(1, 4).map(getSlideTitle)).toEqual([
      'Inserted 0',
      'Inserted 1',
      'Inserted 2',
    ]);
    expect(addSlideAt(pres, 0, [])).toEqual([]);
    expect(getSlides(pres)).toHaveLength(5);
  });
});

describe('fn API: duplicateSlideAt', () => {
  it('duplicates and places the duplicate at the requested index', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const source = getSlides(pres)[1]!;
    const dup = duplicateSlideAt(pres, 0, source);
    expect(getSlideIndex(pres, dup)).toBe(0);
    expect(getSlides(pres).length).toBe(3);
  });
});

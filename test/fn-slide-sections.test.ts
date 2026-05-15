// Slide sections (p14:sectionLst).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  Presentation,
  _internalPackageOf,
  getSlideSections,
  getSlides,
  loadPresentation,
  savePresentation,
  setSlideSections,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: slide sections', () => {
  it('reports empty array when no sectionLst is present', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getSlideSections(pres)).toEqual([]);
  });

  it('writes + reads back named sections', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slides = getSlides(pres);
    setSlideSections(pres, [
      { name: 'Intro', slides: [slides[0]!] },
      { name: 'Main', slides: [slides[1]!] },
    ]);
    const sections = getSlideSections(pres);
    expect(sections.map((s) => s.name)).toEqual(['Intro', 'Main']);
    expect(sections[0]!.slides).toHaveLength(1);
    expect(sections[1]!.slides).toHaveLength(1);

    // Round-trip through save → reload.
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    expect(getSlideSections(reloaded).map((s) => s.name)).toEqual(['Intro', 'Main']);
  });

  it('empty section list drops the extension entirely', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    setSlideSections(pres, [{ name: 'X', slides: [] }]);
    expect(getSlideSections(pres)).toHaveLength(1);
    setSlideSections(pres, []);
    expect(getSlideSections(pres)).toEqual([]);

    // The extLst should have no sectionLst entry now.
    const cls = await Presentation.load(await savePresentation(pres));
    const pkg = _internalPackageOf(cls);
    const presPart = pkg.parts.find((p) => p.name === '/ppt/presentation.xml');
    expect(presPart).not.toBeUndefined();
    const xml = new TextDecoder().decode(presPart!.data);
    expect(xml).not.toContain('521415D9-36F7-43E2-AB2F-B90AF26B5E84');
  });
});

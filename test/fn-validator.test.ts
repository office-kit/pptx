// Lightweight invariant validator (`validatePresentation`).
//
// Confirms that a real, well-formed deck reports no issues, then forges
// known-bad packages and asserts the corresponding error fires.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  Presentation,
  _internalPackageOf,
  loadPresentation,
  savePresentation,
  validatePresentation,
} from '../src/api/index.ts';
import { partName } from '../src/internal/opc/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: validatePresentation', () => {
  it('reports no issues for a clean fixture', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const issues = validatePresentation(pres);
    expect(issues).toEqual([]);
  });

  it('reports a missing layout rel as an error', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    // Round-trip into the class-API to grab the package, drop the layout
    // rel from slide1, save, reload, run the validator on the fn handle.
    const clsApi = await Presentation.load(await savePresentation(pres));
    const pkg = _internalPackageOf(clsApi);
    const slideName = partName('/ppt/slides/slide1.xml');
    const rels = pkg.getRels(slideName);
    expect(rels).not.toBeNull();
    rels!.items = rels!.items.filter(
      (r) => !r.type.endsWith('/slideLayout'),
    );
    pkg.setRels(slideName, rels!);

    const broken = await loadPresentation(await clsApi.save());
    const issues = validatePresentation(broken);
    expect(issues.some((i) => i.message.includes('slideLayout'))).toBe(true);
  });

  it('reports a dangling slide rel as an error', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const clsApi = await Presentation.load(await savePresentation(pres));
    const pkg = _internalPackageOf(clsApi);
    // Delete the second slide part but leave its rel + sldId in place.
    pkg.removePart(partName('/ppt/slides/slide2.xml'));

    const broken = await loadPresentation(await clsApi.save());
    const issues = validatePresentation(broken);
    expect(issues.some((i) => i.message.includes('slide2.xml'))).toBe(true);
  });

  it('reports nothing extra after a successful addSlide round-trip', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const clsApi = await Presentation.load(await savePresentation(pres));
    const layout = clsApi.slideLayouts.find((l) => l.name === 'Title and Content');
    if (!layout) throw new Error('expected Title and Content layout');
    clsApi.addSlide({ layout });
    const after = await loadPresentation(await clsApi.save());
    expect(validatePresentation(after)).toEqual([]);
  });
});

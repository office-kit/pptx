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

  it('flags a dangling image rel after media removal', async () => {
    const { addSlideImage, getSlides, inches } = await import('../src/api/index.ts');
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    // Add an image to install a slide → image rel.
    addSlideImage(
      slide,
      new Uint8Array([
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
        0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00,
        0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
        0x42, 0x60, 0x82,
      ]),
      { x: inches(0), y: inches(0), w: inches(1), h: inches(1), format: 'png' },
    );

    // Save → reload → remove the media part directly via the class
    // package (the slide still references it).
    const cls = await Presentation.load(await savePresentation(pres));
    const pkg = _internalPackageOf(cls);
    const media = pkg.parts.find((p) => /^\/ppt\/media\/image\d+\.png$/.test(p.name));
    if (!media) throw new Error('expected media part');
    pkg.removePart(media.name);
    const broken = await loadPresentation(await cls.save());
    const issues = validatePresentation(broken);
    expect(issues.some((i) => i.message.includes('image'))).toBe(true);
  });
});

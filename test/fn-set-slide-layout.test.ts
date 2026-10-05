// setSlideLayout — rebind a slide's layout rel without touching its content.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  INTERNAL_PACKAGE,
  LAYOUT_PART_NAME,
  SLIDE_PART_NAME,
} from '../src/api/_internal-symbols.ts';
import { partName, resolveTarget } from '../src/internal/opc/index.ts';
import { REL_TYPES } from '../src/internal/presentationml/index.ts';
import {
  addSlide,
  findSlideLayout,
  getSlideLayout,
  getSlideLayoutName,
  getSlideLayoutPartName,
  getShapeXmlString,
  getSlideShapes,
  getSlides,
  loadPresentation,
  savePresentation,
  setSlideLayout,
  setShapeHyperlink,
  setShapeText,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: setSlideLayout', () => {
  it('rebinds the slide to a new layout (content preserved)', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const titleAndContent = findSlideLayout(pres, 'Title and Content');
    expect(titleAndContent).not.toBeNull();
    const slide = addSlide(pres, { layout: titleAndContent! });
    const beforeShapeCount = getSlideShapes(slide).length;

    const blank = findSlideLayout(pres, 'Blank');
    expect(blank).not.toBeNull();
    setSlideLayout(slide, blank!);

    // Same content (no shapes added or removed).
    expect(getSlideShapes(slide).length).toBe(beforeShapeCount);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reLayout = getSlideLayout(getSlides(reloaded)[0]!);
    expect(reLayout && getSlideLayoutName(reLayout)).toBe('Blank');
  });

  it.each([
    '/custom/layout.xml',
    '/ppt/slideLayouts/nested/layout.xml',
    '/layout.xml',
    '/custom/slideLayout7.xml',
  ])(
    'preserves the actual layout part path %s when replacing or adding a relationship',
    async (path) => {
      const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
      const originalLayout = findSlideLayout(pres, 'Blank')!;
      const pkg = pres[INTERNAL_PACKAGE];
      const oldName = originalLayout[LAYOUT_PART_NAME];
      const old = pkg.getPart(oldName)!;
      const newName = partName(path);
      pkg.addPart(newName, old.contentType, old.data);
      const layoutRels = pkg.getRels(oldName)!;
      pkg.setRels(newName, {
        items: layoutRels.items.map((rel) => ({
          ...rel,
          target: rel.targetMode === 'External' ? rel.target : resolveTarget(oldName, rel.target),
        })),
      });
      // The handle represents an imported layout outside the conventional folder.
      const layout = { ...originalLayout, [LAYOUT_PART_NAME]: newName };
      const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Title and Content')! });
      const shapes = [...getSlideShapes(slide)];
      // A link only keeps its relationship while a run carries it, so the
      // placeholder gets text before the link — otherwise the rel this test
      // watches is released as unused the moment it is created.
      setShapeText(shapes[0]!, 'Layout review');
      setShapeHyperlink(shapes[0]!, 'https://example.com/layout-review');
      const xml = shapes.map(getShapeXmlString);
      for (const hasExisting of [true, false]) {
        const rels = pkg.getRels(slide[SLIDE_PART_NAME])!;
        if (!hasExisting) {
          rels.items = rels.items.filter((rel) => rel.type !== REL_TYPES.slideLayout);
          pkg.setRels(slide[SLIDE_PART_NAME], rels);
        }
        const previous = rels.items.find((rel) => rel.type === REL_TYPES.slideLayout)?.id;
        const others = rels.items
          .filter((rel) => rel.type !== REL_TYPES.slideLayout)
          .map((rel) => ({ ...rel }));
        expect(others).not.toHaveLength(0);
        setSlideLayout(slide, layout);
        expect(getSlideLayoutPartName(getSlideLayout(slide)!)).toBe(path);
        const after = pkg.getRels(slide[SLIDE_PART_NAME])!;
        expect(after.items.filter((rel) => rel.type !== REL_TYPES.slideLayout)).toEqual(others);
        if (previous)
          expect(after.items.find((rel) => rel.type === REL_TYPES.slideLayout)!.id).toBe(previous);
        expect(getSlideShapes(slide)).toEqual(shapes);
        expect(shapes.map(getShapeXmlString)).toEqual(xml);
        const loaded = await loadPresentation(await savePresentation(pres));
        const savedSlide = getSlides(loaded).at(-1)!;
        expect(getSlideLayoutPartName(getSlideLayout(savedSlide)!)).toBe(path);
        expect(getSlideShapes(savedSlide).map(getShapeXmlString)).toEqual(xml);
      }
    },
  );

  it('rejects a missing layout without changing the slide relationships', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Title and Content')! });
    const layout = {
      ...findSlideLayout(pres, 'Blank')!,
      [LAYOUT_PART_NAME]: partName('/missing/layout.xml'),
    };
    const before = JSON.stringify(pres[INTERNAL_PACKAGE].getRels(slide[SLIDE_PART_NAME]));
    expect(() => setSlideLayout(slide, layout)).toThrow('not in package');
    expect(JSON.stringify(pres[INTERNAL_PACKAGE].getRels(slide[SLIDE_PART_NAME]))).toBe(before);
  });
});

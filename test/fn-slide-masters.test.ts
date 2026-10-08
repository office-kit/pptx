// The Slide Master tab's structural edits: inserting and deleting masters and
// layouts, renaming and preserving masters, the master's placeholders, the
// layouts' Title / Footers / Hide Background Graphics settings and Insert
// Placeholder. Every result is saved, reloaded, checked by the package
// validator and validated part by part against the ECMA-376 schemas.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideLayout,
  addSlideLayoutPlaceholder,
  addSlideMaster,
  createPresentation,
  emu,
  findSlidePlaceholder,
  getSlideLayoutName,
  getSlideLayoutPartName,
  getSlideLayoutPlaceholders,
  getSlideLayouts,
  getSlideMasterLayouts,
  getSlideMasterName,
  getSlideMasterPartName,
  getSlideMasterPartNames,
  getSlideMasterPlaceholders,
  getSlides,
  isSlideLayoutBackgroundGraphicsHidden,
  isSlideMasterPreserved,
  loadPresentation,
  type PresentationData,
  readPackagePart,
  removeSlideLayout,
  removeSlideMaster,
  setSlideLayoutBackgroundGraphicsHidden,
  setSlideLayoutFootersIncluded,
  setSlideLayoutTitleIncluded,
  setSlideMasterName,
  setSlideMasterPlaceholderIncluded,
  setSlideMasterPreserved,
  type SlideLayoutData,
} from '../src/api/index.ts';
import { expectPackageValid } from './lib/expect-package-valid.ts';

const fixture = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));
const template = async (): Promise<PresentationData> => loadPresentation(await readFile(fixture));
const decode = (bytes: Uint8Array | null): string => new TextDecoder().decode(bytes!);

const firstMaster = (pres: PresentationData): string => getSlideMasterPartNames(pres)[0]!;
const types = (placeholders: ReadonlyArray<{ type: string | null }>) =>
  placeholders.map((p) => p.type);
const layoutNamed = (pres: PresentationData, master: string, name: string): SlideLayoutData =>
  getSlideMasterLayouts(pres, master).find((l) => getSlideLayoutName(l) === name)!;

describe('slide masters', () => {
  it('lists a master’s layouts in its sldLayoutIdLst order', async () => {
    const pres = await template();
    const names = getSlideMasterLayouts(pres, firstMaster(pres)).map(getSlideLayoutName);
    expect(names[0]).toBe('Title Slide');
    expect(names).toHaveLength(getSlideLayouts(pres).length);
  });

  it('reads the master’s placeholders with their positions', async () => {
    const pres = createPresentation();
    const placeholders = getSlideMasterPlaceholders(pres, firstMaster(pres));
    expect(types(placeholders)).toEqual(['title', 'body']);
    expect(placeholders[0]!.bounds).not.toBeNull();
  });

  it('Insert Slide Master adds the default master with eleven layouts and its own theme', async () => {
    const pres = await template();
    const before = getSlideMasterPartNames(pres);
    const master = addSlideMaster(pres);

    expect(getSlideMasterPartNames(pres)).toEqual([...before, master]);
    const layouts = getSlideMasterLayouts(pres, master);
    expect(layouts.map(getSlideLayoutName)).toEqual([
      'Title Slide',
      'Title and Content',
      'Section Header',
      'Two Content',
      'Comparison',
      'Title Only',
      'Blank',
      'Content with Caption',
      'Picture with Caption',
      'Title and Vertical Text',
      'Vertical Title and Text',
    ]);
    for (const layout of layouts) expect(getSlideMasterPartName(layout)).toBe(master);
    expect(types(getSlideMasterPlaceholders(pres, master))).toEqual([
      'title',
      'body',
      'dt',
      'ftr',
      'sldNum',
    ]);
    expect(getSlideMasterName(pres, master)).toBe('Custom Design');
    expect(isSlideMasterPreserved(pres, master)).toBe(true);
    // A second insert is numbered as the reference desktop app numbers it.
    expect(getSlideMasterName(pres, addSlideMaster(pres))).toBe('1_Custom Design');

    const reloaded = await expectPackageValid(pres);
    expect(getSlideMasterPartNames(reloaded)).toHaveLength(before.length + 2);
    // Master and layout ids share one space and never repeat.
    const ids = [
      ...decode(readPackagePart(reloaded, '/ppt/presentation.xml')).matchAll(
        /<p:sldMasterId id="(\d+)"/g,
      ),
      ...getSlideMasterPartNames(reloaded).flatMap((name) => [
        ...decode(readPackagePart(reloaded, name)).matchAll(/<p:sldLayoutId id="(\d+)"/g),
      ]),
    ].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it('scales the inserted master to the slide size', async () => {
    const pres = createPresentation({ size: '4:3' });
    const master = addSlideMaster(pres);
    const title = getSlideMasterPlaceholders(pres, master)[0]!.bounds!;
    // 838200 EMU of a 12192000 EMU wide slide, on a 9144000 EMU one.
    expect(title.x).toBe(Math.round((838200 * 9144000) / 12192000));
    await expectPackageValid(pres);
  });

  it('a slide can use a layout of the inserted master', async () => {
    const pres = await template();
    const master = addSlideMaster(pres);
    const slide = addSlide(pres, { layout: layoutNamed(pres, master, 'Comparison') });
    await expectPackageValid(pres);
    expect(getSlides(pres)).toContain(slide);
  });

  it('renames a master through its theme', async () => {
    const pres = await template();
    const master = firstMaster(pres);
    setSlideMasterName(pres, master, 'ブランド');
    const reloaded = await expectPackageValid(pres);
    expect(getSlideMasterName(reloaded, master)).toBe('ブランド');
  });

  it('Preserve sets and clears the preserve attribute', async () => {
    const pres = await template();
    const master = firstMaster(pres);
    expect(isSlideMasterPreserved(pres, master)).toBe(false);
    setSlideMasterPreserved(pres, master, true);
    expect(isSlideMasterPreserved(await expectPackageValid(pres), master)).toBe(true);
    setSlideMasterPreserved(pres, master, false);
    expect(decode(readPackagePart(pres, master))).not.toContain('preserve=');
  });

  it('deletes an unused master with its layouts and theme', async () => {
    const pres = await template();
    const layoutsBefore = getSlideLayouts(pres).length;
    const master = addSlideMaster(pres);
    const theme = '/ppt/theme/theme2.xml';
    expect(readPackagePart(pres, theme)).not.toBeNull();

    removeSlideMaster(pres, master);

    expect(getSlideMasterPartNames(pres)).toEqual([firstMaster(pres)]);
    expect(getSlideLayouts(pres)).toHaveLength(layoutsBefore);
    expect(readPackagePart(pres, master)).toBeNull();
    expect(readPackagePart(pres, theme)).toBeNull();
    await expectPackageValid(pres);
  });

  it('refuses to delete a master that slides use, or the only master', async () => {
    const pres = await template();
    expect(() => removeSlideMaster(pres, firstMaster(pres))).toThrow(/at least one/);
    const master = addSlideMaster(pres);
    addSlide(pres, { layout: getSlideMasterLayouts(pres, master)[0]! });
    expect(() => removeSlideMaster(pres, master)).toThrow(/slides still use/);
  });

  it('Master Layout removes and restores the master’s placeholders', async () => {
    const pres = createPresentation();
    const master = addSlideMaster(pres);
    setSlideMasterPlaceholderIncluded(pres, master, 'ftr', false);
    setSlideMasterPlaceholderIncluded(pres, master, 'title', false);
    expect(types(getSlideMasterPlaceholders(pres, master))).toEqual(['body', 'dt', 'sldNum']);

    setSlideMasterPlaceholderIncluded(pres, master, 'ftr', true);
    setSlideMasterPlaceholderIncluded(pres, master, 'title', true);
    const restored = getSlideMasterPlaceholders(pres, master);
    expect(types(restored)).toEqual(['title', 'body', 'dt', 'ftr', 'sldNum']);
    expect(restored[3]!.bounds).toEqual({ x: 4038600, y: 6356350, w: 4114800, h: 365125 });
    // Shape ids stay unique within the master.
    const ids = [...decode(readPackagePart(pres, master)).matchAll(/<p:cNvPr id="(\d+)"/g)].map(
      (m) => m[1],
    );
    expect(new Set(ids).size).toBe(ids.length);
    await expectPackageValid(pres);
  });
});

describe('slide layouts', () => {
  it('Insert Layout adds a Custom Layout after the given position', async () => {
    const pres = await template();
    const master = firstMaster(pres);
    const layout = addSlideLayout(pres, master, { index: 1 });

    expect(getSlideLayoutName(layout)).toBe('Custom Layout');
    expect(getSlideMasterLayouts(pres, master).map(getSlideLayoutPartName)[1]).toBe(
      getSlideLayoutPartName(layout),
    );
    expect(types(getSlideLayoutPlaceholders(layout))).toEqual(['title', 'dt', 'ftr', 'sldNum']);
    expect(getSlideLayoutName(addSlideLayout(pres, master))).toBe('1_Custom Layout');

    const reloaded = await expectPackageValid(pres);
    expect(getSlideMasterLayouts(reloaded, master).map(getSlideLayoutName).slice(0, 2)).toEqual([
      'Title Slide',
      'Custom Layout',
    ]);
    // The reference desktop app marks layouts it inserts as user-drawn and preserved.
    expect(decode(readPackagePart(reloaded, getSlideLayoutPartName(layout)))).toContain(
      'preserve="1" userDrawn="1"',
    );
  });

  it('deletes an unused layout, and refuses one that slides use', async () => {
    const pres = await template();
    const master = firstMaster(pres);
    const used = getSlideMasterLayouts(pres, master)[0]!;
    addSlide(pres, { layout: used });
    expect(() => removeSlideLayout(used)).toThrow(/slides still use/);

    const spare = getSlideMasterLayouts(pres, master)[1]!;
    const count = getSlideMasterLayouts(pres, master).length;
    removeSlideLayout(spare);
    expect(getSlideMasterLayouts(pres, master)).toHaveLength(count - 1);
    expect(readPackagePart(pres, getSlideLayoutPartName(spare))).toBeNull();
    await expectPackageValid(pres);
  });

  it('refuses to delete a master’s only layout', async () => {
    const pres = createPresentation();
    const master = firstMaster(pres);
    const layouts = getSlideMasterLayouts(pres, master);
    removeSlideLayout(layouts[0]!);
    removeSlideLayout(layouts[1]!);
    expect(() => removeSlideLayout(layouts[2]!)).toThrow(/at least one layout/);
  });

  it('the Title and Footers checkboxes remove and restore placeholders', async () => {
    const pres = await template();
    const layout = layoutNamed(pres, firstMaster(pres), 'Title and Content');
    setSlideLayoutTitleIncluded(layout, false);
    setSlideLayoutFootersIncluded(layout, false);
    expect(types(getSlideLayoutPlaceholders(layout))).toEqual([null]);

    setSlideLayoutTitleIncluded(layout, true);
    setSlideLayoutFootersIncluded(layout, true);
    expect(types(getSlideLayoutPlaceholders(layout))).toEqual([
      'title',
      null,
      'dt',
      'ftr',
      'sldNum',
    ]);
    expect(
      getSlideLayoutPlaceholders(layout)
        .slice(2)
        .map((p) => p.idx),
    ).toEqual([10, 11, 12]);
    await expectPackageValid(pres);
  });

  it('Hide Background Graphics writes showMasterSp', async () => {
    const pres = await template();
    const layout = layoutNamed(pres, firstMaster(pres), 'Title Only');
    setSlideLayoutBackgroundGraphicsHidden(layout, true);
    expect(isSlideLayoutBackgroundGraphicsHidden(layout)).toBe(true);
    const reloaded = await expectPackageValid(pres);
    expect(
      isSlideLayoutBackgroundGraphicsHidden(
        getSlideLayouts(reloaded).find((l) => getSlideLayoutName(l) === 'Title Only')!,
      ),
    ).toBe(true);
    setSlideLayoutBackgroundGraphicsHidden(layout, false);
    expect(isSlideLayoutBackgroundGraphicsHidden(layout)).toBe(false);
  });

  it('Insert Placeholder adds each kind with a fresh idx', async () => {
    const pres = await template();
    const layout = layoutNamed(pres, firstMaster(pres), 'Title Only');
    const kinds = [
      'content',
      'verticalContent',
      'text',
      'verticalText',
      'picture',
      'chart',
      'table',
      'smartArt',
      'media',
      'onlineImage',
    ] as const;
    const bounds = { x: emu(914400), y: emu(1828800), w: emu(3657600), h: emu(2743200) };
    const indices = kinds.map((kind) => addSlideLayoutPlaceholder(layout, kind, bounds));

    const placeholders = getSlideLayoutPlaceholders(layout);
    expect(indices.map((i) => placeholders[i]!.type)).toEqual([
      null,
      null,
      'body',
      'body',
      'pic',
      'chart',
      'tbl',
      'dgm',
      'media',
      'clipArt',
    ]);
    const idx = indices.map((i) => placeholders[i]!.idx);
    expect(new Set(idx).size).toBe(idx.length);
    expect(Math.min(...(idx as number[]))).toBeGreaterThan(12);
    expect(placeholders[indices[4]!]!.bounds).toEqual(bounds);

    // A slide added on the layout gets the new placeholders.
    const slide = addSlide(pres, { layout });
    expect(findSlidePlaceholder(slide, 'pic')).not.toBeNull();
    expect(findSlidePlaceholder(slide, 'chart')).not.toBeNull();
    await expectPackageValid(pres);
  });

  it('Insert Placeholder rejects out-of-range bounds', async () => {
    const pres = await template();
    const layout = getSlideLayouts(pres)[0]!;
    expect(() =>
      addSlideLayoutPlaceholder(layout, 'text', { x: emu(0), y: emu(0), w: emu(-1), h: emu(1) }),
    ).toThrow(RangeError);
  });
});

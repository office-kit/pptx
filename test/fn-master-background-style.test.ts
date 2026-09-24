import { readFile } from 'node:fs/promises';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { expect, it } from 'vitest';
import {
  getEffectiveColorMap,
  getSlides,
  loadPresentation,
  savePresentation,
  setSlideMasterBackgroundStyle,
} from '../src/api/index.ts';

it.each([
  [1, 1001, 'bg1', false],
  [2, 1001, 'bg2', false],
  [3, 1001, 'bg2', true],
  [4, 1001, 'bg1', true],
  [6, 1002, 'bg2', false],
  [7, 1002, 'bg2', true],
  [12, 1003, 'bg1', true],
] as const)(
  'writes native background style %i to the shared master',
  async (style, index, color, dark) => {
    const bytes = await readFile(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url));
    const before = unzipSync(bytes);
    const pres = await loadPresentation(bytes);
    setSlideMasterBackgroundStyle(getSlides(pres)[1]!, style);
    const saved = unzipSync(await savePresentation(pres));
    const master = strFromU8(saved['ppt/slideMasters/slideMaster1.xml']!);
    expect(master).toContain(`<p:bgRef idx="${index}"><a:schemeClr val="${color}"/></p:bgRef>`);
    for (const [name, value] of Object.entries(before)) {
      if (/^ppt\/(slides|slideLayouts|theme)\/[^/]+\.xml$/.test(name))
        expect(strFromU8(saved[name]!)).toBe(strFromU8(value));
    }
    const reloaded = await loadPresentation(zipSync(saved));
    for (const slide of getSlides(reloaded)) {
      expect(getEffectiveColorMap(slide)).toMatchObject({
        bg1: dark ? 'dk1' : 'lt1',
        tx1: dark ? 'lt1' : 'dk1',
        bg2: dark ? 'dk2' : 'lt2',
        tx2: dark ? 'lt2' : 'dk2',
      });
    }
  },
);

it('rejects invalid style numbers and missing theme fills without changing the package', async () => {
  const zip = unzipSync(
    await readFile(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url)),
  );
  zip['ppt/theme/theme1.xml'] = strToU8(
    strFromU8(zip['ppt/theme/theme1.xml']!).replace(
      /<a:bgFillStyleLst>[\s\S]*?<\/a:bgFillStyleLst>/,
      '<a:bgFillStyleLst/>',
    ),
  );
  const pres = await loadPresentation(zipSync(zip));
  const slide = getSlides(pres)[0]!;
  const before = unzipSync(await savePresentation(pres));
  for (const value of [0, 13, 1.5, NaN])
    expect(() => setSlideMasterBackgroundStyle(slide, value)).toThrow(RangeError);
  expect(() => setSlideMasterBackgroundStyle(slide, 6)).toThrow(
    'Theme has no background fill style 2',
  );
  expect(unzipSync(await savePresentation(pres))).toEqual(before);
});

it('uses the selected slide master and preserves unrelated masters and accent mappings', async () => {
  const zip = unzipSync(
    await readFile(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url)),
  );
  const originalMaster = 'ppt/slideMasters/slideMaster1.xml';
  const otherMaster = 'ppt/slideMasters/slideMaster2.xml';
  zip[otherMaster] = strToU8(
    strFromU8(zip[originalMaster]!).replace('accent1="accent1"', 'accent1="accent6"'),
  );
  zip['ppt/slideMasters/_rels/slideMaster2.xml.rels'] =
    zip['ppt/slideMasters/_rels/slideMaster1.xml.rels']!;
  for (const [name, data] of Object.entries(zip)) {
    if (name.startsWith('ppt/slideLayouts/_rels/'))
      zip[name] = strToU8(strFromU8(data).replace('slideMaster1.xml', 'slideMaster2.xml'));
  }
  zip['[Content_Types].xml'] = strToU8(
    strFromU8(zip['[Content_Types].xml']!).replace(
      '</Types>',
      '<Override PartName="/ppt/slideMasters/slideMaster2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/></Types>',
    ),
  );
  const pres = await loadPresentation(zipSync(zip));
  setSlideMasterBackgroundStyle(getSlides(pres)[0]!, 4);
  const saved = unzipSync(await savePresentation(pres));
  expect(strFromU8(saved[originalMaster]!)).toBe(strFromU8(zip[originalMaster]!));
  expect(strFromU8(saved[otherMaster]!)).toContain(
    '<p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef>',
  );
  expect(getEffectiveColorMap(getSlides(pres)[0]!)).toMatchObject({
    bg1: 'dk1',
    accent1: 'accent6',
  });
});

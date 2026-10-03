import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  getSlideBackground,
  getSlides,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';

const fixtureWithBackground = async (background: string) => {
  const zip = unzipSync(
    await readFile(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url)),
  );
  const slide = strFromU8(zip['ppt/slides/slide1.xml']!);
  zip['ppt/slides/slide1.xml'] = strToU8(slide.replace('<p:cSld>', `<p:cSld>${background}`));
  return zipSync(zip);
};

it.each([
  [
    'preset solid fill',
    '<p:bg><p:bgPr><a:solidFill><a:prstClr val="red"/></a:solidFill></p:bgPr></p:bg>',
    { kind: 'solid', color: '#FF0000' },
  ],
  [
    'system solid fill',
    '<p:bg><p:bgPr><a:solidFill><a:sysClr val="windowText" lastClr="123456"/></a:solidFill></p:bgPr></p:bg>',
    { kind: 'solid', color: '#123456' },
  ],
  [
    'preset background reference',
    '<p:bg><p:bgRef idx="1001"><a:prstClr val="red"/></p:bgRef></p:bg>',
    { kind: 'solid', color: '#FF0000' },
  ],
  [
    'system background reference',
    '<p:bg><p:bgRef idx="1001"><a:sysClr val="windowText" lastClr="123456"/></p:bgRef></p:bg>',
    { kind: 'solid', color: '#123456' },
  ],
])('reads and round-trips %s', async (_name, background, expected) => {
  const pres = await loadPresentation(await fixtureWithBackground(background));
  const target = getSlides(pres)[0]!;
  expect(getSlideBackground(target)).toEqual(expected);

  const reloaded = await loadPresentation(await savePresentation(pres));
  expect(getSlideBackground(getSlides(reloaded)[0]!)).toEqual(expected);
});

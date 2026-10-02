import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  getSlideBackground,
  getSlideLayout,
  getSlideLayoutBackground,
  getSlideMasterBackground,
  getSlides,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';

import { renderSlideToRgba } from '../packages/preview/src/node.ts';

const fixture = new URL('./fixtures/minimal/two-slides.pptx', import.meta.url);
const withBackground = (xml: string, background: string): string =>
  /<p:bg\b/.test(xml)
    ? xml.replace(/<p:bg\b[\s\S]*?<\/p:bg>/, background)
    : xml.replace(/(<p:cSld\b[^>]*>)/, `$1${background}`);

it.each([
  ['sRGB', '<a:srgbClr val="336699"><a:tint val="50000"/></a:srgbClr>', '#BEC6D4'],
  ['theme', '<a:schemeClr val="accent1"><a:tint val="50000"/></a:schemeClr>', '#C2CDE1'],
] as const)(
  'applies %s color transforms to a solid slide background',
  async (_, color, expected) => {
    const zip = unzipSync(await readFile(fixture));
    const slideName = 'ppt/slides/slide1.xml';
    const xml = strFromU8(zip[slideName]!);
    zip[slideName] = strToU8(
      xml.replace(
        /(<p:cSld\b[^>]*>)/,
        `$1<p:bg><p:bgPr><a:solidFill>${color}</a:solidFill><a:effectLst/></p:bgPr></p:bg>`,
      ),
    );

    const presentation = await loadPresentation(zipSync(zip));
    const slide = getSlides(presentation)[0]!;
    expect(getSlideBackground(slide)).toEqual({ kind: 'solid', color: expected });

    const reloaded = await loadPresentation(await savePresentation(presentation));
    expect(getSlideBackground(getSlides(reloaded)[0]!)).toEqual({
      kind: 'solid',
      color: expected,
    });
    const { image } = renderSlideToRgba(reloaded, getSlides(reloaded)[0]!, { width: 96 });
    expect(Array.from(image.data.slice(0, 3))).toEqual([
      Number.parseInt(expected.slice(1, 3), 16),
      Number.parseInt(expected.slice(3, 5), 16),
      Number.parseInt(expected.slice(5, 7), 16),
    ]);
  },
);

it('keeps an untransformed scheme background token while exposing alpha separately', async () => {
  const zip = unzipSync(await readFile(fixture));
  const slideName = 'ppt/slides/slide1.xml';
  zip[slideName] = strToU8(
    strFromU8(zip[slideName]!).replace(
      /(<p:cSld\b[^>]*>)/,
      '$1<p:bg><p:bgPr><a:solidFill><a:schemeClr val="accent1"><a:alpha val="50000"/></a:schemeClr></a:solidFill><a:effectLst/></p:bgPr></p:bg>',
    ),
  );
  const presentation = await loadPresentation(zipSync(zip));
  expect(getSlideBackground(getSlides(presentation)[0]!)).toEqual({
    kind: 'solid',
    color: 'scheme:accent1',
    opacity: 0.5,
  });
});

it('applies a slide color-map override before transforming a scheme background', async () => {
  const zip = unzipSync(await readFile(fixture));
  const slideName = 'ppt/slides/slide1.xml';
  const background =
    '<p:bg><p:bgPr><a:solidFill><a:schemeClr val="accent1"><a:tint val="50000"/></a:schemeClr></a:solidFill><a:effectLst/></p:bgPr></p:bg>';
  zip[slideName] = strToU8(
    withBackground(strFromU8(zip[slideName]!), background).replace(
      /<p:clrMapOvr\b[\s\S]*?<\/p:clrMapOvr>/,
      '<p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent2" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>',
    ),
  );
  const presentation = await loadPresentation(zipSync(zip));
  expect(getSlideBackground(getSlides(presentation)[0]!)).toEqual({
    kind: 'solid',
    color: '#E2C2C2',
  });
});

it('applies transforms when the background is inherited from the layout or master', async () => {
  const zip = unzipSync(await readFile(fixture));
  const background =
    '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="336699"><a:tint val="50000"/></a:srgbClr></a:solidFill><a:effectLst/></p:bgPr></p:bg>';
  for (const name of Object.keys(zip).filter((name) =>
    /^ppt\/(slideLayouts\/slideLayout\d+|slideMasters\/slideMaster\d+)\.xml$/.test(name),
  )) {
    zip[name] = strToU8(withBackground(strFromU8(zip[name]!), background));
  }
  const presentation = await loadPresentation(zipSync(zip));
  const slide = getSlides(presentation)[0]!;
  const layout = getSlideLayout(slide)!;
  expect(getSlideLayoutBackground(layout)).toEqual({ kind: 'solid', color: '#BEC6D4' });
  expect(getSlideMasterBackground(presentation, layout)).toEqual({
    kind: 'solid',
    color: '#BEC6D4',
  });
});

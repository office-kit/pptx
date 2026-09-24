import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  getSlideBackgroundGradientFill,
  getSlideLayoutBackgroundGradientFill,
  getSlideMasterBackgroundGradientFill,
  getSlideLayout,
  getSlides,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';

const gradient = `<a:gradFill rotWithShape="0"><a:gsLst><a:gs pos="0"><a:schemeClr val="accent1"><a:alpha val="40000"/><a:lumMod val="70000"/><a:lumOff val="30000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:srgbClr val="123456"><a:lumMod val="80000"/></a:srgbClr></a:gs></a:gsLst><a:path path="rect"><a:fillToRect l="100000" t="100000"/></a:path><a:tileRect l="10000" t="20000" r="30000" b="40000"/></a:gradFill>`;

async function withBackground(fill: string) {
  const zip = unzipSync(
    await readFile(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url)),
  );
  for (const [name, bytes] of Object.entries(zip)) {
    if (!/^ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(name)) continue;
    const xml = strFromU8(bytes).replace(/<p:bg\b[^>]*>[\s\S]*?<\/p:bg>/g, '');
    zip[name] = strToU8(
      xml.replace(/(<p:cSld\b[^>]*>)/, `$1<p:bg><p:bgPr>${fill}</p:bgPr></p:bg>`),
    );
  }
  return loadPresentation(zipSync(zip));
}

function backgrounds(pres: Awaited<ReturnType<typeof loadPresentation>>) {
  const slide = getSlides(pres)[0]!;
  const layout = getSlideLayout(slide)!;
  return [
    getSlideBackgroundGradientFill(slide),
    getSlideLayoutBackgroundGradientFill(layout),
    getSlideMasterBackgroundGradientFill(pres, layout),
  ];
}

describe('gradient background details', () => {
  it('retains stop transforms and rectangular direction on slides, layouts and masters after reload', async () => {
    const pres = await withBackground(gradient);
    const expected = {
      stops: [
        { offset: 0, color: 'scheme:accent1', opacity: 0.4, brightness: 0.3 },
        { offset: 1, color: '#123456', brightness: -0.2 },
      ],
      angleDeg: 0,
      path: 'rect',
      rotateWithShape: false,
      focus: { left: 1, top: 1, right: 0, bottom: 0 },
      tileRect: { left: 0.1, top: 0.2, right: 0.3, bottom: 0.4 },
    };
    for (const value of backgrounds(pres)) {
      expect(value).toMatchObject({
        ...expected,
        stops: [expected.stops[0], { ...expected.stops[1], brightness: expect.closeTo(-0.2) }],
      });
    }
    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(backgrounds(reloaded)).toEqual(backgrounds(pres));
  });

  it('reads scaled linear direction and returns null for non-gradient backgrounds', async () => {
    const linear = gradient.replace(
      /<a:path[\s\S]*?<\/a:path>/,
      '<a:lin ang="2700000" scaled="1"/>',
    );
    for (const value of backgrounds(await withBackground(linear))) {
      expect(value).toMatchObject({ angleDeg: 45, scaled: true });
      expect(value?.path).toBeUndefined();
    }
    expect(
      backgrounds(await withBackground('<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>')),
    ).toEqual([null, null, null]);
  });
});

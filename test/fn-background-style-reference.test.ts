import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  getSlideBackground,
  getSlideBackgroundGradientFill,
  getSlideLayoutBackground,
  getSlideLayoutBackgroundGradientFill,
  getSlideMasterBackground,
  getSlideMasterBackgroundGradientFill,
  getSlideLayout,
  getSlides,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

const referencedGradient =
  '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>';

it.each(['slides', 'slideLayouts', 'slideMasters'] as const)(
  'resolves a theme background gradient on %s without changing the package',
  async (level) => {
    const zip = unzipSync(
      await readFile(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url)),
    );
    for (const [name, bytes] of Object.entries(zip)) {
      if (/^ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(name)) {
        let xml = strFromU8(bytes).replace(/<p:bg\b[^>]*>[\s\S]*?<\/p:bg>/g, '');
        if (name.startsWith(`ppt/${level}/`))
          xml = xml.replace(
            /(<p:cSld\b[^>]*>)/,
            '$1<p:bg><p:bgRef idx="1003"><a:schemeClr val="bg1"/></p:bgRef></p:bg>',
          );
        if (name.startsWith('ppt/slideMasters/'))
          xml = xml.replace('bg1="lt1"', 'bg1="dk1"').replace('tx1="dk1"', 'tx1="lt1"');
        zip[name] = strToU8(xml);
      }
      if (/^ppt\/theme\/[^/]+\.xml$/.test(name))
        zip[name] = strToU8(
          strFromU8(bytes).replace(
            /<a:bgFillStyleLst>[\s\S]*?<\/a:bgFillStyleLst>/,
            `<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>${referencedGradient}${referencedGradient.repeat(3)}</a:bgFillStyleLst>`,
          ),
        );
    }
    const pres = await loadPresentation(zipSync(zip));
    const slide = getSlides(pres)[0]!;
    const layout = getSlideLayout(slide)!;
    const fill =
      level === 'slides'
        ? getSlideBackground(slide)
        : level === 'slideLayouts'
          ? getSlideLayoutBackground(layout)
          : getSlideMasterBackground(pres, layout);
    expect(fill.kind).toBe('gradient');
    const gradient =
      level === 'slides'
        ? getSlideBackgroundGradientFill(slide)
        : level === 'slideLayouts'
          ? getSlideLayoutBackgroundGradientFill(layout)
          : getSlideMasterBackgroundGradientFill(pres, layout);
    expect(gradient).toMatchObject({
      path: 'circle',
      focus: { left: 0.5, top: 0.5, right: 0.5, bottom: 0.5 },
      stops: [
        { offset: 0, resolvedColor: '#7C7C7C' },
        { offset: 1, resolvedColor: '#000000' },
      ],
    });
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    expect(svg).toContain('<radialGradient');
    expect(svg).toContain('stop-color="#7C7C7C"');
    const saved = unzipSync(await savePresentation(pres));
    for (const [name, bytes] of Object.entries(zip)) {
      if (/^ppt\/(slides|slideLayouts|slideMasters|theme)\/[^/]+\.xml$/.test(name))
        expect(strFromU8(saved[name]!)).toBe(strFromU8(bytes));
    }
  },
);

it('uses the owning master theme when another theme is first in the presentation', async () => {
  const zip = unzipSync(
    await readFile(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url)),
  );
  const master = 'ppt/slideMasters/slideMaster1.xml';
  zip['ppt/slideMasters/slideMaster2.xml'] = zip[master]!;
  zip['ppt/slideMasters/_rels/slideMaster2.xml.rels'] = strToU8(
    strFromU8(zip['ppt/slideMasters/_rels/slideMaster1.xml.rels']!).replace(
      'theme1.xml',
      'theme2.xml',
    ),
  );
  zip['ppt/theme/theme2.xml'] = strToU8(
    strFromU8(zip['ppt/theme/theme1.xml']!)
      .replace(/<a:dk1>[\s\S]*?<\/a:dk1>/, '<a:dk1><a:srgbClr val="FF0000"/></a:dk1>')
      .replace(
        /<a:bgFillStyleLst>[\s\S]*?<\/a:bgFillStyleLst>/,
        `<a:bgFillStyleLst>${referencedGradient.repeat(3)}</a:bgFillStyleLst>`,
      ),
  );
  zip['[Content_Types].xml'] = strToU8(
    strFromU8(zip['[Content_Types].xml']!).replace(
      '</Types>',
      '<Override PartName="/ppt/slideMasters/slideMaster2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>',
    ),
  );
  for (const [name, bytes] of Object.entries(zip)) {
    if (/^ppt\/slideLayouts\/_rels\/.+\.rels$/.test(name))
      zip[name] = strToU8(strFromU8(bytes).replace('slideMaster1.xml', 'slideMaster2.xml'));
    if (/^ppt\/slides\/[^/]+\.xml$/.test(name))
      zip[name] = strToU8(
        strFromU8(bytes)
          .replace(/<p:bg\b[^>]*>[\s\S]*?<\/p:bg>/g, '')
          .replace(
            /(<p:cSld\b[^>]*>)/,
            '$1<p:bg><p:bgRef idx="1001"><a:schemeClr val="dk1"/></p:bgRef></p:bg>',
          ),
      );
  }
  const pres = await loadPresentation(zipSync(zip));
  const slide = getSlides(pres)[0]!;
  expect(getSlideBackgroundGradientFill(slide)?.stops.map((stop) => stop.resolvedColor)).toEqual([
    '#FF7C7C',
    '#950000',
  ]);
  expect(renderSlideToSvg(pres, slide, { textLayout: 'svg' })).toContain('stop-color="#FF7C7C"');
});

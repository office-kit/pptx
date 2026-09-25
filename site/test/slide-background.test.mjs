import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import {
  loadPresentation,
  getSlides,
  getSlideLayout,
  setSlideBackground,
  clearSlideBackground,
  setSlideLayoutBackground,
  clearSlideLayoutBackground,
} from '@office-kit/pptx';
import { readSlideBackground } from '../src/lib/editor/core/slide-background.ts';

test('background controls follow slide, layout and master inheritance without looking through a solid override', async () => {
  const zip = unzipSync(
    await readFile(new URL('../../test/fixtures/minimal/two-slides.pptx', import.meta.url)),
  );
  for (const [name, bytes] of Object.entries(zip)) {
    if (!/^ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(name)) continue;
    let xml = strFromU8(bytes).replace(/<p:bg\b[^>]*>[\s\S]*?<\/p:bg>/g, '');
    if (name.startsWith('ppt/slideMasters/'))
      xml = xml.replace(
        /(<p:cSld\b[^>]*>)/,
        '$1<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="123456"/></a:gs><a:gs pos="100000"><a:srgbClr val="ABCDEF"/></a:gs></a:gsLst><a:lin ang="2700000"/></a:gradFill></p:bgPr></p:bg>',
      );
    zip[name] = strToU8(xml);
  }
  const pres = await loadPresentation(zipSync(zip));
  const slide = getSlides(pres)[0];
  const layout = getSlideLayout(slide);
  const inherited = readSlideBackground(pres, slide);
  assert.equal(inherited.fill.kind, 'gradient');
  assert.equal(inherited.gradient.angleDeg, 45);
  setSlideLayoutBackground(layout, '#112233');
  assert.deepEqual(readSlideBackground(pres, slide), {
    fill: { kind: 'solid', color: '#112233' },
    gradient: null,
    pattern: null,
  });
  setSlideBackground(slide, '#445566');
  assert.deepEqual(readSlideBackground(pres, slide), {
    fill: { kind: 'solid', color: '#445566' },
    gradient: null,
    pattern: null,
  });
  clearSlideBackground(slide);
  clearSlideLayoutBackground(layout);
  assert.deepEqual(readSlideBackground(pres, slide), inherited);
});

test('pattern controls read the effective theme colors from layout and master', async () => {
  for (const source of ['slideLayouts', 'slideMasters']) {
    const zip = unzipSync(
      await readFile(new URL('../../test/fixtures/minimal/two-slides.pptx', import.meta.url)),
    );
    for (const [name, bytes] of Object.entries(zip)) {
      if (!/^ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(name)) continue;
      let xml = strFromU8(bytes).replace(/<p:bg\b[^>]*>[\s\S]*?<\/p:bg>/g, '');
      if (name.startsWith(`ppt/${source}/`))
        xml = xml.replace(
          /(<p:cSld\b[^>]*>)/,
          '$1<p:bg><p:bgPr><a:pattFill prst="wave"><a:fgClr><a:schemeClr val="accent2"/></a:fgClr><a:bgClr><a:schemeClr val="bg2"/></a:bgClr></a:pattFill></p:bgPr></p:bg>',
        );
      zip[name] = strToU8(xml);
    }
    const pres = await loadPresentation(zipSync(zip));
    const slide = getSlides(pres)[0];
    const inherited = readSlideBackground(pres, slide, { preserveTheme: true });
    assert.equal(inherited.fill.kind, 'pattern');
    assert.deepEqual(inherited.pattern, {
      preset: 'wave',
      foreground: 'accent2',
      background: 'bg2',
    });
    assert.match(readSlideBackground(pres, slide).pattern.foreground, /^#[0-9A-F]{6}$/);
    setSlideBackground(slide, '#112233');
    assert.equal(readSlideBackground(pres, slide).pattern, null);
  }
});

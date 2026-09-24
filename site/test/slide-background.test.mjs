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
  });
  setSlideBackground(slide, '#445566');
  assert.deepEqual(readSlideBackground(pres, slide), {
    fill: { kind: 'solid', color: '#445566' },
    gradient: null,
  });
  clearSlideBackground(slide);
  clearSlideLayoutBackground(layout);
  assert.deepEqual(readSlideBackground(pres, slide), inherited);
});

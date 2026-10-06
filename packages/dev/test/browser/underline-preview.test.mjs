import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  inches,
  loadPresentation,
  setShapeRunFormat,
} from '@office-kit/pptx';
import { renderSlideToSvg } from '@office-kit/pptx-preview';

test('HTML underline patterns remain distinct without changing text layout', async () => {
  const pres = await loadPresentation(
    await readFile(new URL('../../../../test/fixtures/minimal/blank.pptx', import.meta.url)),
  );
  const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Blank') });
  const box = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(5),
    h: inches(1),
    text: 'Underline words',
  });
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage({ viewport: { width: 960, height: 720 } });
    const images = new Map();
    let originalBounds;
    for (const underline of [
      'none',
      'sng',
      'dbl',
      'heavy',
      'dotted',
      'dottedHeavy',
      'dash',
      'dashHeavy',
      'dashLong',
      'dashLongHeavy',
      'dotDash',
      'dotDashHeavy',
      'dotDotDash',
      'dotDotDashHeavy',
      'wavy',
      'wavyHeavy',
      'wavyDbl',
      'words',
    ]) {
      setShapeRunFormat(box, 0, 0, { size: 36, underline, strike: true, color: '#154687' });
      const svg = renderSlideToSvg(pres, slide, { textLayout: 'foreignObject' });
      await page.setContent(`<style>body{margin:0}</style>${svg}`);
      await page.evaluate(() => document.fonts.ready);
      assert.equal(
        await page.evaluate(
          (source) =>
            new DOMParser().parseFromString(source, 'image/svg+xml').querySelector('parsererror')
              ?.textContent ?? null,
          svg,
        ),
        null,
      );
      const bounds = await page
        .locator('foreignObject')
        .first()
        .evaluate((node) => {
          const range = document.createRange();
          range.selectNodeContents(node);
          const r = range.getBoundingClientRect();
          return { x: r.x, y: r.y, width: r.width, height: r.height };
        });
      originalBounds ??= bounds;
      assert.deepEqual(bounds, originalBounds, `${underline} changed text layout`);
      const png = await page.screenshot();
      const encoded = png.toString('base64');
      assert.ok(!images.has(encoded), `${underline} rendered like ${images.get(encoded)}`);
      images.set(encoded, underline);
    }
  } finally {
    await browser.close();
  }
});

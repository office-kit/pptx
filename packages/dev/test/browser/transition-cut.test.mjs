import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { compile, Presentation, Slide, Text } from '@office-kit/pptx-dsl';
import {
  getSlides,
  savePresentation,
  setSlideTransition,
  setSlideBackground,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'cut through black holds black without fading and cleans up on completion or exit',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-cut-'));
    let preview, browser;
    try {
      const deck = await compile(
        Presentation({
          children: [0, 1, 2].map((i) =>
            Slide({
              children: Text({ x: 1, y: 1, width: 8, height: 2, children: `カット / Cut ${i}` }),
            }),
          ),
        }),
      );
      getSlides(deck).forEach((slide, i) => {
        setSlideBackground(slide, i % 2 ? '#FFD166' : '#73C2FB');
        setSlideTransition(slide, {
          effect: 'cut',
          speed: 'slow',
          thruBlack: i === 1,
          advanceOnClick: false,
        });
      });
      await writeFile(join(dir, 'source.pptx'), await savePresentation(deck));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(join(dir, 'source.pptx'))})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ reducedMotion: 'no-preference' });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.waitForFunction(() => state.slides.length === 3);
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      await page.keyboard.press('ArrowRight');
      const frame = await page.evaluate(() => {
        for (const animation of transitionAnimations) {
          animation.pause();
          animation.currentTime = 500;
        }
        const incoming = slide.shadowRoot.querySelector('.transition-layer:not(.transition-old)');
        const outgoing = slide.shadowRoot.querySelector('.transition-old');
        return {
          count: transitionAnimations.length,
          opacity: incoming && getComputedStyle(incoming).opacity,
          oldVisibility: outgoing && getComputedStyle(outgoing).visibility,
          background: getComputedStyle(slide).backgroundColor,
        };
      });
      assert.deepEqual(frame, {
        count: 1,
        opacity: '0',
        oldVisibility: 'hidden',
        background: 'rgb(0, 0, 0)',
      });
      await page.locator('#stage').click({ position: { x: 5, y: 5 } });
      assert.equal(await page.evaluate(() => index), 1, 'click does not advance a disabled slide');
      await page.evaluate(() => {
        for (const animation of transitionAnimations) animation.finish();
      });
      await page.waitForFunction(() => !transitionCleanup);
      assert.equal(await page.locator('.transition-layer').count(), 0);
      assert.equal(await page.locator('#slide').evaluate((node) => node.style.background), '');
      assert.match(
        await page.locator('#slide').evaluate((node) => node.shadowRoot.textContent),
        /カット \/ Cut 1/,
      );
      await page.keyboard.press('ArrowRight');
      assert.equal(await page.locator('.transition-layer').count(), 0, 'plain cut is immediate');
      await page.keyboard.press('ArrowLeft');
      assert.equal(await page.locator('.transition-layer').count(), 2);
      await page.getByRole('button', { name: 'Exit · Esc', exact: true }).click();
      assert.equal(await page.locator('#slide').evaluate((node) => node.style.background), '');
      assert.equal(await page.locator('.transition-layer').count(), 0);
      await page.emulateMedia({ reducedMotion: 'reduce' });
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      await page.keyboard.press('ArrowLeft');
      await page.keyboard.press('ArrowRight');
      assert.equal(await page.locator('.transition-layer').count(), 0);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

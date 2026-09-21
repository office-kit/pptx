import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { compile, Presentation, Slide, Text } from '@office-kit/pptx-dsl';
import { getSlides, savePresentation, setSlideTransition } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'presentation playback honors saved click and automatic advance settings',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-presentation-'));
    let preview;
    let browser;
    try {
      const deck = await compile(
        Presentation({
          children: ['Timed', 'Manual', 'Long timer', 'End'].map((title) =>
            Slide({ children: Text({ x: 1, y: 1, width: 5, height: 1, children: title }) }),
          ),
        }),
      );
      const slides = getSlides(deck);
      setSlideTransition(slides[0], {
        effect: 'none',
        advanceOnClick: false,
        advanceAfterMs: 1500,
      });
      setSlideTransition(slides[2], {
        effect: 'none',
        advanceOnClick: false,
        advanceAfterMs: 4294967295,
      });
      setSlideTransition(slides[3], { effect: 'none', advanceAfterMs: 0 });
      await writeFile(join(dir, 'source.pptx'), await savePresentation(deck));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(join(dir, 'source.pptx'))})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.waitForFunction(() => state.slides.length === 4);
      const persisted = await (await fetch(preview.url + '/state')).json();
      assert.equal(persisted.transitions[0].advanceAfterMs, 1500);
      assert.equal(persisted.transitions[0].advanceOnClick, false);
      const current = () => page.locator('#count').textContent();
      const wait = (ms) =>
        page.evaluate((ms) => new Promise((resolve) => setTimeout(resolve, ms)), ms);
      // Ordinary preview has no running playback timer.
      await wait(1700);
      assert.equal(await current(), 'Slide 1 of 4');
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      await page.locator('#stage').click({ position: { x: 15, y: 15 } });
      assert.equal(await current(), 'Slide 1 of 4');
      const timer = await page.evaluate(() => advanceTimer);
      await page.evaluate(() => refresh());
      assert.equal(
        await page.evaluate(() => advanceTimer),
        timer,
        'unchanged server refresh must not restart the timer',
      );
      await page.waitForFunction(() => index === 1);
      // No timing on this slide: remain until a click.
      await wait(1700);
      assert.equal(await current(), 'Slide 2 of 4');
      await page.locator('#stage').click({ position: { x: 15, y: 15 } });
      assert.equal(await current(), 'Slide 3 of 4');
      await wait(100);
      assert.equal(
        await current(),
        'Slide 3 of 4',
        'large OOXML delays must not overflow to immediate timeouts',
      );
      // Explicit navigation still works when stage click advance is disabled.
      await page.keyboard.press('ArrowRight');
      await wait(100);
      assert.equal(await current(), 'Slide 4 of 4');
      assert.equal(await page.locator('#present-next').isDisabled(), true);
      await page.keyboard.press('Home');
      await page.getByRole('button', { name: 'Exit · Esc', exact: true }).click();
      await wait(1700);
      assert.equal(
        await current(),
        'Slide 1 of 4',
        'leaving presentation cancels automatic advance',
      );
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      await page.waitForFunction(() => index === 1);
      await page.keyboard.press('Escape');
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

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
  'random transitions vary between visits without changing saved settings',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-random-'));
    let preview, browser;
    try {
      const options = [{ effect: 'none' }, { effect: 'random' }, { effect: 'random' }];
      const deck = await compile(
        Presentation({
          children: options.map((_, i) =>
            Slide({
              children: Text({
                x: 1,
                y: 1,
                width: 8,
                height: 2,
                children: `日本語 / English ${i}`,
              }),
            }),
          ),
        }),
      );
      getSlides(deck).forEach((slide, i) =>
        setSlideTransition(slide, { ...options[i], speed: 'slow' }),
      );
      getSlides(deck).forEach((slide, i) =>
        setSlideBackground(slide, i % 2 ? '#FFD166' : '#73C2FB'),
      );
      await writeFile(join(dir, 'source.pptx'), await savePresentation(deck));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(join(dir, 'source.pptx'))})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({
        viewport: { width: 1280, height: 800 },
        reducedMotion: 'no-preference',
      });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.waitForFunction(() => state.slides.length === 3);
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      const savedTransitions = await page.evaluate(() => state.transitions);
      const signatures = new Set();
      for (const random of [0, 0.5, 0.999999]) {
        await page.evaluate((random) => {
          Math.random = () => random;
        }, random);
        await page.keyboard.press('ArrowRight');
        const result = await page.evaluate(() => {
          const animations = slide.shadowRoot.getAnimations({ subtree: true });
          for (const animation of animations) {
            animation.pause();
            animation.currentTime = 500;
          }
          return {
            layers: slide.shadowRoot.querySelectorAll('.transition-layer').length,
            count: animations.length,
            durations: animations.map((animation) => animation.effect.getTiming().duration),
            signature: animations.map((animation) => animation.effect.getKeyframes()),
          };
        });
        assert.equal(result.layers, 2);
        assert.ok(result.count >= 1);
        assert.ok(result.durations.every((duration) => duration === 1000));
        signatures.add(JSON.stringify(result.signature));
        // Revisiting a random slide picks again, and cancels the unfinished effect.
        await page.keyboard.press('ArrowLeft');
      }
      assert.equal(signatures.size, 3, 'random chooses visually different effects');
      await page.keyboard.press('ArrowRight');
      await page.keyboard.press('ArrowRight');
      await page.evaluate(() => {
        for (const animation of transitionAnimations) animation.finish();
      });
      await page.waitForFunction(() => !transitionCleanup);
      assert.equal(await page.locator('.transition-layer').count(), 0);
      assert.match(
        await page.locator('#slide').evaluate((node) => node.shadowRoot.textContent),
        /日本語 \/ English 2/,
      );
      await page.keyboard.press('ArrowLeft');
      assert.equal(await page.locator('.transition-layer').count(), 2);
      await page.emulateMedia({ reducedMotion: 'reduce' });
      await page.waitForFunction(() => !transitionCleanup);
      for (let i = 0; i >= 0; i--) {
        await page.keyboard.press('ArrowLeft');
        assert.equal(await page.locator('.transition-layer').count(), 0);
      }
      assert.deepEqual(await page.evaluate(() => state.transitions), savedTransitions);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

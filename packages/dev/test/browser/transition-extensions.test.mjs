import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { compile, Presentation, Slide, Text } from '@office-kit/pptx-dsl';
import { getSlides, savePresentation, setSlideTransition } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

// PowerPoint 2010+ transitions have no flat equivalent in the browser; the
// presentation plays the nearest one rather than cutting straight to the slide.
test(
  'PowerPoint 2010+ transitions play an approximation in the presentation',
  { timeout: 90000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-extension-transitions-'));
    let preview, browser;
    try {
      const options = [
        { effect: 'none' },
        { effect: 'morph' },
        { effect: 'vortex', direction: 'r' },
        { effect: 'prism', direction: 'l', isContent: true },
        { effect: 'doors', direction: 'vert' },
        { effect: 'warp', direction: 'in' },
        { effect: 'prstTrans', preset: 'curtains' },
        { effect: 'reveal', direction: 'l', thruBlack: true },
        { effect: 'wheelReverse', spokes: 1 },
        { effect: 'ripple', direction: 'center' },
      ];
      const deck = await compile(
        Presentation({
          children: options.map((_, i) =>
            Slide({ children: Text({ x: 1, y: 1, width: 8, height: 2, children: `Slide ${i}` }) }),
          ),
        }),
      );
      getSlides(deck).forEach((slide, i) =>
        setSlideTransition(slide, { ...options[i], durationMs: 1000 }),
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
      await page
        .frameLocator('#editor-frame')
        .locator('.statusbar')
        .getByRole('button', { name: 'Reading View', exact: true })
        .click();
      await page.waitForFunction((count) => state.slides.length === count, options.length);
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      for (let i = 1; i < options.length; i++) {
        await page.keyboard.press('ArrowRight');
        const result = await page.evaluate(() => ({
          animations: slide.shadowRoot
            .getAnimations({ subtree: true })
            .map((animation) => animation.effect.getTiming().duration),
          layers: slide.shadowRoot.querySelectorAll('.transition-layer').length,
        }));
        assert.ok(result.animations.length > 0, `${options[i].effect} animates`);
        assert.ok(
          result.animations.every((duration) => duration === 1000),
          `${options[i].effect} runs for its p14:dur`,
        );
        assert.equal(result.layers, 2, options[i].effect);
        await page.evaluate(() => {
          for (const animation of transitionAnimations) animation.finish();
        });
        await page.waitForFunction(() => !transitionCleanup);
        assert.match(
          await page.locator('#slide').evaluate((node) => node.shadowRoot.textContent),
          new RegExp(`Slide ${i}`),
        );
      }
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

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
  'blinds and comb reveal alternating bands and clean up on interruption',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-band-transitions-'));
    let preview, browser;
    try {
      const options = [
        { effect: 'none' },
        { effect: 'blinds' },
        { effect: 'blinds', direction: 'vert' },
        { effect: 'comb' },
        { effect: 'comb', direction: 'vert' },
      ];
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
      await page.waitForFunction(() => state.slides.length === 5);
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      for (let i = 1; i < options.length; i++) {
        await page.keyboard.press('ArrowRight');
        const result = await page.evaluate((options) => {
          const animations = slide.shadowRoot.getAnimations({ subtree: true });
          for (const animation of animations) {
            animation.pause();
            animation.currentTime = 500;
          }
          const incoming = slide.shadowRoot.querySelector('.transition-layer:not(.transition-old)');
          const bounds = slide.getBoundingClientRect();
          const visible = (x, y) =>
            slide.shadowRoot
              .elementFromPoint(bounds.left + bounds.width * x, bounds.top + bounds.height * y)
              ?.closest('.transition-layer') === incoming;
          const samples = [];
          for (let band = 0; band < 8; band++) {
            for (const fraction of [0.25, 0.75]) {
              const across = options.effect === 'blinds' ? 0.5 : fraction;
              const along = (band + (options.effect === 'blinds' ? fraction : 0.5)) / 8;
              samples.push(
                (options.effect === 'blinds') !== (options.direction === 'vert')
                  ? visible(along, across)
                  : visible(across, along),
              );
            }
          }
          return {
            count: animations.length,
            layers: slide.shadowRoot.querySelectorAll('.transition-layer').length,
            duration: animations[0]?.effect.getTiming().duration,
            samples,
          };
        }, options[i]);
        assert.equal(result.count, 1, options[i].effect);
        assert.equal(result.layers, 2, 'interruptions remove previous bands');
        assert.equal(result.duration, 1000);
        assert.deepEqual(
          result.samples,
          Array.from({ length: 8 }, (_, band) =>
            options[i].effect === 'comb' && band % 2 ? [false, true] : [true, false],
          ).flat(),
          JSON.stringify(options[i]),
        );
        await page.screenshot({ path: `/tmp/pptx-band-transition-${i}.png` });
      }
      await page.evaluate(() => {
        for (const animation of transitionAnimations) animation.finish();
      });
      await page.waitForFunction(() => !transitionCleanup);
      assert.equal(await page.locator('.transition-layer').count(), 0);
      assert.match(
        await page.locator('#slide').evaluate((node) => node.shadowRoot.textContent),
        /日本語 \/ English 4/,
      );
      await page.keyboard.press('ArrowLeft');
      assert.equal(await page.locator('.transition-layer').count(), 2);
      await page.emulateMedia({ reducedMotion: 'reduce' });
      await page.waitForFunction(() => !transitionCleanup);
      for (let i = 2; i >= 0; i--) {
        await page.keyboard.press('ArrowLeft');
        assert.equal(await page.locator('.transition-layer').count(), 0);
      }
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

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
  'wheel plays saved spoke counts clockwise and cleans up interrupted playback',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-wheel-'));
    let preview, browser;
    try {
      const options = [
        { effect: 'none' },
        ...[undefined, 1, 2, 3, 4, 8, 0, 4294967295].map((spokes) => ({
          effect: 'wheel',
          ...(spokes === undefined ? {} : { spokes }),
        })),
      ];
      const deck = await compile(
        Presentation({
          children: options.map((_, i) =>
            Slide({
              children: Text({
                x: 0.1,
                y: 0.1,
                width: 8,
                height: 0.3,
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
      await page.waitForFunction((count) => state.slides.length === count, options.length);
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      for (let i = 1; i < options.length; i++) {
        await page.keyboard.press('ArrowRight');
        for (const time of [0, 500, 1000]) {
          const result = await page.evaluate((time) => {
            const animations = slide.shadowRoot.getAnimations({ subtree: true });
            for (const animation of animations) {
              animation.pause();
              animation.currentTime = time;
            }
            const incoming = slide.shadowRoot.querySelector(
              '.transition-layer:not(.transition-old)',
            );
            return {
              count: animations.length,
              layers: slide.shadowRoot.querySelectorAll('.transition-layer').length,
              duration: animations[0]?.effect.getTiming().duration,
              mask: incoming && getComputedStyle(incoming).maskImage,
              opacity: incoming && getComputedStyle(incoming).opacity,
            };
          }, time);
          assert.equal(result.count, 1);
          assert.equal(result.layers, 2);
          assert.equal(result.duration, 1000);
          if (options[i].spokes === 0) {
            assert.equal(Number(result.opacity), time / 1000);
            continue;
          }
          assert.match(result.mask, /gradient/);
          if (options[i].spokes > 8) continue;
          const png = await page.locator('#slide').screenshot();
          const spokes = options[i].spokes ?? 4;
          const pixels = await page.evaluate(
            async ({ png, spokes }) => {
              const img = new Image();
              img.src = 'data:image/png;base64,' + png;
              await img.decode();
              const canvas = document.createElement('canvas');
              canvas.width = img.width;
              canvas.height = img.height;
              const ctx = canvas.getContext('2d');
              ctx.drawImage(img, 0, 0);
              const samples = [];
              for (let sector = 0; sector < spokes; sector++)
                for (const fraction of [0.25, 0.75]) {
                  const angle = -Math.PI / 2 + ((sector + fraction) * 2 * Math.PI) / spokes;
                  const radius = Math.min(img.width, img.height) * 0.35;
                  samples.push(
                    Array.from(
                      ctx.getImageData(
                        Math.round(img.width / 2 + radius * Math.cos(angle)),
                        Math.round(img.height / 2 + radius * Math.sin(angle)),
                        1,
                        1,
                      ).data,
                    ).slice(0, 3),
                  );
                }
              return samples;
            },
            { png: png.toString('base64'), spokes },
          );
          const incoming = i % 2 ? [255, 209, 102] : [115, 194, 251];
          const outgoing = i % 2 ? [115, 194, 251] : [255, 209, 102];
          pixels.forEach((pixel, index) =>
            assert.deepEqual(
              pixel,
              time === 1000 || (time === 500 && index % 2 === 0) ? incoming : outgoing,
              `spokes=${spokes} time=${time} sample=${index}`,
            ),
          );
          if (time === 500) await page.screenshot({ path: `/tmp/pptx-wheel-${i}.png` });
        }
        // Leave the first effect unfinished to exercise interruption by the next slide.
        await page.evaluate(() => {
          for (const animation of transitionAnimations) animation.currentTime = 500;
        });
      }
      await page.evaluate(() => {
        for (const animation of transitionAnimations) animation.finish();
      });
      await page.waitForFunction(() => !transitionCleanup);
      assert.equal(await page.locator('.transition-layer').count(), 0);
      assert.match(
        await page.locator('#slide').evaluate((node) => node.shadowRoot.textContent),
        /日本語 \/ English 8/,
      );
      await page.keyboard.press('ArrowLeft');
      assert.equal(await page.locator('.transition-layer').count(), 2);
      await page.emulateMedia({ reducedMotion: 'reduce' });
      await page.waitForFunction(() => !transitionCleanup);
      for (let i = 0; i >= 0; i--) {
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

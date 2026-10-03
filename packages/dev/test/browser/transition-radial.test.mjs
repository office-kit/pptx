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
  'wedge and newsflash play saved transitions and clean up interrupted playback',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-radial-'));
    let preview, browser;
    try {
      const options = [{ effect: 'none' }, { effect: 'wedge' }, { effect: 'newsflash' }];
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
      for (let i = 1; i < options.length; i++) {
        await page.keyboard.press('ArrowRight');
        for (const time of options[i].effect === 'wedge'
          ? [0, 250, 500, 750, 1000]
          : [0, 500, 1000]) {
          const result = await page.evaluate(
            ({ time }) => {
              const animations = slide.shadowRoot.getAnimations({ subtree: true });
              for (const animation of animations) {
                animation.pause();
                animation.currentTime = time;
              }
              const incoming = slide.shadowRoot.querySelector(
                '.transition-layer:not(.transition-old)',
              );
              const bounds = slide.getBoundingClientRect();
              const controls = document.getElementById('presentation-controls');
              controls.style.pointerEvents = 'none';
              const samples = [];
              for (let point = 0; point < 8; point++) {
                const angle = -Math.PI / 2 + ((point + 0.5) * Math.PI) / 4;
                const x = 0.5 + 0.4 * Math.cos(angle);
                const y = 0.5 + 0.4 * Math.sin(angle);
                samples.push(
                  slide.shadowRoot
                    .elementFromPoint(
                      bounds.left + bounds.width * x,
                      bounds.top + bounds.height * y,
                    )
                    ?.closest('.transition-layer') === incoming,
                );
              }
              controls.style.pointerEvents = '';
              return {
                count: animations.length,
                layers: slide.shadowRoot.querySelectorAll('.transition-layer').length,
                duration: animations[0]?.effect.getTiming().duration,
                samples,
                frames: animations[0]?.effect.getKeyframes(),
                transform: incoming && getComputedStyle(incoming).transform,
              };
            },
            { time },
          );
          assert.equal(result.count, 1);
          assert.equal(result.layers, 2, 'interruptions remove previous bands');
          assert.equal(result.duration, 1000);
          if (options[i].effect === 'wedge') {
            assert.deepEqual(
              result.samples,
              time === 0
                ? Array(8).fill(false)
                : time === 1000
                  ? Array(8).fill(true)
                  : time === 250
                    ? [true, false, false, false, false, false, false, true]
                    : time === 750
                      ? [true, true, true, false, false, true, true, true]
                      : [true, true, false, false, false, false, true, true],
            );
          } else {
            const matrix = result.transform
              .match(/matrix\(([^)]+)\)/)[1]
              .split(',')
              .map(Number);
            assert.ok(Math.abs(matrix[0] - time / 1000) < 0.001);
            assert.ok(Math.abs(matrix[1]) < 0.001);
            assert.match(result.frames[0].transform, /rotate\(720deg\)/);
            assert.match(result.frames.at(-1).transform, /rotate\(0deg\)/);
          }
          if (time === 500) await page.screenshot({ path: `/tmp/pptx-radial-${i}.png` });
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
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

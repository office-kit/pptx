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

test('dissolve reveals scattered cells at all transition speeds', { timeout: 60000 }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-dissolve-'));
  let preview, browser;
  try {
    const options = [
      { effect: 'none' },
      ...['slow', 'med', 'fast'].map((speed) => ({ effect: 'dissolve', speed })),
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
    getSlides(deck).forEach((slide, i) => setSlideTransition(slide, options[i]));
    getSlides(deck).forEach((slide, i) => setSlideBackground(slide, i % 2 ? '#FFD166' : '#73C2FB'));
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
    await page.waitForFunction(() => state.slides.length === 4);
    await page.getByRole('button', { name: 'Present', exact: true }).click();
    for (let i = 1; i < options.length; i++) {
      await page.keyboard.press('ArrowRight');
      const duration = { slow: 1000, med: 600, fast: 300 }[options[i].speed];
      let previous = [];
      for (const progress of [0, 0.25, 0.5, 0.75, 1]) {
        const result = await page.evaluate(
          ({ progress, duration }) => {
            const animations = slide.shadowRoot.getAnimations({ subtree: true });
            for (const animation of animations) {
              animation.pause();
              animation.currentTime = duration * progress;
            }
            const incoming = slide.shadowRoot.querySelector(
              '.transition-layer:not(.transition-old)',
            );
            const bounds = slide.getBoundingClientRect();
            const controls = document.getElementById('presentation-controls');
            controls.style.pointerEvents = 'none';
            const samples = [];
            for (let row = 0; row < 18; row++)
              for (let col = 0; col < 32; col++) {
                const x = (col + 0.5) / 32;
                const y = (row + 0.5) / 18;
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
            };
          },
          { progress, duration },
        );
        assert.equal(result.count, 1);
        assert.equal(result.layers, 2, 'interruptions remove previous cells');
        assert.equal(result.duration, duration);
        assert.equal(result.samples.filter(Boolean).length, 576 * progress);
        previous.forEach((visible, index) => {
          if (visible) assert.equal(result.samples[index], true, 'revealed cells stay visible');
        });
        previous = result.samples;
        if (progress === 0.5) await page.screenshot({ path: `/tmp/pptx-dissolve-${i}.png` });
      }
      // Leave the first effect unfinished to exercise interruption by the next slide.
      await page.evaluate(() => {
        for (const animation of transitionAnimations)
          animation.currentTime = Number(animation.effect.getTiming().duration) / 2;
      });
    }
    await page.evaluate(() => {
      for (const animation of transitionAnimations) animation.finish();
    });
    await page.waitForFunction(() => !transitionCleanup);
    assert.equal(await page.locator('.transition-layer').count(), 0);
    assert.match(
      await page.locator('#slide').evaluate((node) => node.shadowRoot.textContent),
      /日本語 \/ English 3/,
    );
    await page.keyboard.press('ArrowLeft');
    assert.equal(await page.locator('.transition-layer').count(), 2);
    await page.emulateMedia({ reducedMotion: 'reduce' });
    await page.waitForFunction(() => !transitionCleanup);
    for (let i = 1; i >= 0; i--) {
      await page.keyboard.press('ArrowLeft');
      assert.equal(await page.locator('.transition-layer').count(), 0);
    }
    assert.deepEqual(errors, []);
  } finally {
    await browser?.close();
    await preview?.close();
    await rm(dir, { recursive: true, force: true });
  }
});

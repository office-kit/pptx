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
  'split and shape transitions play from saved slides and clean up on interruption',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-shape-transitions-'));
    let preview, browser;
    try {
      const options = [
        { effect: 'none' },
        { effect: 'split', orientation: 'horz', direction: 'out' },
        { effect: 'split', orientation: 'horz', direction: 'in' },
        { effect: 'split', orientation: 'vert', direction: 'out' },
        { effect: 'split', orientation: 'vert', direction: 'in' },
        { effect: 'circle' },
        { effect: 'diamond' },
        { effect: 'plus' },
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
      await page.waitForFunction(() => state.slides.length === 8);
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      for (let i = 1; i < options.length; i++) {
        await page.keyboard.press('ArrowRight');
        const result = await page.evaluate(() => {
          const animations = slide.shadowRoot.getAnimations({ subtree: true });
          for (const animation of animations) {
            animation.pause();
            animation.currentTime = 500;
          }
          const animation = animations[0];
          return {
            count: animations.length,
            layers: slide.shadowRoot.querySelectorAll('.transition-layer').length,
            frames: animation?.effect.getKeyframes(),
            clip: animation && getComputedStyle(animation.effect.target).clipPath,
            outgoing: animation?.effect.target.classList.contains('transition-old'),
            duration: animation?.effect.getTiming().duration,
          };
        });
        assert.equal(result.count, 1, options[i].effect);
        assert.equal(
          result.layers,
          2,
          'interrupted animations retain only current and previous layers',
        );
        assert.equal(result.duration, 1000);
        assert.notEqual(result.clip, 'none');
        assert.notEqual(result.clip, result.frames[0].clipPath);
        assert.notEqual(result.clip, result.frames[1].clipPath);
        assert.equal(result.outgoing, options[i].direction === 'in');
        if (i <= 4) {
          const collapsed = i <= 2 ? 'inset(50% 0px)' : 'inset(0px 50%)';
          assert.equal(result.frames[options[i].direction === 'in' ? 1 : 0].clipPath, collapsed);
        }
      }
      await page.screenshot({ path: '/tmp/pptx-plus-transition.png' });
      await page.evaluate(() => {
        for (const animation of transitionAnimations) animation.finish();
      });
      await page.waitForFunction(() => !transitionCleanup);
      assert.equal(await page.locator('.transition-layer').count(), 0);
      assert.match(
        await page.locator('#slide').evaluate((node) => node.shadowRoot.textContent),
        /日本語 \/ English 7/,
      );
      await page.keyboard.press('ArrowLeft');
      assert.equal(await page.locator('.transition-layer').count(), 2);
      await page.emulateMedia({ reducedMotion: 'reduce' });
      await page.waitForFunction(() => !transitionCleanup);
      for (let i = 5; i >= 0; i--) {
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

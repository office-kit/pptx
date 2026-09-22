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
  'presentation renders transitions, cancels interrupted layers, and respects reduced motion',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-transitions-'));
    let preview, browser;
    try {
      const effects = ['none', 'fade', 'push', 'wipe', 'cover', 'pull', 'zoom', 'fade'];
      const deck = await compile(
        Presentation({
          children: effects.map((effect, i) =>
            Slide({
              children: Text({ x: 1, y: 1, width: 8, height: 2, children: `${i + 1}: ${effect}` }),
            }),
          ),
        }),
      );
      getSlides(deck).forEach((slide, i) =>
        setSlideTransition(slide, {
          effect: effects[i],
          speed: 'slow',
          ...(i === 2
            ? { direction: 'r' }
            : i === 3
              ? { direction: 'd' }
              : i === 6
                ? { direction: 'out' }
                : {}),
          ...(i === 7 ? { thruBlack: true, advanceAfterMs: 50 } : {}),
        }),
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
      await page.keyboard.press('ArrowRight');
      assert.equal(
        await page.locator('.transition-layer').count(),
        0,
        'ordinary preview switches immediately',
      );
      await page.keyboard.press('Home');
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      for (let i = 1; i <= 7; i++) {
        await page.keyboard.press('ArrowRight');
        const animation = await page.evaluate(() => {
          const animations = slide.shadowRoot.getAnimations({ subtree: true });
          for (const animation of animations) {
            animation.pause();
            animation.currentTime = 500;
          }
          return {
            count: animations.length,
            duration: animations[0]?.effect.getTiming().duration,
            frames: animations.map((a) => a.effect.getKeyframes()),
            oldHidden: slide.shadowRoot
              .querySelector('.transition-old')
              ?.getAttribute('aria-hidden'),
            oldInert: slide.shadowRoot.querySelector('.transition-old')?.inert,
            incomingOpacity: getComputedStyle(
              slide.shadowRoot.querySelector('.transition-layer:not(.transition-old)'),
            ).opacity,
            current: index,
          };
        });
        assert.equal(animation.current, i);
        assert.equal(animation.duration, 1000);
        assert.equal(animation.oldHidden, 'true');
        assert.equal(animation.oldInert, true);
        assert.equal(animation.count, [2, 7].includes(i) ? 2 : 1);
        if (i === 1)
          assert.ok(Number(animation.incomingOpacity) > 0 && Number(animation.incomingOpacity) < 1);
        if (i === 2) assert.equal(animation.frames[0][0].transform, 'translate(-100%, 0%)');
        if (i === 3) assert.equal(animation.frames[0][0].clipPath, 'inset(0px 0px 100%)');
        if (i === 6) assert.equal(animation.frames[0][0].transform, 'scale(2)');
        if (i === 7) assert.equal(animation.incomingOpacity, '0');
        assert.equal(
          await page.locator('.transition-layer').count(),
          2,
          'interrupted effects leave only two layers',
        );
      }
      // A server refresh with unchanged slides must not restart the animation.
      const preserved = await page.evaluate(async () => {
        const animation = transitionAnimations[0];
        await refresh();
        return animation === transitionAnimations[0] && animation.currentTime === 500;
      });
      assert.equal(preserved, true);
      await page.emulateMedia({ reducedMotion: 'reduce' });
      await page.waitForFunction(() => !transitionCleanup);
      assert.equal(await page.locator('.transition-layer').count(), 0);
      assert.equal(await page.locator('#slide').evaluate((node) => node.style.background), '');
      await page.keyboard.press('ArrowLeft');
      assert.equal(await page.locator('.transition-layer').count(), 0);
      await page.emulateMedia({ reducedMotion: 'no-preference' });
      await page.keyboard.press('ArrowLeft');
      assert.equal(await page.locator('.transition-layer').count(), 2);
      const pausedIndex = await page.evaluate(() => {
        state.transitions[index].advanceAfterMs = 50;
        for (const animation of transitionAnimations) animation.pause();
        return index;
      });
      await page.evaluate(() => new Promise((resolve) => setTimeout(resolve, 100)));
      assert.equal(
        await page.evaluate(() => index),
        pausedIndex,
        'advance waits for the transition',
      );
      await page.evaluate(() => {
        for (const animation of transitionAnimations) animation.finish();
      });
      await page.waitForFunction((previous) => index === previous + 1, pausedIndex);
      await page.waitForFunction(() => !transitionCleanup);
      assert.equal(
        await page.locator('.transition-layer').count(),
        0,
        'finished animation releases both layers',
      );
      assert.match(
        await page.locator('#slide').evaluate((node) => node.shadowRoot.textContent),
        /7: zoom/,
      );
      await page.keyboard.press('ArrowLeft');
      assert.equal(await page.locator('.transition-layer').count(), 2);
      await page.getByRole('button', { name: 'Exit · Esc', exact: true }).click();
      assert.equal(await page.locator('.transition-layer').count(), 0);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

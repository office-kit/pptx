import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test(
  'configured slideshow ranges, looping and timers control playback',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-link-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={8} height={1}>Link target</Text></Slide><Slide><Text x={1} y={1} width={8} height={1}>Return</Text></Slide></Presentation>`,
    );
    const proc = spawn(process.execPath, [
      fileURLToPath(new URL('../../dist/cli.mjs', import.meta.url)),
      'dev',
      file,
      '--port',
      '0',
    ]);
    let browser;
    try {
      const url = await new Promise((resolve, reject) => {
        let output = '';
        proc.stdout.on('data', (data) => {
          output += data;
          const match = output.match(/Preview: (http:\/\/\S+)/);
          if (match) resolve(match[1]);
        });
        proc.on('error', reject);
        proc.on('exit', (code) => reject(new Error('Server exited ' + code)));
      });
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(url);
      await page.waitForFunction(() => window.office?.getState().editor?.slides.length === 2);
      await page.locator('#tab-transitions').click();
      const soundSelect = page.getByRole('combobox', { name: 'Transition sound', exact: true });
      const soundLoop = page.locator('#transition-sound-loop');
      assert.equal(await soundLoop.isDisabled(), true);
      const wav = Buffer.alloc(46);
      wav.write('RIFF');
      wav.writeUInt32LE(38, 4);
      wav.write('WAVEfmt ', 8);
      wav.writeUInt32LE(16, 16);
      wav.writeUInt16LE(1, 20);
      wav.writeUInt16LE(1, 22);
      wav.writeUInt32LE(8000, 24);
      wav.writeUInt32LE(16000, 28);
      wav.writeUInt16LE(2, 32);
      wav.writeUInt16LE(16, 34);
      wav.write('data', 36);
      wav.writeUInt32LE(2, 40);
      const chooserPromise = page.waitForEvent('filechooser');
      await soundSelect.selectOption('other');
      await (
        await chooserPromise
      ).setFiles({ name: 'Chime.wav', mimeType: 'audio/wav', buffer: wav });
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].transitionSound?.name === 'Chime.wav',
      );
      await soundLoop.check();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].transitionSound?.loop === true,
      );
      await page.reload();
      await page.waitForFunction(() => window.office?.getState().editor?.slides.length === 2);
      await page.locator('#tab-transitions').click();
      assert.equal(await soundSelect.inputValue(), 'embedded');
      assert.equal(await soundLoop.isChecked(), true);
      await page.evaluate(() => {
        window.soundCalls = [];
        window.Audio = class extends EventTarget {
          constructor(src) {
            super();
            this.src = src;
            this.paused = false;
            window.soundCalls.push(this);
          }
          play() {
            return Promise.resolve();
          }
          pause() {
            this.paused = true;
          }
          removeAttribute() {
            this.src = '';
          }
          load() {}
        };
      });
      await page.locator('[data-edit="transition-preview"]').click();
      assert.equal(
        await page.evaluate(() => window.soundCalls.length === 1 && window.soundCalls[0].loop),
        true,
      );
      await page.evaluate(() => window.office.selectSlide(1));
      assert.equal(await page.evaluate(() => window.soundCalls[0].paused), true);
      await page.evaluate(() => window.office.selectSlide(0));
      await page.locator('#tab-show').click();
      await page.locator('#present-first').click();
      await page.waitForFunction(() => window.office.getState().presenting);
      assert.equal(
        await page.evaluate(() => window.soundCalls.length === 2 && !window.soundCalls[1].paused),
        true,
      );
      const slideBox = await page.locator('#slide').boundingBox();
      await page.mouse.move(slideBox.x + slideBox.width / 2, slideBox.y + slideBox.height / 2);
      await page.keyboard.press('Meta+l');
      await page.locator('#show-laser-pointer').waitFor({ state: 'visible' });
      assert.equal(await page.locator('#slide').getAttribute('data-show-cursor'), 'none');
      assert.equal(
        await page.locator('#slide svg').evaluate((node) => getComputedStyle(node).cursor),
        'none',
      );
      await page.locator('#present-pointer').click();
      await page.getByRole('menuitem', { name: 'Laser Color', exact: true }).hover();
      await page.getByRole('menuitemcheckbox', { name: 'Green', exact: true }).click();
      await page.mouse.move(slideBox.x + slideBox.width / 2 + 20, slideBox.y + slideBox.height / 2);
      assert.equal(await page.locator('#show-laser-pointer').getAttribute('data-color'), 'green');
      await page.keyboard.press('Meta+i');
      await page.locator('#show-laser-pointer').waitFor({ state: 'hidden' });
      await page.mouse.move(slideBox.x + slideBox.width / 2 + 30, slideBox.y + slideBox.height / 2);
      assert.equal(await page.locator('#slide').getAttribute('data-show-cursor'), 'none');
      await page.keyboard.press('Meta+a');
      assert.equal(await page.locator('#slide').getAttribute('data-show-cursor'), 'arrow');
      await page.keyboard.press('Meta+u');
      assert.equal(await page.locator('#slide').getAttribute('data-show-cursor'), null);
      await page.waitForFunction(
        () => document.getElementById('slide').dataset.showCursor === 'none',
      );
      await page.mouse.move(slideBox.x + slideBox.width / 2 + 40, slideBox.y + slideBox.height / 2);
      assert.equal(await page.locator('#slide').getAttribute('data-show-cursor'), null);
      await page.keyboard.press('Meta+l');
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => window.office.getState().index === 1);
      assert.equal(await page.evaluate(() => window.soundCalls[1].paused), false);
      await page.locator('#show-laser-pointer').waitFor({ state: 'visible' });
      await page.keyboard.press('Escape');
      assert.equal(await page.evaluate(() => window.office.getState().presenting), true);
      await page.locator('#show-laser-pointer').waitFor({ state: 'hidden' });
      assert.equal(await page.evaluate(() => window.soundCalls[1].paused), false);
      await page.keyboard.press('Escape');
      await page.waitForFunction(() => !window.office.getState().presenting);
      assert.equal(await page.evaluate(() => window.soundCalls[1].paused), true);
      await page.locator('#show-laser-pointer').waitFor({ state: 'hidden' });
      assert.equal(await page.locator('#slide').getAttribute('data-show-cursor'), null);
      await page.evaluate(() => window.office.selectSlide(0));
      await page.locator('#tab-transitions').click();
      await soundSelect.selectOption('stop');
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].transitionSound?.kind === 'stop',
      );
      assert.equal(await soundLoop.isDisabled(), true);
      await soundSelect.selectOption('none');
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].transitionSound?.kind === 'none',
      );
      const afterTime = page.getByRole('textbox', { name: 'Advance slide after', exact: true });
      assert.equal(await afterTime.isDisabled(), true);
      await page.locator('#advance-after').check();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].transition?.advanceAfterMs === 0,
      );
      await afterTime.fill('00:01.25');
      await afterTime.press('Tab');
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].transition?.advanceAfterMs === 1250,
      );
      await page.locator('#advance-click').uncheck();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].transition?.advanceOnClick === false,
      );
      await page.reload();
      await page.waitForFunction(() => window.office?.getState().editor?.slides.length === 2);
      await page.locator('#tab-transitions').click();
      assert.equal(await afterTime.inputValue(), '00:01.25');
      assert.equal(await page.locator('#advance-click').isChecked(), false);
      await page.locator('[data-edit="transition-apply-all"]').click();
      await page.waitForFunction(() =>
        window.office
          .getState()
          .editor.slides.every(
            (slide) =>
              slide.transition?.advanceAfterMs === 1250 &&
              slide.transition?.advanceOnClick === false,
          ),
      );
      await page.locator('[data-edit="undo"]').first().click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[1].transition === null,
      );
      assert.equal(await page.locator('#advance-click').isChecked(), false);
      const timingBounds = await afterTime.boundingBox();
      const ribbonBounds = await page.locator('#ribbon-transitions').boundingBox();
      assert.ok(
        timingBounds.y >= ribbonBounds.y &&
          timingBounds.y + timingBounds.height <= ribbonBounds.y + ribbonBounds.height,
      );
      await page.locator('[data-edit="undo"]').first().click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].transition?.advanceOnClick === true,
      );
      await afterTime.fill('00:99.00');
      await afterTime.press('Tab');
      assert.equal(await afterTime.evaluate((input) => input.checkValidity()), false);
      assert.equal(
        await page.evaluate(
          () => window.office.getState().editor.slides[0].transition.advanceAfterMs,
        ),
        1250,
      );
      await afterTime.fill('00:02.00');
      await afterTime.press('Tab');
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].transition?.advanceAfterMs === 2000,
      );
      await page.locator('#advance-after').uncheck();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].transition?.advanceAfterMs === undefined,
      );
      await page.locator('[data-effect="push"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].transition?.effect === 'push',
      );
      assert.equal(await page.locator('[data-effect="push"]').getAttribute('aria-pressed'), 'true');
      const duration = page.getByRole('spinbutton', { name: 'Transition duration', exact: true });
      await duration.fill('1.75');
      await duration.press('Tab');
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].transition?.durationMs === 1750,
      );
      await duration.fill('-1');
      await duration.press('Tab');
      assert.equal(await duration.evaluate((input) => input.checkValidity()), false);
      assert.equal(
        await page.evaluate(() => window.office.getState().editor.slides[0].transition.durationMs),
        1750,
      );
      await duration.fill('1.75');
      await duration.press('Tab');
      await page.locator('#transition-options').click();
      assert.equal(
        await page
          .getByRole('menuitemcheckbox', { name: 'From Right', exact: true })
          .getAttribute('aria-checked'),
        'true',
      );
      await page.getByRole('menuitemcheckbox', { name: 'From Top', exact: true }).click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].transition?.direction === 'd',
      );
      await page.reload();
      await page.waitForFunction(() => window.office?.getState().editor?.slides.length === 2);
      await page.locator('#tab-transitions').click();
      assert.equal(await page.locator('[data-effect="push"]').getAttribute('aria-pressed'), 'true');
      assert.equal(await duration.inputValue(), '1.75');
      await page.locator('#transition-options').click();
      assert.equal(
        await page
          .getByRole('menuitemcheckbox', { name: 'From Top', exact: true })
          .getAttribute('aria-checked'),
        'true',
      );
      await page.keyboard.press('Escape');
      await page.locator('[data-edit="transition-preview"]').click();
      const playback = page.locator('[data-transition-playback="push"]');
      await playback.waitFor();
      assert.equal(
        await playback.evaluate(
          (node) => node.children[1].getAnimations()[0].effect.getTiming().duration,
        ),
        1750,
      );
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      await playback.evaluate((node) =>
        [...node.children]
          .flatMap((layer) => layer.getAnimations())
          .forEach((animation) => animation.finish()),
      );
      await playback.waitFor({ state: 'detached' });
      await page.getByRole('button', { name: 'More transition effects', exact: true }).click();
      await page.getByRole('menuitemcheckbox', { name: 'Reverse Wheel', exact: true }).click();
      await page.locator('#transition-options').click();
      await page.getByRole('menuitemcheckbox', { name: '3 Spokes', exact: true }).click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].transition?.spokes === 3,
      );
      await page.reload();
      await page.waitForFunction(
        () => window.office?.getState().editor?.slides[0].transition?.spokes === 3,
      );
      await page.locator('#tab-transitions').click();
      assert.equal(await duration.inputValue(), '1.75');
      await page.locator('#transition-options').click();
      assert.equal(
        await page
          .getByRole('menuitemcheckbox', { name: '3 Spokes', exact: true })
          .getAttribute('aria-checked'),
        'true',
      );
      await page.keyboard.press('Escape');
      await page.locator('[data-edit="transition-preview"]').click();
      await page.locator('[data-transition-playback="wheelReverse"]').waitFor();
      await page.getByRole('button', { name: 'More transition effects', exact: true }).click();
      await page.getByRole('menuitemcheckbox', { name: 'Dissolve', exact: true }).click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].transition?.effect === 'dissolve',
      );
      assert.equal(await page.locator('#transition-options').isDisabled(), true);
      await page.locator('[data-effect="none"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].transition?.effect === 'none',
      );
      await page.locator('[data-edit="undo"]').first().click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].transition?.effect === 'dissolve',
      );
      assert.equal(await duration.inputValue(), '1.75');
      // The remaining cases isolate show navigation/timers from transition duration.
      await page.locator('[data-effect="none"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].transition?.effect === 'none',
      );
      const openSettings = async () => {
        await page.locator('#tab-show').click();
        await page.locator('#show-properties').click();
      };
      const dialog = page.getByRole('dialog', { name: 'Set Up Show', exact: true });
      await openSettings();
      assert.equal(await dialog.getByLabel('Show scrollbar', { exact: true }).isDisabled(), true);
      assert.equal(await dialog.getByLabel('From slide').isDisabled(), true);
      assert.equal(await dialog.getByRole('radio', { name: 'Custom show:' }).isDisabled(), true);
      await dialog.getByRole('radio', { name: 'From:', exact: true }).check();
      await dialog.getByLabel('From slide').fill('2');
      await dialog.getByLabel('To slide').fill('1');
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      assert.equal(await dialog.isVisible(), true);
      await dialog.getByLabel('To slide').fill('2');
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.equal(
        await page.evaluate(() => window.office.getState().editor.showProperties.slides.kind),
        'all',
      );
      await openSettings();
      await dialog.getByRole('radio', { name: 'From:', exact: true }).check();
      await dialog.getByLabel('From slide').fill('2');
      await dialog.getByLabel('To slide').fill('2');
      await dialog
        .getByRole('radio', { name: 'Browsed by an individual (window)', exact: true })
        .check();
      await dialog.getByLabel('Show scrollbar', { exact: true }).uncheck();
      await dialog.getByRole('radio', { name: 'Manually', exact: true }).check();
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await dialog.waitFor({ state: 'detached' });
      await openSettings();
      assert.equal(await dialog.getByLabel('From slide').inputValue(), '2');
      assert.equal(await dialog.getByLabel('Show scrollbar', { exact: true }).isChecked(), false);
      assert.equal(
        await dialog.getByRole('radio', { name: 'Manually', exact: true }).isChecked(),
        true,
      );
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      await page.locator('[data-edit="undo"]').first().click();
      await page.waitForFunction(
        () => window.office.getState().editor.showProperties.slides.kind === 'all',
      );
      await openSettings();
      await dialog
        .getByRole('radio', { name: 'Browsed at a kiosk (full screen)', exact: true })
        .check();
      assert.equal(await dialog.getByLabel("Loop continuously until 'Esc'").isChecked(), true);
      assert.equal(await dialog.getByLabel("Loop continuously until 'Esc'").isDisabled(), true);
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      const configure = async (patch) => {
        await page.evaluate(async (patch) => {
          const state = window.office.getState();
          const response = await fetch('/edit', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              revision: state.revision,
              command: {
                type: 'show-properties',
                slide: 0,
                showProperties: { ...state.editor.showProperties, ...patch },
              },
            }),
          });
          if (!response.ok) throw new Error(await response.text());
          await window.office.refresh();
        }, patch);
      };
      const start = async () => {
        await page.locator('#tab-show').click();
        await page.locator('#present-first').click();
        await page.waitForFunction(() => window.office.getState().presenting);
      };
      const stop = async () => {
        await page.keyboard.press('Escape');
        await page.waitForFunction(() => !window.office.getState().presenting);
      };
      await configure({
        mode: { kind: 'browse', showScrollbar: true },
        slides: { kind: 'range', start: 2, end: 2 },
        loop: false,
      });
      await start();
      assert.equal(await page.evaluate(() => window.office.getState().index), 1);
      assert.equal(await page.evaluate(() => !!document.fullscreenElement), false);
      await page.keyboard.press('ArrowRight');
      await page.locator('#show-screen.end').waitFor({ state: 'visible' });
      await stop();
      await configure({ slides: { kind: 'all' }, loop: true });
      await start();
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => window.office.getState().index === 1);
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => window.office.getState().index === 0);
      assert.equal(await page.locator('#show-screen').isVisible(), false);
      await stop();
      await page.evaluate(async () => {
        const state = window.office.getState();
        const response = await fetch('/edit', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            revision: state.revision,
            command: {
              type: 'custom-shows',
              slide: 0,
              customShows: [
                {
                  id: 4,
                  name: 'Configured',
                  slides: [
                    state.editor.slides[1].key,
                    state.editor.slides[0].key,
                    state.editor.slides[1].key,
                  ],
                },
              ],
            },
          }),
        });
        if (!response.ok) throw new Error(await response.text());
        await window.office.refresh();
      });
      await configure({ slides: { kind: 'customShow', id: 4 }, loop: false });
      await page.reload();
      await page.waitForFunction(
        () => window.office?.getState().editor?.showProperties.slides.kind === 'customShow',
      );
      await start();
      assert.equal(await page.evaluate(() => window.office.getState().index), 1);
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => window.office.getState().index === 0);
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => window.office.getState().index === 1);
      await page.keyboard.press('ArrowRight');
      await page.locator('#show-screen.end').waitFor({ state: 'visible' });
      await stop();
      // Scroll positions belong to the show sequence, including repeated slides.
      await start();
      const scrollbar = page.locator('#browse-scrollbar');
      assert.equal(await scrollbar.isVisible(), true);
      await scrollbar.hover();
      await page.mouse.wheel(0, await scrollbar.evaluate((bar) => bar.clientHeight));
      await page.waitForFunction(() => window.office.getState().index === 0);
      await scrollbar.evaluate((bar) => {
        bar.scrollTop = bar.clientHeight * 2;
      });
      await page.waitForFunction(() => window.office.getState().index === 1);
      await page.keyboard.press('ArrowLeft');
      await page.waitForFunction(() => window.office.getState().index === 0);
      assert.equal(await scrollbar.evaluate((bar) => bar.scrollTop / bar.clientHeight), 1);
      await stop();
      assert.equal(await scrollbar.isVisible(), false);
      await configure({ mode: { kind: 'browse', showScrollbar: false } });
      await start();
      assert.equal(await scrollbar.isVisible(), false);
      await stop();
      await configure({ slides: { kind: 'all' } });
      await openSettings();
      await dialog
        .getByRole('radio', { name: 'Browsed at a kiosk (full screen)', exact: true })
        .check();
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await dialog.waitFor({ state: 'detached' });
      await start();
      assert.equal(await page.locator('#presentation-controls').isVisible(), false);
      await page.keyboard.press('ArrowRight');
      await page.keyboard.press('n');
      await page.locator('#stage').click({ position: { x: 10, y: 10 } });
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      await stop();
      await configure({ mode: { kind: 'browse', showScrollbar: true } });
      // Core tests verify timing extraction; inject model timing to isolate the playback clock.
      const timing = async () =>
        page.evaluate(() => {
          for (const slide of window.office.getState().editor.slides)
            slide.transition = { effect: 'none', advanceAfterMs: 250, advanceOnClick: false };
        });
      await configure({ loop: false, useTimings: true });
      await timing();
      await start();
      await page.waitForFunction(() => window.office.getState().index === 1);
      await page.locator('#show-screen.end').waitFor({ state: 'visible' });
      await stop();
      await configure({ useTimings: false });
      await timing();
      await start();
      await page.waitForTimeout(600);
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      await page.locator('#stage').click({ position: { x: 10, y: 10 } });
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => window.office.getState().index === 1);
      await stop();
      await configure({ useTimings: true });
      await timing();
      await start();
      await stop();
      const index = await page.evaluate(() => window.office.getState().index);
      await page.waitForTimeout(600);
      assert.equal(await page.evaluate(() => window.office.getState().index), index);
      // The kiosk deadline is measured from show start, not the latest input.
      await configure({ mode: { kind: 'kiosk', restart: 1000 }, useTimings: false });
      await page.clock.install();
      await page.clock.pauseAt(new Date());
      await start();
      await page.evaluate(() => window.office.selectSlide(1));
      await page.clock.runFor(900);
      await page.locator('#stage').click({ position: { x: 10, y: 10 } });
      assert.equal(await page.evaluate(() => window.office.getState().index), 1);
      await page.clock.runFor(101);
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      await stop();
      await page.evaluate(() => window.office.selectSlide(1));
      await page.clock.runFor(1100);
      assert.equal(await page.evaluate(() => window.office.getState().index), 1);
      await configure({ slides: { kind: 'customShow', id: 4 } });
      await start();
      await page.evaluate(() => window.office.selectSlide(0));
      await page.clock.runFor(1001);
      assert.equal(await page.evaluate(() => window.office.getState().index), 1);
      await stop();
      // A nested show-and-return must not replace the original restart destination.
      await configure({ slides: { kind: 'all' } });
      await start();
      await page.clock.runFor(500);
      await page.evaluate(() => window.office.startCustomShow(4, true));
      assert.equal(await page.evaluate(() => window.office.getState().index), 1);
      await page.clock.runFor(501);
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      await stop();
      // Automatic looping resets the kiosk deadline at the end of the sequence.
      await configure({ mode: { kind: 'kiosk', restart: 700 }, useTimings: true });
      await timing();
      await start();
      await page.clock.runFor(800);
      assert.equal(await page.evaluate(() => window.office.getState().index), 1);
      await stop();
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      proc.kill('SIGTERM');
      await new Promise((resolve) =>
        proc.exitCode !== null ? resolve() : proc.once('exit', resolve),
      );
      await rm(dir, { recursive: true, force: true });
    }
  },
);

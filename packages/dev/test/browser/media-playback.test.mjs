import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { compile, Media, Presentation, Slide, Text } from '@office-kit/pptx-dsl';
import {
  getSlideShapes,
  getSlides,
  savePresentation,
  setShapeMediaPlayback,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const wav = (durationMs = 5000) => {
  const sampleRate = 8000;
  const samples = Math.floor((sampleRate * durationMs) / 1000);
  const bytes = new Uint8Array(44 + samples * 2);
  const view = new DataView(bytes.buffer);
  const text = (offset, value) =>
    [...value].forEach((char, index) => view.setUint8(offset + index, char.charCodeAt(0)));
  text(0, 'RIFF');
  view.setUint32(4, 36 + samples * 2, true);
  text(8, 'WAVE');
  text(12, 'fmt ');
  view.setUint32(16, 16, true);
  view.setUint16(20, 1, true);
  view.setUint16(22, 1, true);
  view.setUint32(24, sampleRate, true);
  view.setUint32(28, sampleRate * 2, true);
  view.setUint16(32, 2, true);
  view.setUint16(34, 16, true);
  text(36, 'data');
  view.setUint32(40, samples * 2, true);
  for (let index = 0; index < samples; index++) {
    view.setInt16(
      44 + index * 2,
      Math.round(Math.sin((index * 2 * Math.PI * 440) / sampleRate) * 1000),
      true,
    );
  }
  return bytes;
};

test(
  'embedded audio plays in presentation mode and resets when revisiting its slide',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-media-playback-'));
    let preview;
    let browser;
    try {
      const deck = await compile(
        Presentation({
          children: [
            Slide({
              children: Media({ kind: 'audio', data: wav(), x: 1, y: 1, width: 3, height: 1 }),
            }),
            Slide({ children: Text({ x: 1, y: 1, width: 5, height: 1, children: 'Second' }) }),
          ],
        }),
      );
      const [audio] = getSlideShapes(getSlides(deck)[0]);
      setShapeMediaPlayback(audio, {
        autoplay: true,
        delayMs: 1200,
        loop: true,
        volume: 0.25,
        muted: true,
      });
      await writeFile(join(dir, 'source.pptx'), await savePresentation(deck));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(join(dir, 'source.pptx'))})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
      await page.addInitScript(() => {
        const originalPlay = HTMLMediaElement.prototype.play;
        HTMLMediaElement.prototype.play = function () {
          this.dataset.playCalls = String(Number(this.dataset.playCalls ?? 0) + 1);
          return originalPlay.call(this);
        };
      });
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.waitForFunction(() => state.slides.length === 2);
      const persistedMedia = await page
        .evaluate(async () => (await fetch('/state')).json())
        .then((current) => current.media[0]);
      assert.equal(persistedMedia.playback.delayMs, 1200);
      assert.equal(persistedMedia.playback.loop, true);
      assert.equal(persistedMedia.playback.volume, 0.25);
      assert.equal(persistedMedia.playback.muted, true);
      await page.getByRole('button', { name: 'Present', exact: true }).click();

      const audioElement = page.locator('foreignObject[data-pptx-media] audio');
      await audioElement.waitFor({ state: 'attached' });
      const oldAudio = await audioElement.elementHandle();
      assert.ok(oldAudio);
      await page.waitForTimeout(300);
      assert.equal(await oldAudio.evaluate((element) => element.paused), true);
      assert.equal(await oldAudio.evaluate((element) => element.currentTime), 0);
      await page.waitForFunction(() => {
        const root = document.querySelector('#slide')?.shadowRoot;
        const element = root?.querySelector('foreignObject[data-pptx-media] audio');
        return element?.readyState >= 2;
      });
      const initialTime = await audioElement.evaluate((element) => element.currentTime);
      assert.deepEqual(
        await audioElement.evaluate((element) => ({
          loop: element.loop,
          volume: element.volume,
          muted: element.muted,
        })),
        { loop: true, volume: 0.25, muted: true },
      );
      await page.waitForFunction((initial) => {
        const root = document.querySelector('#slide')?.shadowRoot;
        const element = root?.querySelector('foreignObject[data-pptx-media] audio');
        return element !== null && element.currentTime > initial + 0.05;
      }, initialTime);
      await page.evaluate(() => refresh());
      assert.equal(
        await oldAudio.evaluate((element) => element.isConnected),
        true,
        'unchanged refresh preserves the running media element',
      );

      await page.locator('#stage').focus();
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => index === 1);
      await page.waitForFunction((element) => element.paused, oldAudio);

      await page.locator('#stage').focus();
      await page.keyboard.press('ArrowLeft');
      await page.waitForFunction(() => index === 0);
      const revisited = page.locator('foreignObject[data-pptx-media] audio');
      await revisited.waitFor({ state: 'attached' });
      await page.waitForFunction(() => {
        const root = document.querySelector('#slide')?.shadowRoot;
        const element = root?.querySelector('foreignObject[data-pptx-media] audio');
        return element?.readyState >= 2;
      });
      assert.ok((await revisited.evaluate((element) => element.currentTime)) < 0.5);
      const pendingAudio = await revisited.elementHandle();
      assert.ok(pendingAudio);
      assert.equal(await pendingAudio.evaluate((element) => element.dataset.playCalls), undefined);
      await revisited.focus();
      await page.keyboard.press('Escape');
      await revisited.waitFor({ state: 'detached' });
      assert.equal(await page.evaluate(() => presenting), false);
      await page.waitForTimeout(1300);
      assert.equal(await pendingAudio.evaluate((element) => element.dataset.playCalls), undefined);
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      const manualAudio = page.locator('foreignObject[data-pptx-media] audio');
      await manualAudio.waitFor({ state: 'attached' });
      await manualAudio.evaluate(async (element) => {
        await element.play();
        element.pause();
      });
      await page.waitForTimeout(1300);
      assert.equal(await manualAudio.evaluate((element) => element.paused), true);
      assert.equal(await manualAudio.evaluate((element) => element.dataset.playCalls), '1');
      assert.equal(await oldAudio.evaluate((element) => element.paused), true);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

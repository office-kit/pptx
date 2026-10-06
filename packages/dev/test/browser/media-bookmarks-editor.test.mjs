import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  addBlankSlide,
  addSlideMedia,
  createPresentation,
  getShapeMediaPlayback,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeMediaPlayback,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const wav = (durationMs = 5033) => {
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
  for (let index = 0; index < samples; index++)
    view.setInt16(
      44 + index * 2,
      Math.round(Math.sin((index * 2 * Math.PI * 440) / sampleRate) * 12000),
      true,
    );
  return bytes;
};

test(
  'media bookmarks add, persist, undo, and remove from the playback ribbon',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-media-bookmarks-'));
    let browser;
    let preview;
    try {
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      addSlideMedia(slide, {
        kind: 'audio',
        data: wav(10000),
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(2),
      });
      addSlideMedia(slide, {
        kind: 'audio',
        data: wav(10000),
        x: inches(6),
        y: inches(1),
        w: inches(2),
        h: inches(2),
      });
      const source = await savePresentation(pres);
      const sourcePath = join(dir, 'source.pptx');
      await writeFile(sourcePath, source);
      const entry = join(dir, 'deck.tsx');
      await writeFile(
        entry,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(sourcePath)})} />;`,
      );
      preview = await startPreview(entry);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 900 } });
      const editor = page.frameLocator('#editor-frame');
      await page.goto(preview.url);
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.hit').first().click();
      const inline = editor.locator('[aria-label="Media controls"]');
      assert.deepEqual(
        await inline
          .locator('button')
          .evaluateAll((nodes) => nodes.map((node) => node.getAttribute('aria-label'))),
        ['Play', 'Back', 'Forward', 'Mute'],
      );
      const controlsBox = await inline.boundingBox();
      const canvasBox = await editor.locator('.canvas-area').boundingBox();
      assert.ok(controlsBox.x >= canvasBox.x - 1);
      assert.ok(controlsBox.x + controlsBox.width <= canvasBox.x + canvasBox.width + 1);
      await page.screenshot({ path: '/tmp/pptx-inline-media-controls.png' });
      await editor.getByRole('tab', { name: 'Playback', exact: true }).click();
      const audio = editor.locator('audio');
      await audio.waitFor({ state: 'attached' });
      await audio.evaluate((element) =>
        element.readyState >= 1
          ? undefined
          : new Promise((resolve) =>
              element.addEventListener('loadedmetadata', resolve, { once: true }),
            ),
      );
      await audio.evaluate((element) => {
        element.currentTime = 6.7470924;
        element.dispatchEvent(new Event('timeupdate'));
      });
      const oldAudio = await audio.elementHandle();
      const startTime = await audio.evaluate((element) => element.currentTime);
      await inline.getByRole('button', { name: 'Play', exact: true }).click();
      await page.waitForFunction(
        ({ element, start }) => element.currentTime > start + 1,
        { element: oldAudio, start: startTime },
        { timeout: 5000 },
      );
      assert.equal(await oldAudio.evaluate((element) => element.paused), false);
      await editor.locator('.hit').nth(1).click();
      assert.equal(await oldAudio.evaluate((element) => element.paused), true);
      await editor.locator('.hit').first().click();
      await inline.getByRole('button', { name: 'Back', exact: true }).click();
      const stepped = await audio.evaluate((element) => element.currentTime);
      await inline.getByRole('button', { name: 'Forward', exact: true }).click();
      assert.ok(
        Math.abs((await audio.evaluate((element) => element.currentTime)) - stepped - 0.25) < 0.03,
      );
      await audio.evaluate((element) => {
        element.currentTime = 6.7470924;
        element.dispatchEvent(new Event('timeupdate'));
      });
      const add = editor.getByRole('button', { name: 'Add Bookmark', exact: true });
      await add.click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const readPlayback = async () => {
        const saved = await loadPresentation(
          new Uint8Array(await (await fetch(`${preview.url}/deck.pptx`)).arrayBuffer()),
        );
        return getShapeMediaPlayback(getSlideShapes(getSlides(saved)[0])[0]);
      };
      await inline.getByRole('button', { name: 'Mute', exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal((await readPlayback()).muted, true);
      await inline.getByRole('button', { name: 'Unmute', exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal((await readPlayback()).muted, false);
      const savedPlayback = await readPlayback();
      assert.deepEqual(
        savedPlayback.bookmarks?.map(({ name }) => name),
        ['Bookmark 1'],
      );
      assert.ok(Math.abs(savedPlayback.bookmarks[0].timeMs - 6747.0924) < 0.001);
      await editor.getByRole('button', { name: 'Volume', exact: true }).click();
      await editor.getByRole('menuitemradio', { name: 'Mute', exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal((await readPlayback()).muted, true);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal((await readPlayback()).muted, false);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal((await readPlayback()).bookmarks?.length ?? 0, 0);
      await add.click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await page.reload();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.hit').first().click();
      await editor.getByRole('tab', { name: 'Playback', exact: true }).click();
      await editor.locator('.bookmark[aria-label="Bookmark 1"]').click();
      assert.equal(
        await editor.getByRole('button', { name: 'Remove Bookmark', exact: true }).isEnabled(),
        true,
      );
      await editor.getByRole('button', { name: 'Remove Bookmark', exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal((await readPlayback()).bookmarks?.length ?? 0, 0);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'inline preview honors trim, replay, rewind, loop, and fade settings',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-media-inline-trim-'));
    let browser;
    let preview;
    try {
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      addSlideMedia(slide, {
        kind: 'audio',
        data: wav(4000),
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(2),
      });
      const shape = getSlideShapes(getSlides(pres)[0])[0];
      setShapeMediaPlayback(shape, {
        trim: { startMs: 1000, endMs: 1000 },
        rewindAfterPlaying: true,
        fade: { inMs: 500, outMs: 500 },
        volume: 1,
        muted: false,
      });
      const sourcePath = join(dir, 'source.pptx');
      await writeFile(sourcePath, await savePresentation(pres));
      const entry = join(dir, 'deck.tsx');
      await writeFile(
        entry,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(sourcePath)})} />;`,
      );
      preview = await startPreview(entry);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 900 } });
      const editor = page.frameLocator('#editor-frame');
      await page.goto(preview.url);
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.hit').first().click();
      const inline = editor.locator('[aria-label="Media controls"]');
      const audio = editor.locator('audio');
      await audio.waitFor({ state: 'attached' });
      await audio.evaluate(
        (element) =>
          new Promise((resolve) =>
            element.readyState >= 1
              ? resolve()
              : element.addEventListener('loadedmetadata', resolve, { once: true }),
          ),
      );
      assert.ok((await audio.evaluate((element) => element.currentTime)) >= 0.99);
      await audio.evaluate((element) => {
        element.currentTime = 3;
        element.dispatchEvent(new Event('timeupdate'));
      });
      await inline.getByRole('button', { name: 'Play', exact: true }).click();
      assert.ok((await audio.evaluate((element) => element.currentTime)) <= 1.05);
      await audio.evaluate((element) => element.pause());
      for (const [time, volume] of [
        [1.25, 0.5],
        [2, 1],
        [2.75, 0.5],
      ]) {
        await audio.evaluate((element, time) => {
          element.currentTime = time;
          element.dispatchEvent(new Event('timeupdate'));
        }, time);
        assert.ok(Math.abs((await audio.evaluate((element) => element.volume)) - volume) < 0.02);
      }
      await inline.getByRole('button', { name: 'Play', exact: true }).click();
      await new Promise((resolve) => setTimeout(resolve, 700));
      assert.equal(await audio.evaluate((element) => element.paused), true);
      assert.ok(Math.abs((await audio.evaluate((element) => element.currentTime)) - 1) < 0.02);
      await editor.getByRole('tab', { name: 'Playback', exact: true }).click();
      await audio.evaluate((element) => {
        element.currentTime = 3;
        element.dispatchEvent(new Event('timeupdate'));
      });
      await editor.locator('.ribbon').getByRole('button', { name: 'Play', exact: true }).click();
      await audio.evaluate(
        (element) =>
          new Promise((resolve, reject) => {
            const deadline = Date.now() + 1000;
            const check = () => {
              if (!element.paused && element.currentTime < 1.3) resolve();
              else if (Date.now() > deadline)
                reject(
                  new Error(
                    `Ribbon replay failed: ${element.currentTime}, paused=${element.paused}`,
                  ),
                );
              else setTimeout(check, 10);
            };
            check();
          }),
      );
      await audio.evaluate((element) => element.pause());
      await editor.getByLabel('Loop until stopped', { exact: true }).check();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await audio.evaluate((element) => {
        element.currentTime = 2.9;
        element.dispatchEvent(new Event('timeupdate'));
      });
      await inline.getByRole('button', { name: 'Play', exact: true }).click();
      await new Promise((resolve) => setTimeout(resolve, 400));
      assert.equal(await audio.evaluate((element) => element.paused), false);
      assert.ok((await audio.evaluate((element) => element.currentTime)) < 2);
      await audio.evaluate((element) => element.pause());
      assert.equal(await inline.getByRole('slider', { name: 'Volume', exact: true }).count(), 0);
      await inline.getByRole('button', { name: 'Mute', exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const readMuted = async () => {
        const saved = await loadPresentation(
          new Uint8Array(await (await fetch(`${preview.url}/deck.pptx`)).arrayBuffer()),
        );
        return getShapeMediaPlayback(getSlideShapes(getSlides(saved)[0])[0]).muted;
      };
      assert.equal(await readMuted(), true);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal(await readMuted(), false);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

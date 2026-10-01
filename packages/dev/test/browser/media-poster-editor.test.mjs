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
  getShapeImageBytes,
  getShapeMedia,
  getShapeMediaPlayback,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
  setShapeMediaPlayback,
  savePresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const videoBytes = (page) =>
  page.evaluate(async () => {
    const canvas = document.createElement('canvas');
    canvas.width = 160;
    canvas.height = 90;
    const context = canvas.getContext('2d');
    const stream = canvas.captureStream(10);
    const recorder = new MediaRecorder(stream, { mimeType: 'video/webm;codecs=vp8' });
    const parts = [];
    const stopped = new Promise((resolve) => (recorder.onstop = resolve));
    recorder.ondataavailable = (event) => parts.push(event.data);
    recorder.start();
    let frame = 0;
    const timer = setInterval(() => {
      context.fillStyle = frame++ % 2 ? '#2d9cdb' : '#eb5757';
      context.fillRect(0, 0, canvas.width, canvas.height);
    }, 50);
    await new Promise((resolve) => setTimeout(resolve, 600));
    recorder.stop();
    await stopped;
    clearInterval(timer);
    stream.getTracks().forEach((track) => track.stop());
    return Array.from(new Uint8Array(await new Blob(parts).arrayBuffer()));
  });

const firstVideoFrameBytes = (page, media) =>
  page.evaluate(async (bytes) => {
    const video = document.createElement('video');
    video.muted = true;
    video.preload = 'auto';
    const url = URL.createObjectURL(new Blob([new Uint8Array(bytes)], { type: 'video/webm' }));
    try {
      await new Promise((resolve, reject) => {
        const deadline = window.setTimeout(
          () => reject(new Error('video did not decode its first frame')),
          5000,
        );
        video.addEventListener(
          'loadeddata',
          () => {
            window.clearTimeout(deadline);
            resolve();
          },
          { once: true },
        );
        video.addEventListener(
          'error',
          () => {
            window.clearTimeout(deadline);
            reject(new Error('video failed to decode its first frame'));
          },
          { once: true },
        );
        video.src = url;
        video.load();
      });
      const canvas = document.createElement('canvas');
      canvas.width = video.videoWidth;
      canvas.height = video.videoHeight;
      canvas.getContext('2d').drawImage(video, 0, 0);
      const blob = await new Promise((resolve, reject) =>
        canvas.toBlob(
          (value) => (value ? resolve(value) : reject(new Error('failed to encode first frame'))),
          'image/png',
        ),
      );
      return Array.from(new Uint8Array(await blob.arrayBuffer()));
    } finally {
      video.removeAttribute('src');
      video.load();
      URL.revokeObjectURL(url);
    }
  }, Array.from(media));

test(
  'video poster frame changes preserve media bytes, support file replacement, undo, and reload',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-media-poster-editor-'));
    let browser;
    let preview;
    try {
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 900 } });
      const media = Uint8Array.from(await videoBytes(page));
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      const videoShape = addSlideMedia(slide, {
        kind: 'video',
        data: media,
        x: inches(1),
        y: inches(1),
        w: inches(6),
        h: inches(3.375),
      });
      setShapeMediaPlayback(videoShape, {
        autoplay: true,
        delayMs: 250,
        loop: true,
        rewindAfterPlaying: true,
        volume: 0.8,
      });
      const source = join(dir, 'source.pptx');
      await writeFile(source, await savePresentation(pres));
      const entry = join(dir, 'deck.tsx');
      await writeFile(
        entry,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
      );
      preview = await startPreview(entry);
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      await saved();
      await editor.locator('.hit').first().click();
      await editor.getByRole('tab', { name: 'Video Format', exact: true }).click();
      const video = editor.locator('.media-preview video');
      await video.waitFor({ state: 'attached' });
      await video.evaluate(
        (element) =>
          new Promise((resolve, reject) => {
            const deadline = window.setTimeout(
              () => reject(new Error('video did not present a decoded frame')),
              5000,
            );
            if (element.readyState >= 2) {
              window.clearTimeout(deadline);
              resolve();
            } else {
              element.addEventListener(
                'loadeddata',
                () => {
                  window.clearTimeout(deadline);
                  resolve();
                },
                { once: true },
              );
            }
          }),
      );
      const panel = editor.locator('#ribbon-panel');
      const mediaShape = async () =>
        getSlideShapes(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        )[0];
      const readPosterAndMedia = async () => {
        const shape = await mediaShape();
        return {
          poster: Buffer.from(getShapeImageBytes(shape)),
          media: Buffer.from(getShapeMedia(shape).bytes),
        };
      };
      const original = await readPosterAndMedia();
      assert.deepEqual(original.media, Buffer.from(media));
      const readPlayback = async () => {
        const saved = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getShapeMediaPlayback(getSlideShapes(getSlides(saved)[0])[0]);
      };
      const originalPlayback = await readPlayback();
      assert.deepEqual(originalPlayback, {
        autoplay: true,
        delayMs: 250,
        loop: true,
        volume: 0.8,
        muted: false,
        fullScreen: false,
        hideWhenStopped: false,
        rewindAfterPlaying: true,
      });

      // The saved badge describes the editor's local state and can become
      // visible before the preview server has finished serializing the deck.
      // Poll the authoritative PPTX until the requested mutation is present.
      const waitForPersisted = async (predicate, label) => {
        const deadline = Date.now() + 5000;
        let lastValue;
        while (Date.now() < deadline) {
          lastValue = await readPosterAndMedia();
          if (predicate(lastValue)) return lastValue;
          await new Promise((resolve) => setTimeout(resolve, 50));
        }
        assert.fail(`timed out waiting for persisted ${label}`);
      };
      const waitForPlayback = async () => {
        const deadline = Date.now() + 5000;
        let lastValue;
        while (Date.now() < deadline) {
          lastValue = await readPlayback();
          if (JSON.stringify(lastValue) === JSON.stringify(originalPlayback)) return lastValue;
          await new Promise((resolve) => setTimeout(resolve, 50));
        }
        assert.deepEqual(lastValue, originalPlayback);
      };

      // A selected video keeps its poster visible until playback or seeking
      // begins; the decoded <video> element must not paint over it while idle.
      assert.equal(
        await video.evaluate((element) => getComputedStyle(element).visibility),
        'hidden',
      );
      await panel.getByRole('button', { name: /Poster Frame/i }).click();
      const currentFrameItem = panel.getByRole('menuitem', { name: 'Current Frame', exact: true });
      assert.equal(await currentFrameItem.isDisabled(), true);
      await editor.locator('body').press('Escape');
      await editor.getByRole('button', { name: 'Play', exact: true }).click();
      await video.evaluate(
        (element) =>
          new Promise((resolve, reject) => {
            const deadline = Date.now() + 3000;
            const check = () => {
              if (!element.paused && element.currentTime > 0.1) resolve();
              else if (Date.now() >= deadline) reject(new Error('video did not advance'));
              else setTimeout(check, 20);
            };
            check();
          }),
      );
      await editor.getByRole('button', { name: 'Pause', exact: true }).click();
      await video.evaluate((element) => {
        if (!element.paused) throw new Error('video did not pause');
      });

      // The current-frame command captures the decoded frame into the poster image.
      await panel.getByRole('button', { name: /Poster Frame/i }).click();
      await panel.getByRole('menuitem', { name: 'Current Frame', exact: true }).click();
      const currentFrame = await waitForPersisted(
        ({ poster, media: persistedMedia }) =>
          !poster.equals(original.poster) && persistedMedia.equals(original.media),
        'current-frame poster',
      );
      assert.notDeepEqual(currentFrame.poster, original.poster);
      assert.deepEqual(await waitForPlayback(), originalPlayback);

      // PowerPoint's Reset command replaces the poster with the video's first
      // decoded frame. Decode a separate video element so the expected bytes
      // do not depend on the editor's current playback position.
      const firstFrame = Buffer.from(await firstVideoFrameBytes(page, media));

      const replacement = Buffer.from(
        await page.evaluate(() => {
          const canvas = document.createElement('canvas');
          canvas.width = 320;
          canvas.height = 180;
          const context = canvas.getContext('2d');
          context.fillStyle = '#45b36b';
          context.fillRect(0, 0, canvas.width, canvas.height);
          return canvas.toDataURL('image/png').split(',')[1];
        }),
        'base64',
      );
      await panel.getByRole('button', { name: /Poster Frame/i }).click();
      await panel
        .getByRole('menuitem', { name: /(?:Picture|Image) from File\.\.\./i, exact: true })
        .click();
      const dialog = editor.getByRole('dialog', { name: /Insert image|Replace image/i });
      await dialog
        .getByLabel('Image file', { exact: true })
        .setInputFiles({ name: 'poster.png', mimeType: 'image/png', buffer: replacement });
      await dialog.getByRole('button', { name: /Insert image|Replace image/i }).click();
      await dialog.waitFor({ state: 'hidden' });
      const filePoster = await waitForPersisted(
        ({ poster, media: persistedMedia }) =>
          poster.equals(replacement) && persistedMedia.equals(original.media),
        'file poster',
      );
      assert.deepEqual(filePoster.poster, replacement);
      assert.deepEqual(filePoster.media, original.media);
      assert.deepEqual(await waitForPlayback(), originalPlayback);

      // Keep the earlier poster replacement undo path covered before testing
      // Reset's own history entry.
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await waitForPersisted(
        ({ poster, media: persistedMedia }) =>
          poster.equals(currentFrame.poster) && persistedMedia.equals(original.media),
        'undo file poster',
      );
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await waitForPersisted(
        ({ poster, media: persistedMedia }) =>
          poster.equals(filePoster.poster) && persistedMedia.equals(original.media),
        'redo file poster',
      );

      await video.evaluate(
        (element) =>
          new Promise((resolve, reject) => {
            const target = Math.min(0.35, Math.max(0, element.duration - 0.05));
            const deadline = window.setTimeout(
              () => reject(new Error('video did not seek before poster reset')),
              3000,
            );
            element.addEventListener(
              'seeked',
              () => {
                window.clearTimeout(deadline);
                element.pause();
                resolve();
              },
              { once: true },
            );
            element.currentTime = target;
          }),
      );
      await panel.getByRole('button', { name: /Poster Frame/i }).click();
      const playbackPositionBeforeReset = await video.evaluate((element) => element.currentTime);
      await panel.getByRole('menuitem', { name: 'Reset', exact: true }).click();
      const reset = await waitForPersisted(
        ({ poster, media: persistedMedia }) =>
          poster.equals(firstFrame) && persistedMedia.equals(original.media),
        'reset poster',
      );
      assert.deepEqual(reset.poster, firstFrame);
      assert.deepEqual(reset.media, original.media);
      assert.deepEqual(await waitForPlayback(), originalPlayback);
      assert.ok(
        Math.abs(
          (await video.evaluate((element) => element.currentTime)) - playbackPositionBeforeReset,
        ) < 0.01,
      );

      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      const undone = await waitForPersisted(
        ({ poster, media: persistedMedia }) =>
          poster.equals(filePoster.poster) && persistedMedia.equals(original.media),
        'undo poster',
      );
      assert.deepEqual(undone.poster, filePoster.poster);
      assert.deepEqual(undone.media, original.media);
      assert.deepEqual(await waitForPlayback(), originalPlayback);

      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      const redone = await waitForPersisted(
        ({ poster, media: persistedMedia }) =>
          poster.equals(firstFrame) && persistedMedia.equals(original.media),
        'redo reset poster',
      );
      assert.deepEqual(redone.poster, firstFrame);
      assert.deepEqual(redone.media, original.media);
      assert.deepEqual(await waitForPlayback(), originalPlayback);

      await page.reload();
      await saved();
      await editor.locator('.hit').first().click();
      await editor.getByRole('tab', { name: 'Video Format', exact: true }).click();
      const reloaded = await waitForPersisted(
        ({ poster, media: persistedMedia }) =>
          poster.equals(firstFrame) && persistedMedia.equals(original.media),
        'reloaded poster',
      );
      assert.deepEqual(reloaded.poster, firstFrame);
      assert.deepEqual(reloaded.media, original.media);
      assert.deepEqual(await waitForPlayback(), originalPlayback);

      await editor.locator('.lang select').selectOption('ja');
      await editor.getByRole('button', { name: '表紙画像', exact: true }).click();
      await editor.getByRole('menuitem', { name: 'リセット', exact: true }).press('Escape');
      assert.equal(await editor.getByRole('menu').count(), 0);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

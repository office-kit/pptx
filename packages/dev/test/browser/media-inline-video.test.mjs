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
  getShapeId,
  inches,
  savePresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const videoBytes = (page, colors) =>
  page.evaluate(async ([first, second]) => {
    const canvas = document.createElement('canvas');
    canvas.width = 160;
    canvas.height = 90;
    const context = canvas.getContext('2d');
    const stream = canvas.captureStream(10);
    const recorder = new MediaRecorder(stream, { mimeType: 'video/webm;codecs=vp8' });
    const parts = [];
    recorder.ondataavailable = (event) => parts.push(event.data);
    const stopped = new Promise((resolve) => (recorder.onstop = resolve));
    recorder.start();
    let frame = 0;
    const timer = setInterval(() => {
      context.fillStyle = frame++ % 2 ? first : second;
      context.fillRect(0, 0, 160, 90);
    }, 50);
    await new Promise((resolve) => setTimeout(resolve, 600));
    recorder.stop();
    await stopped;
    clearInterval(timer);
    stream.getTracks().forEach((track) => track.stop());
    return Array.from(new Uint8Array(await new Blob(parts).arrayBuffer()));
  }, colors);

const waitForDecodedFrame = (video) =>
  video.evaluate(
    (element) =>
      new Promise((resolve, reject) => {
        const deadline = window.setTimeout(
          () => reject(new Error('video did not present a decoded frame')),
          3000,
        );
        element.requestVideoFrameCallback((_timestamp, metadata) => {
          window.clearTimeout(deadline);
          resolve(metadata.presentedFrames ?? 1);
        });
      }),
  );

test(
  'selected video inline preview decodes, stops on selection change, and replaces source across slides',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-media-inline-video-'));
    let browser;
    let preview;
    try {
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 900 } });
      const firstBytes = await videoBytes(page, ['red', 'blue']);
      const secondBytes = await videoBytes(page, ['lime', 'purple']);

      const presentation = createPresentation();
      const firstSlide = addBlankSlide(presentation);
      const secondSlide = addBlankSlide(presentation);
      const firstShape = addSlideMedia(firstSlide, {
        kind: 'video',
        data: Uint8Array.from(firstBytes),
        x: inches(1),
        y: inches(1),
        w: inches(6),
        h: inches(3.375),
      });
      const secondShape = addSlideMedia(secondSlide, {
        kind: 'video',
        data: Uint8Array.from(secondBytes),
        x: inches(8.5),
        y: inches(1),
        w: inches(4.5),
        h: inches(3.375),
      });
      assert.equal(
        getShapeId(firstShape),
        getShapeId(secondShape),
        'fixture must exercise the same shape ID on different slides',
      );

      const sourcePath = join(dir, 'source.pptx');
      await writeFile(sourcePath, await savePresentation(presentation));
      const entry = join(dir, 'deck.tsx');
      await writeFile(
        entry,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(sourcePath)})} />;`,
      );
      preview = await startPreview(entry);
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();

      await editor.locator('.hit').first().click();
      const firstVideo = editor.locator('.media-preview video');
      await firstVideo.waitFor({ state: 'attached' });
      await firstVideo.evaluate(
        (element) =>
          new Promise((resolve) =>
            element.readyState >= 1
              ? resolve()
              : element.addEventListener('loadedmetadata', resolve, { once: true }),
          ),
      );
      assert.deepEqual(
        await firstVideo.evaluate((element) => [element.videoWidth, element.videoHeight]),
        [160, 90],
      );
      const firstSource = await firstVideo.getAttribute('src');
      const firstHandle = await firstVideo.elementHandle();
      assert.ok(firstHandle);
      await editor.getByRole('button', { name: 'Play', exact: true }).click();
      await firstVideo.evaluate(
        (element) =>
          new Promise((resolve, reject) => {
            const deadline = Date.now() + 3000;
            const check = () => {
              if (!element.paused && element.currentTime > 0.1) resolve();
              else if (Date.now() >= deadline)
                reject(new Error('video did not advance during playback'));
              else setTimeout(check, 20);
            };
            check();
          }),
      );
      assert.ok((await waitForDecodedFrame(firstVideo)) > 0);

      await editor.locator('.thumb-row').nth(1).click();
      assert.equal(
        await firstHandle.evaluate((element) => element.paused),
        true,
        'changing slides must pause the previously selected video',
      );
      await editor.locator('.hit').first().click();
      const secondVideo = editor.locator('.media-preview video');
      await secondVideo.waitFor({ state: 'attached' });
      await secondVideo.evaluate(
        (element) =>
          new Promise((resolve) =>
            element.readyState >= 1
              ? resolve()
              : element.addEventListener('loadedmetadata', resolve, { once: true }),
          ),
      );
      assert.deepEqual(
        await secondVideo.evaluate((element) => [element.videoWidth, element.videoHeight]),
        [160, 90],
      );
      assert.notEqual(await secondVideo.getAttribute('src'), firstSource);
      const controlsInsideCanvas = async () =>
        editor.locator('.media-preview .controls').evaluate((node) => {
          const canvas = node.closest('.canvas-area').getBoundingClientRect();
          const controls = node.getBoundingClientRect();
          return {
            left: controls.left >= canvas.left - 1,
            right: controls.right <= canvas.right + 1,
          };
        });
      assert.deepEqual(await controlsInsideCanvas(), { left: true, right: true });
      await page.setViewportSize({ width: 900, height: 900 });
      await page.waitForTimeout(100);
      assert.deepEqual(await controlsInsideCanvas(), { left: true, right: true });
      await editor.getByRole('button', { name: 'Play', exact: true }).click();
      await secondVideo.evaluate(
        (element) =>
          new Promise((resolve, reject) => {
            const deadline = Date.now() + 3000;
            const check = () => {
              if (!element.paused && element.currentTime > 0.1) resolve();
              else if (Date.now() >= deadline)
                reject(new Error('replacement video did not advance'));
              else setTimeout(check, 20);
            };
            check();
          }),
      );
      assert.ok((await waitForDecodedFrame(secondVideo)) > 0);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

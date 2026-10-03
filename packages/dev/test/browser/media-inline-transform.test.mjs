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
  inches,
  savePresentation,
  setShapeRotation,
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
    recorder.ondataavailable = (event) => parts.push(event.data);
    const stopped = new Promise((resolve) => (recorder.onstop = resolve));
    recorder.start();
    const timer = setInterval(() => {
      context.fillStyle = '#4f46e5';
      context.fillRect(0, 0, 160, 90);
    }, 50);
    await new Promise((resolve) => setTimeout(resolve, 500));
    recorder.stop();
    await stopped;
    clearInterval(timer);
    stream.getTracks().forEach((track) => track.stop());
    return Array.from(new Uint8Array(await new Blob(parts).arrayBuffer()));
  });

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
  'rotated embedded media keeps its inline toolbar horizontal and usable',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-media-inline-transform-'));
    let browser;
    let preview;
    try {
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 900 } });
      const presentation = createPresentation();
      const slide = addBlankSlide(presentation);
      const shape = addSlideMedia(slide, {
        kind: 'video',
        data: Uint8Array.from(await videoBytes(page)),
        x: inches(4),
        y: inches(2),
        w: inches(5),
        h: inches(2.8125),
      });
      setShapeRotation(shape, 45);
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
      const inline = editor.locator('.media-preview');
      const video = inline.locator('video');
      await video.waitFor({ state: 'attached' });
      await editor.locator('[aria-label="Media controls"]').waitFor({ state: 'visible' });
      const toolbar = editor.locator('[aria-label="Media controls"]');
      const toolbarBox = await toolbar.boundingBox();
      assert.ok(toolbarBox, 'rotated media must render its toolbar');
      assert.ok(
        toolbarBox.width > toolbarBox.height * 4,
        `toolbar should remain horizontal, got ${toolbarBox.width}×${toolbarBox.height}`,
      );
      const buttons = toolbar.locator('button');
      assert.equal(await buttons.count(), 4);
      for (let index = 0; index < 4; index++) {
        const buttonBox = await buttons.nth(index).boundingBox();
        assert.ok(buttonBox && buttonBox.width > 0 && buttonBox.height > 0);
      }
      await toolbar.getByRole('button', { name: 'Play', exact: true }).click();
      await video.evaluate(
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
      assert.ok((await waitForDecodedFrame(video)) > 0);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

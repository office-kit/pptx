import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { compile, Media, Presentation, Slide } from '@office-kit/pptx-dsl';
import {
  getSlideShapes,
  getSlides,
  savePresentation,
  setShapeMediaPlayback,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'embedded video decodes and plays without advancing the slide',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-video-playback-'));
    let browser, preview;
    try {
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage();
      // A real browser-generated clip exercises decoding without an external encoder dependency.
      const bytes = await page.evaluate(async () => {
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
          context.fillStyle = frame++ % 2 ? 'red' : 'blue';
          context.fillRect(0, 0, 160, 90);
        }, 50);
        await new Promise((resolve) => setTimeout(resolve, 600));
        recorder.stop();
        await stopped;
        clearInterval(timer);
        stream.getTracks().forEach((track) => track.stop());
        return Array.from(new Uint8Array(await new Blob(parts).arrayBuffer()));
      });
      const deck = await compile(
        Presentation({
          children: Slide({
            children: Media({
              kind: 'video',
              data: Uint8Array.from(bytes),
              x: 1,
              y: 1,
              width: 4,
              height: 2.25,
            }),
          }),
        }),
      );
      setShapeMediaPlayback(getSlideShapes(getSlides(deck)[0])[0], {
        autoplay: false,
        muted: true,
        hideWhenStopped: true,
      });
      await writeFile(join(dir, 'source.pptx'), await savePresentation(deck));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(join(dir, 'source.pptx'))})} />;`,
      );
      preview = await startPreview(file);
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      const video = page.locator('#slide video');
      await video.waitFor();
      assert.equal(await video.evaluate((element) => element.paused), true);
      await video.focus();
      await page.keyboard.press('Space');
      await page.waitForFunction(
        () => document.querySelector('#slide').shadowRoot.querySelector('video')?.currentTime > 0.1,
      );
      assert.deepEqual(
        await video.evaluate((element) => [element.videoWidth, element.videoHeight]),
        [160, 90],
      );
      assert.equal(await page.evaluate(() => index), 0);
      await page.waitForFunction(
        () => document.querySelector('#slide').shadowRoot.querySelector('video')?.ended,
      );
      assert.equal(
        await page
          .locator('#slide foreignObject')
          .evaluate((element) => getComputedStyle(element).visibility),
        'hidden',
      );
      await page.keyboard.press('Escape');
      await video.waitFor({ state: 'detached' });
      assert.equal(await page.locator('#slide image').count(), 1);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

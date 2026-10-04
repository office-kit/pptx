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
  'full-screen video uses one media element and returns to its poster on end or exit',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-video-fullscreen-'));
    let browser;
    let preview;
    try {
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
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
        await new Promise((resolve) => setTimeout(resolve, 1500));
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
        fullScreen: true,
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

      const inlinePlay = page.getByRole('button', { name: /Play media|メディアを再生/ });
      await inlinePlay.focus();
      await inlinePlay.click();
      const overlay = page.locator('[data-pptx-media-fullscreen]');
      const overlayVideo = overlay.locator('video');
      await overlayVideo.waitFor({ state: 'attached' });
      assert.deepEqual(await overlay.boundingBox(), { x: 0, y: 0, width: 1280, height: 800 });
      const original = await overlayVideo.elementHandle();
      assert.ok(original);
      assert.equal(
        await overlayVideo.evaluate((element) => document.activeElement === element),
        true,
      );
      await page.waitForFunction((element) => element.currentTime > 0.1, original);
      await overlayVideo.evaluate((element) => element.pause());
      assert.equal(await overlayVideo.evaluate((element) => element.paused), true);
      await overlayVideo.evaluate((element) => element.play());
      assert.equal(await overlayVideo.evaluate((element) => element.paused), false);
      assert.equal(await overlayVideo.evaluate((element) => element.isConnected), true);
      await page.waitForFunction((element) => element.ended, original);
      await page.waitForFunction(() => {
        const element = document.querySelector('[data-pptx-media-fullscreen]');
        return element instanceof HTMLElement && element.hidden;
      });
      assert.equal(await overlayVideo.evaluate((element) => element.isConnected), true);
      assert.equal(await page.locator('#slide foreignObject[data-pptx-media]').count(), 1);
      assert.match(
        await page.locator('#slide foreignObject[data-pptx-media] div').getAttribute('style'),
        /background/,
      );
      assert.equal(
        await inlinePlay.evaluate((element) => element.getRootNode().activeElement === element),
        true,
      );

      await inlinePlay.click();
      await overlayVideo.waitFor({ state: 'attached' });
      assert.equal(
        await overlayVideo.evaluate((element, prior) => element === prior, original),
        true,
      );
      await page.keyboard.press('Escape');
      await overlay.waitFor({ state: 'detached' });
      assert.equal(await original.evaluate((element) => element.isConnected), false);
      assert.equal(await page.locator('#slide image').count(), 1);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'autoplay full-screen video hides the inline host and returns focus to the stage',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-video-fullscreen-autoplay-'));
    let browser;
    let preview;
    try {
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
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
        const timer = setInterval(() => {
          context.fillStyle = 'purple';
          context.fillRect(0, 0, 160, 90);
        }, 50);
        await new Promise((resolve) => setTimeout(resolve, 1500));
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
        autoplay: true,
        muted: true,
        fullScreen: true,
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
      const overlay = page.locator('[data-pptx-media-fullscreen]');
      const video = overlay.locator('video');
      await video.waitFor({ state: 'attached' });
      const original = await video.elementHandle();
      assert.ok(original);
      assert.equal(await video.evaluate((element) => document.activeElement === element), true);
      await page.waitForFunction((element) => element.ended, original);
      await page.waitForFunction(() => {
        const element = document.querySelector('[data-pptx-media-fullscreen]');
        return element instanceof HTMLElement && element.hidden;
      });
      assert.equal(
        await page
          .locator('#slide foreignObject[data-pptx-media]')
          .evaluate((element) => getComputedStyle(element).visibility),
        'hidden',
      );
      assert.equal(
        await page.locator('#stage').evaluate((element) => document.activeElement === element),
        true,
      );
      assert.equal(await original.evaluate((element) => element.isConnected), true);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'native video fullscreen during a full-screen show stays until the viewer leaves it',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-video-native-fullscreen-'));
    let browser;
    let preview;
    try {
      const deck = await compile(
        Presentation({
          children: Slide({
            children: Media({
              kind: 'video',
              // An MP4 header is enough for a <video> with native controls.
              data: Uint8Array.from([
                0, 0, 0, 0x18, 0x66, 0x74, 0x79, 0x70, 0x6d, 0x70, 0x34, 0x32, 0, 0, 0, 0, 0x6d,
                0x70, 0x34, 0x32,
              ]),
              x: 1,
              y: 1,
              width: 4,
              height: 2.25,
            }),
          }),
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
      const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      await page.waitForFunction(() => document.fullscreenElement === document.documentElement);
      const video = page.locator('video');
      assert.equal(await video.evaluate((element) => element.controls), true);
      // The video lives in a shadow root, so the document reports its host;
      // `:fullscreen` tells the video itself apart from the page.
      await video.evaluate((element) => {
        window.fullscreenTrace = [];
        document.addEventListener('fullscreenchange', () =>
          window.fullscreenTrace.push(
            element.matches(':fullscreen')
              ? 'video'
              : document.fullscreenElement === document.documentElement
                ? 'page'
                : null,
          ),
        );
      });
      // Playwright evaluates with a user gesture, as the native control's
      // fullscreen button would provide.
      await video.evaluate((element) => element.requestFullscreen());
      await page.waitForFunction(() => window.fullscreenTrace.length === 1);
      assert.equal(await video.evaluate((element) => element.matches(':fullscreen')), true);
      // Leaving the video's fullscreen returns to the show's own fullscreen.
      await page.evaluate(() => document.exitFullscreen());
      await page.waitForFunction(() => window.fullscreenTrace.length === 2);
      assert.deepEqual(await page.evaluate(() => window.fullscreenTrace), ['video', 'page']);
      assert.equal(
        await page.evaluate(() => document.fullscreenElement === document.documentElement),
        true,
      );
      assert.equal(await page.evaluate(() => presenting), true);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

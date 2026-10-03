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

for (const fullScreen of [false, true]) {
  test(
    `presenter media controls follow audience playback and navigation (${fullScreen ? 'fullscreen' : 'inline'})`,
    { timeout: 60000 },
    async () => {
      const directory = await mkdtemp(join(tmpdir(), 'office-presenter-media-'));
      let browser;
      let preview;
      try {
        browser = await chromium.launch({ headless: true });
        const context = await browser.newContext({ viewport: { width: 1400, height: 900 } });
        const page = await context.newPage();
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
          await new Promise((resolve) => setTimeout(resolve, 5000));
          recorder.stop();
          await stopped;
          clearInterval(timer);
          stream.getTracks().forEach((track) => track.stop());
          return Array.from(new Uint8Array(await new Blob(parts).arrayBuffer()));
        });
        const deck = await compile(
          Presentation({
            children: [
              Slide({
                children: Media({
                  kind: 'video',
                  data: Uint8Array.from(bytes),
                  x: 1,
                  y: 1,
                  width: 4,
                  height: 2.25,
                }),
              }),
              Slide({
                children: Text({ x: 1, y: 1, width: 5, height: 1, children: 'Second slide' }),
              }),
            ],
          }),
        );
        setShapeMediaPlayback(getSlideShapes(getSlides(deck)[0])[0], {
          autoplay: false,
          muted: false,
          fullScreen,
        });
        await writeFile(join(directory, 'source.pptx'), await savePresentation(deck));
        const file = join(directory, 'deck.tsx');
        await writeFile(
          file,
          `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(join(directory, 'source.pptx'))})} />;`,
        );
        preview = await startPreview(file);
        await page.goto(preview.url);
        await page.getByRole('button', { name: 'Preview', exact: true }).click();
        await page.waitForFunction(() => state.slides.length === 2);
        const popup = page.waitForEvent('popup');
        await page.getByRole('button', { name: 'Presenter view', exact: true }).click();
        const presenter = await popup;
        await presenter.getByText('Slide 1 of 2', { exact: true }).waitFor();

        const audienceVideo = fullScreen
          ? page.locator('video').first()
          : page.locator('#slide video');
        const presenterVideo = presenter.locator('#current video');
        await audienceVideo.waitFor({ state: 'attached' });
        await presenterVideo.waitFor({ state: 'attached' });
        assert.equal(await audienceVideo.evaluate((element) => element.paused), true);
        assert.equal(await presenterVideo.evaluate((element) => element.paused), true);
        assert.equal(await audienceVideo.evaluate((element) => element.muted), false);
        assert.equal(await presenterVideo.evaluate((element) => element.muted), true);

        await presenter.getByRole('button', { name: 'Play media', exact: true }).click();
        await page.waitForFunction(
          (fullScreen) =>
            fullScreen
              ? document.querySelector('video')?.paused === false
              : document.querySelector('#slide').shadowRoot.querySelector('video')?.paused ===
                false,
          fullScreen,
        );
        await presenter.getByRole('button', { name: 'Pause media', exact: true }).waitFor();
        if (fullScreen) {
          const canvas = await presenter.locator('#current').boundingBox();
          const overlay = await presenter.locator('[data-pptx-media-fullscreen]').boundingBox();
          assert.ok(canvas && overlay);
          for (const dimension of ['x', 'y', 'width', 'height']) {
            assert.ok(Math.abs(canvas[dimension] - overlay[dimension]) < 1);
          }
        }
        await presenter.waitForFunction(
          () =>
            (
              document.querySelector('#current video') ??
              document.querySelector('#current')?.shadowRoot?.querySelector('video')
            )?.paused === false,
        );
        assert.equal(await presenterVideo.evaluate((element) => element.paused), false);
        await presenter.getByRole('button', { name: 'Pause media', exact: true }).click();
        await page.waitForFunction(
          (fullScreen) =>
            fullScreen
              ? document.querySelector('video')?.paused === true
              : document.querySelector('#slide').shadowRoot.querySelector('video')?.paused === true,
          fullScreen,
        );
        await presenter.waitForFunction(
          () =>
            (
              document.querySelector('#current video') ??
              document.querySelector('#current')?.shadowRoot?.querySelector('video')
            )?.paused === true,
        );
        assert.equal(await audienceVideo.evaluate((element) => element.paused), true);
        assert.equal(await presenterVideo.evaluate((element) => element.paused), true);

        await audienceVideo.evaluate((element) => element.play());
        await presenterVideo.waitFor({ state: 'attached' });
        await presenter.waitForFunction(
          () =>
            (
              document.querySelector('#current video') ??
              document.querySelector('#current')?.shadowRoot?.querySelector('video')
            )?.paused === false,
        );
        assert.equal(await presenterVideo.evaluate((element) => element.paused), false);
        const position = presenter.getByRole('slider', { name: 'Media position', exact: true });
        await position.fill('50');
        await page.waitForFunction(
          (fullScreen) =>
            (fullScreen
              ? document.querySelector('video')?.currentTime
              : document.querySelector('#slide').shadowRoot.querySelector('video')?.currentTime) >
            1,
          fullScreen,
        );
        assert.ok((await audienceVideo.evaluate((element) => element.currentTime)) > 1);
        await presenter.screenshot({
          path: `/tmp/pptx-presenter-media-${fullScreen ? 'fullscreen' : 'inline'}.png`,
        });

        await presenter.getByRole('button', { name: 'Next', exact: true }).click();
        await presenter.getByText('Slide 2 of 2', { exact: true }).waitFor();
        await page.waitForFunction(() => index === 1);
        await presenter.evaluate(() =>
          window.opener.postMessage(
            {
              type: 'presenter-command',
              action: 'media',
              mediaKey: 'stale-media-key',
              index: { shapeId: 1, action: 'play' },
            },
            location.origin,
          ),
        );
        await audienceVideo.waitFor({ state: 'detached' });
        await presenterVideo.waitFor({ state: 'detached' });
        await presenter.getByRole('button', { name: 'Previous', exact: true }).click();
        await presenter.getByText('Slide 1 of 2', { exact: true }).waitFor();
        const reenteredAudienceVideo = fullScreen
          ? page.locator('video').first()
          : page.locator('#slide video');
        const reenteredPresenterVideo = presenter.locator('#current video');
        await reenteredAudienceVideo.waitFor({ state: 'attached' });
        await reenteredPresenterVideo.waitFor({ state: 'attached' });
        await presenter.getByRole('button', { name: 'Exit presentation', exact: true }).click();
        await page.waitForFunction(() => !presenting);
        await reenteredAudienceVideo.waitFor({ state: 'detached' });
        await reenteredPresenterVideo.waitFor({ state: 'detached' });
        await presenter.close();
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(directory, { recursive: true, force: true });
      }
    },
  );
}

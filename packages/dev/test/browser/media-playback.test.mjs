import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { compile, Media, Presentation, Slide, Text } from '@office-kit/pptx-dsl';
import {
  getSlideShapes,
  getShapeId,
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

test(
  'native background audio timing autoplays hidden audio across the next slide',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-native-media-background-'));
    let preview;
    let browser;
    try {
      const deck = await compile(
        Presentation({
          children: [
            Slide({
              children: Media({ kind: 'audio', data: wav(15000), x: 1, y: 1, width: 3, height: 1 }),
            }),
            Slide({ children: Text({ x: 1, y: 1, width: 5, height: 1, children: 'Second' }) }),
          ],
        }),
      );
      const parts = unzipSync(await savePresentation(deck));
      const slidePath = 'ppt/slides/slide1.xml';
      const [audioShape] = getSlideShapes(getSlides(deck)[0]);
      const timing = (
        await readFile(
          new URL('../../../../test/fixtures/native-media-background-timing.xml', import.meta.url),
          'utf8',
        )
      ).replaceAll('spid="2"', `spid="${getShapeId(audioShape)}"`);
      const slideXml = strFromU8(parts[slidePath]);
      parts[slidePath] = strToU8(
        slideXml.includes('<p:timing')
          ? slideXml.replace(/<p:timing\b[\s\S]*?<\/p:timing>/, timing)
          : slideXml.replace('</p:sld>', `${timing}</p:sld>`),
      );
      const source = join(dir, 'source.pptx');
      await writeFile(source, zipSync(parts));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage();
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.waitForFunction(() => state.slides.length === 2);
      const persistedMedia = await page
        .evaluate(async () => (await fetch('/state')).json())
        .then((current) => current.media[0]);
      assert.equal(persistedMedia.playback.autoplay, true);
      assert.equal(persistedMedia.playback.loop, true);
      assert.equal(persistedMedia.playback.slideCount, 999);
      assert.equal(persistedMedia.playback.hideWhenStopped, true);
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      const audio = page.locator('foreignObject[data-pptx-media] audio');
      await audio.waitFor({ state: 'attached' });
      const original = await audio.elementHandle();
      assert.ok(original);
      await page.waitForFunction(() => {
        const root = document.querySelector('#slide')?.shadowRoot;
        const element = root?.querySelector('foreignObject[data-pptx-media] audio');
        return element?.currentTime > 0.05;
      });
      assert.equal(
        await audio.locator('xpath=..').evaluate((element) => getComputedStyle(element).visibility),
        'hidden',
      );
      const initialTime = await audio.evaluate((element) => element.currentTime);
      await page.locator('#stage').focus();
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => index === 1);
      assert.equal(await original.evaluate((element) => element.isConnected), true);
      await page.waitForFunction((time) => {
        const element = document.querySelector('audio');
        return element !== null && element.currentTime > time + 0.05;
      }, initialTime);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test('audio rewinds after natural playback when requested', { timeout: 60000 }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-media-rewind-'));
  let preview;
  let browser;
  try {
    const deck = await compile(
      Presentation({
        children: [
          Slide({
            children: Media({ kind: 'audio', data: wav(1000), x: 1, y: 1, width: 3, height: 1 }),
          }),
        ],
      }),
    );
    const [audio] = getSlideShapes(getSlides(deck)[0]);
    setShapeMediaPlayback(audio, {
      autoplay: false,
      muted: true,
      rewindAfterPlaying: true,
    });
    const source = join(dir, 'source.pptx');
    await writeFile(source, await savePresentation(deck));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
    );
    preview = await startPreview(file);
    browser = await chromium.launch({ headless: true });
    const page = await browser.newPage();
    await page.goto(preview.url);
    await page.getByRole('button', { name: 'Preview', exact: true }).click();
    await page.getByRole('button', { name: 'Present', exact: true }).click();
    const audioElement = page.locator('foreignObject[data-pptx-media] audio');
    await audioElement.waitFor({ state: 'attached' });
    const ended = page.evaluate(
      () =>
        new Promise((resolve) => {
          const element = document.querySelector('#slide')?.shadowRoot?.querySelector('audio');
          if (!(element instanceof HTMLAudioElement)) throw new Error('audio element not found');
          element.addEventListener('ended', resolve, { once: true });
        }),
    );
    await audioElement.evaluate((element) => element.play());
    await page.waitForFunction(() => {
      const element = document.querySelector('#slide')?.shadowRoot?.querySelector('audio');
      return element !== null && element.currentTime > 0.05;
    });
    await ended;
    await page.waitForTimeout(50);
    assert.equal(await audioElement.evaluate((element) => element.paused), true);
    assert.ok((await audioElement.evaluate((element) => element.currentTime)) < 0.05);
  } finally {
    await browser?.close();
    await preview?.close();
    await rm(dir, { recursive: true, force: true });
  }
});

test(
  'audio continues across its configured slides without mixing same-id media',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-media-across-'));
    let preview;
    let browser;
    try {
      const deck = await compile(
        Presentation({
          children: [
            Slide({
              children: Media({ kind: 'audio', data: wav(15000), x: 1, y: 1, width: 3, height: 1 }),
            }),
            Slide({
              children: Media({ kind: 'audio', data: wav(15000), x: 1, y: 1, width: 3, height: 1 }),
            }),
            Slide({ children: Text({ x: 1, y: 1, width: 5, height: 1, children: 'Third' }) }),
          ],
        }),
      );
      const slides = getSlides(deck);
      assert.equal(
        getShapeId(getSlideShapes(slides[0])[0]),
        getShapeId(getSlideShapes(slides[1])[0]),
      );
      setShapeMediaPlayback(getSlideShapes(slides[0])[0], {
        autoplay: true,
        hideWhenStopped: true,
        slideCount: 2,
        muted: true,
      });
      setShapeMediaPlayback(getSlideShapes(slides[1])[0], { autoplay: false, muted: true });
      const source = join(dir, 'source.pptx');
      await writeFile(source, await savePresentation(deck));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage();
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.waitForFunction(() => state.slides.length === 3);
      const popup = page.waitForEvent('popup');
      await page.getByRole('button', { name: 'Presenter view', exact: true }).click();
      const presenter = await popup;
      const first = await page.locator('#slide audio').elementHandle();
      assert.ok(first);
      await page.waitForFunction((element) => element.currentTime > 0.1 && !element.paused, first);
      assert.equal(
        await first.evaluate((element) => getComputedStyle(element).visibility),
        'hidden',
        'Hide During Show keeps background audio invisible while playing',
      );
      const firstTime = await first.evaluate((element) => element.currentTime);
      const firstMediaKey = await page.evaluate(() => mediaPlayerKey);
      await presenter.getByRole('button', { name: 'Next', exact: true }).click();
      await page.waitForFunction(() => index === 1);
      assert.equal(
        await first.evaluate((element) => element.paused),
        false,
        'slide navigation must preserve playing audio',
      );
      await page.waitForFunction(({ element, time }) => element.currentTime > time + 0.1, {
        element: first,
        time: firstTime,
      });
      await presenter.getByText('Slide 2 of 3', { exact: true }).waitFor();
      const second = page.locator('#slide audio');
      await second.waitFor({ state: 'attached' });
      assert.equal(await second.evaluate((element) => element.paused), true);
      await presenter.getByRole('button', { name: 'Play media', exact: true }).click();
      await page.waitForFunction(
        () => document.querySelector('#slide').shadowRoot.querySelector('audio')?.paused === false,
      );
      await presenter.evaluate(
        ({ mediaKey, shapeId }) => {
          window.opener.postMessage(
            {
              type: 'presenter-command',
              action: 'media',
              mediaKey,
              index: { shapeId, action: 'pause' },
            },
            location.origin,
          );
        },
        { mediaKey: firstMediaKey, shapeId: getShapeId(getSlideShapes(slides[1])[0]) },
      );
      await page.waitForTimeout(100);
      assert.equal(
        await second.evaluate((element) => element.paused),
        false,
        'stale presenter commands cannot control another slide',
      );
      await presenter.getByRole('button', { name: 'Pause media', exact: true }).click();
      await page.waitForFunction(
        () => document.querySelector('#slide').shadowRoot.querySelector('audio')?.paused === true,
      );
      assert.equal(
        await first.evaluate((element) => element.paused),
        false,
        'presenter commands target only current slide media',
      );
      await presenter.getByRole('button', { name: 'Next', exact: true }).click();
      await page.waitForFunction(() => index === 2);
      await page.waitForFunction(
        (element) => element.paused && !element.hasAttribute('src'),
        first,
      );
      await presenter.getByRole('button', { name: 'Previous', exact: true }).click();
      await page.waitForFunction(() => index === 1);
      await presenter.getByRole('button', { name: 'Previous', exact: true }).click();
      await page.waitForFunction(() => index === 0);
      const restarted = await page.locator('#slide audio').elementHandle();
      assert.ok(restarted);
      await page.waitForFunction((element) => !element.paused, restarted);
      await presenter.getByRole('button', { name: 'Next', exact: true }).click();
      await page.waitForFunction(() => index === 1);
      await presenter.getByRole('button', { name: 'Exit presentation', exact: true }).click();
      await page.waitForFunction(() => !presenting);
      await page.waitForFunction(
        (element) => element.paused && !element.hasAttribute('src'),
        restarted,
      );
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

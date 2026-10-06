import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  addBlankSlide,
  addSlideMedia,
  cm,
  createPresentation,
  getShapeBounds,
  getShapeImageBytes,
  getShapeImageCrop,
  getShapeMedia,
  getShapeMediaPlayback,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeMediaPlayback,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const makeVideo = (page) =>
  page.evaluate(async () => {
    const canvas = document.createElement('canvas');
    canvas.width = 160;
    canvas.height = 90;
    const context = canvas.getContext('2d');
    const stream = canvas.captureStream(20);
    const recorder = new MediaRecorder(stream, { mimeType: 'video/webm;codecs=vp8' });
    const chunks = [];
    const stopped = new Promise((resolve) => (recorder.onstop = resolve));
    recorder.ondataavailable = (event) => chunks.push(event.data);
    recorder.start();
    let frame = 0;
    const timer = setInterval(() => {
      context.fillStyle = frame++ % 2 ? '#2d9cdb' : '#eb5757';
      context.fillRect(0, 0, canvas.width, canvas.height);
    }, 50);
    await new Promise((resolve) => setTimeout(resolve, 700));
    recorder.stop();
    await stopped;
    clearInterval(timer);
    stream.getTracks().forEach((track) => track.stop());
    return Array.from(new Uint8Array(await new Blob(chunks).arrayBuffer()));
  });

const closeTo = (actual, expected, tolerance = 0.002) =>
  assert.ok(Math.abs(actual - expected) <= tolerance, `${actual} is not close to ${expected}`);

const cmRect = (bounds) => ({
  x: bounds.x / cm(1),
  y: bounds.y / cm(1),
  w: bounds.w / cm(1),
  h: bounds.h / cm(1),
});

const pictureRect = (state) => {
  const crop = state.crop ?? { left: 0, top: 0, right: 0, bottom: 0 };
  const bounds = cmRect(state.bounds);
  const w = bounds.w / (1 - crop.left - crop.right);
  const h = bounds.h / (1 - crop.top - crop.bottom);
  return {
    x: bounds.x - crop.left * w,
    y: bounds.y - crop.top * h,
    w,
    h,
  };
};

test(
  'Video Crop details match PowerPoint position fields and preserve media history',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-video-crop-'));
    let browser;
    let preview;
    try {
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 900 } });
      const media = Uint8Array.from(await makeVideo(page));
      const deck = createPresentation();
      const slide = addBlankSlide(deck);
      const video = addSlideMedia(slide, {
        kind: 'video',
        data: media,
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(2.25),
      });
      setShapeMediaPlayback(video, {
        autoplay: true,
        loop: true,
        volume: 0.73,
        muted: true,
        fullScreen: true,
        hideWhenStopped: true,
      });
      const source = join(dir, 'source.pptx');
      await writeFile(source, await savePresentation(deck));
      const entry = join(dir, 'deck.tsx');
      await writeFile(
        entry,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
      );
      preview = await startPreview(entry);
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const readVideo = async () => {
        const savedDeck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        const shape = getSlideShapes(getSlides(savedDeck)[0])[0];
        const clip = getShapeMedia(shape);
        assert.ok(clip && clip.kind === 'video');
        return {
          bounds: getShapeBounds(shape),
          crop: getShapeImageCrop(shape),
          media: Buffer.from(clip.bytes),
          poster: Buffer.from(getShapeImageBytes(shape) ?? []),
          playback: getShapeMediaPlayback(shape),
        };
      };
      await saved();
      await editor.locator('.hit').first().click();
      await editor.locator('.hit').first().click({ button: 'right' });
      await editor.getByRole('menuitem', { name: 'Format Video...', exact: true }).click();
      await editor.getByRole('tab', { name: 'Video Format', exact: true }).click();
      await editor.getByRole('tab', { name: 'Video', exact: true }).click();
      const detail = editor.getByRole('tabpanel', { name: 'Video', exact: true });
      const crop = detail.locator('section[aria-label="Crop"]');
      await crop.waitFor();
      await page.screenshot({ path: '/tmp/pptx-video-crop-pane.png', fullPage: true });
      const field = (name) => crop.getByRole('spinbutton', { name, exact: true });
      for (const name of [
        'Picture width',
        'Picture height',
        'Offset X',
        'Offset Y',
        'Crop width',
        'Crop height',
        'Crop left',
        'Crop top',
      ]) {
        await field(name).waitFor();
      }
      const update = async (name, value) => {
        await field(name).fill(String(value));
        await field(name).press('Tab');
        await saved();
      };

      const initial = await readVideo();
      assert.equal(initial.crop, null);
      closeTo(cmRect(initial.bounds).x, 2.54);
      closeTo(cmRect(initial.bounds).w, 10.16);
      const originalMedia = initial.media;
      const originalPoster = initial.poster;
      const originalPlayback = initial.playback;

      await update('Picture width', 12);
      let state = await readVideo();
      closeTo(cmRect(state.bounds).w, 10.16);
      closeTo(state.crop.left, 0.07667, 0.00002);
      closeTo(state.crop.right, 0.07667, 0.00002);
      assert.deepEqual(state.media, originalMedia);

      // Exercise the remaining PowerPoint position fields, including a negative offset.
      await update('Picture height', 6);
      await update('Offset Y', -0.5);
      state = await readVideo();
      closeTo(state.crop.top, 0.10708, 0.00002);
      closeTo(state.crop.bottom, -0.05958, 0.00002);
      await update('Offset Y', 0.5);
      state = await readVideo();
      closeTo(state.crop.top, -0.05958, 0.00002);
      closeTo(state.crop.bottom, 0.10708, 0.00002);

      const beforeCropHeight = state;
      const beforeCropHeightPicture = pictureRect(beforeCropHeight);
      await update('Crop height', 4);
      state = await readVideo();
      closeTo(cmRect(state.bounds).h, 4, 0.01);
      for (const key of ['x', 'y', 'w', 'h'])
        closeTo(pictureRect(state)[key], beforeCropHeightPicture[key], 0.02);
      const beforeCropTopPicture = pictureRect(state);
      await update('Crop top', 3);
      state = await readVideo();
      closeTo(cmRect(state.bounds).y, 3, 0.01);
      for (const key of ['x', 'y', 'w', 'h'])
        closeTo(pictureRect(state)[key], beforeCropTopPicture[key], 0.02);

      await update('Offset X', 1);
      state = await readVideo();
      closeTo(state.crop.left, -0.00667, 0.00002);
      closeTo(state.crop.right, 0.16001, 0.00002);

      await update('Crop width', 8);
      state = await readVideo();
      closeTo(cmRect(state.bounds).w, 8, 0.01);
      closeTo(state.crop.left, -0.00667, 0.00002);
      closeTo(state.crop.right, 0.34001, 0.00002);

      const beforeLeft = state;
      const beforePicture = pictureRect(beforeLeft);
      await update('Crop left', 3);
      state = await readVideo();
      closeTo(cmRect(state.bounds).x, 3, 0.01);
      closeTo(state.crop.left, 0.03166, 0.00002);
      closeTo(state.crop.right, 0.30167, 0.00002);
      const afterPicture = pictureRect(state);
      for (const key of ['x', 'y', 'w', 'h']) closeTo(afterPicture[key], beforePicture[key], 0.02);
      assert.deepEqual(state.media, originalMedia);
      assert.deepEqual(state.poster, originalPoster);
      assert.deepEqual(state.playback, originalPlayback);

      const reset = crop.getByRole('button', { name: 'Reset', exact: true });
      await reset.click();
      await saved();
      const resetState = await readVideo();
      assert.equal(resetState.crop, null);
      closeTo(cmRect(resetState.bounds).x, beforePicture.x, 0.02);
      closeTo(cmRect(resetState.bounds).w, beforePicture.w, 0.02);
      closeTo(cmRect(resetState.bounds).y, beforePicture.y, 0.02);
      closeTo(cmRect(resetState.bounds).h, beforePicture.h, 0.02);
      assert.deepEqual(resetState.media, originalMedia);
      assert.deepEqual(resetState.poster, originalPoster);
      assert.deepEqual(resetState.playback, originalPlayback);

      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      const undone = await readVideo();
      assert.deepEqual(undone.crop, state.crop);
      assert.deepEqual(undone.bounds, state.bounds);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.equal((await readVideo()).crop, null);
      await page.reload();
      await saved();
      const persisted = await readVideo();
      assert.equal(persisted.crop, null);
      assert.deepEqual(persisted.media, originalMedia);
      assert.deepEqual(persisted.poster, originalPoster);
      assert.deepEqual(persisted.playback, originalPlayback);

      // A crop source rectangle at the ST_Percentage boundary must be rejected
      // without changing the document or leaving the invalid input visible.
      await editor.locator('.hit').first().click();
      await editor.locator('.hit').first().click({ button: 'right' });
      await editor.getByRole('menuitem', { name: 'Format Video...', exact: true }).click();
      await editor.getByRole('tab', { name: 'Video Format', exact: true }).click();
      await editor.getByRole('tab', { name: 'Video', exact: true }).click();
      const reloadedCrop = editor
        .getByRole('tabpanel', { name: 'Video', exact: true })
        .locator('section[aria-label="Crop"]');
      await reloadedCrop.waitFor();
      const reloadedField = (name) => reloadedCrop.getByRole('spinbutton', { name, exact: true });
      await reloadedField('Picture width').fill('0.01');
      await reloadedField('Picture width').press('Tab');
      await saved();
      const beforeInvalid = await readVideo();
      const previousOffset = await reloadedField('Offset X').inputValue();
      await reloadedField('Offset X').fill('5963.92');
      await reloadedField('Offset X').press('Tab');
      await editor
        .getByText('These crop dimensions exceed the supported range.', { exact: true })
        .waitFor();
      assert.equal(await reloadedField('Offset X').inputValue(), previousOffset);
      assert.deepEqual(await readVideo(), beforeInvalid);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

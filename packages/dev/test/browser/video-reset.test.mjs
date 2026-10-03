import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { chromium } from 'playwright';
import {
  addBlankSlide,
  addSlideMedia,
  createPresentation,
  getShapeBounds,
  getShapeFill,
  getShapeImageBrightness,
  getShapeImageContrast,
  getShapeImageBytes,
  getShapeImageDuotone,
  getShapeImageCrop,
  getShapeMedia,
  getShapeMediaPlayback,
  getShapePreset,
  getShapeStroke,
  getShapeFillOpacity,
  getShapeStrokeOpacity,
  getSlideShapes,
  getSlides,
  isShapeImageGrayscale,
  inches,
  loadPresentation,
  savePresentation,
  setShapeFill,
  setShapeImageBrightness,
  setShapeImageContrast,
  setShapeMediaPlayback,
  setShapeStroke,
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

test(
  'Video Format Reset clears corrections as one undoable action and preserves the clip',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-video-reset-'));
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
        w: inches(6),
        h: inches(3.375),
      });
      setShapeFill(video, { color: '#123456', opacity: 0.62 });
      setShapeStroke(video, { color: '#FEDCBA', widthEmu: 38100, opacity: 0.58 });
      const source = join(dir, 'source.pptx');
      const sourceParts = unzipSync(await savePresentation(deck));
      const slidePath = 'ppt/slides/slide1.xml';
      const slideXml = strFromU8(sourceParts[slidePath]);
      const colorEffects =
        '<a:grayscl/><a:duotone><a:srgbClr val="FF0000"/><a:srgbClr val="0000FF"/></a:duotone>';
      sourceParts[slidePath] = strToU8(
        slideXml.replace(/<a:blip([^>]*)\/>/, `<a:blip$1>${colorEffects}</a:blip>`),
      );
      await writeFile(source, zipSync(sourceParts));
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
          brightness: getShapeImageBrightness(shape) ?? 0,
          contrast: getShapeImageContrast(shape) ?? 0,
          bounds: getShapeBounds(shape),
          fill: getShapeFill(shape),
          stroke: getShapeStroke(shape),
          fillOpacity: getShapeFillOpacity(shape),
          strokeOpacity: getShapeStrokeOpacity(shape),
          media: Buffer.from(clip.bytes),
          poster: Buffer.from(getShapeImageBytes(shape) ?? []),
          grayscale: isShapeImageGrayscale(shape),
          duotone: getShapeImageDuotone(savedDeck, shape),
        };
      };

      await saved();
      await editor.locator('.hit').first().click();
      await editor.getByRole('tab', { name: 'Video Format', exact: true }).click();
      const format = editor.getByRole('tablist', { name: 'Format Video', exact: true });
      await format.getByRole('tab', { name: 'Video', exact: true }).click();
      const detail = editor.getByRole('tabpanel', { name: 'Video', exact: true });
      const brightness = detail.getByRole('spinbutton', { name: 'Brightness', exact: true });
      const contrast = detail.getByRole('spinbutton', { name: 'Contrast', exact: true });
      await brightness.fill('30');
      await brightness.press('Tab');
      await saved();
      await contrast.fill('-20');
      await contrast.press('Tab');
      await saved();
      const before = await readVideo();
      assert.deepEqual([before.brightness, before.contrast], [0.3, -0.2]);
      assert.equal(before.grayscale, true);
      assert.deepEqual(before.duotone, { firstColor: '#FF0000', secondColor: '#0000FF' });
      assert.equal(before.fillOpacity, 0.62);
      assert.equal(before.strokeOpacity, 0.58);

      const reset = detail
        .getByRole('region', { name: 'Video options', exact: true })
        .getByRole('button', { name: 'Reset', exact: true });
      await reset.click();
      await saved();
      const cleared = await readVideo();
      assert.deepEqual([cleared.brightness, cleared.contrast], [0, 0]);
      assert.equal(cleared.grayscale, false);
      assert.equal(cleared.duotone, null);
      assert.deepEqual(cleared.bounds, before.bounds);
      assert.deepEqual(cleared.fill, before.fill);
      assert.deepEqual(cleared.stroke, before.stroke);
      assert.deepEqual(cleared.media, before.media);
      assert.deepEqual(cleared.poster, before.poster);
      assert.equal(cleared.fillOpacity, before.fillOpacity);
      assert.equal(cleared.strokeOpacity, before.strokeOpacity);

      // Reset is one PowerPoint history action: Undo restores both corrections.
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      const undone = await readVideo();
      assert.deepEqual([undone.brightness, undone.contrast], [0.3, -0.2]);
      assert.equal(undone.grayscale, true);
      assert.deepEqual(undone.duotone, { firstColor: '#FF0000', secondColor: '#0000FF' });
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      const redone = await readVideo();
      assert.deepEqual([redone.brightness, redone.contrast], [0, 0]);
      assert.equal(redone.grayscale, false);
      assert.equal(redone.duotone, null);

      await page.reload();
      await saved();
      await editor.locator('.hit').first().click();
      await editor.getByRole('tab', { name: 'Video Format', exact: true }).click();
      await editor.getByRole('tab', { name: 'Video', exact: true }).click();
      // Reset remains available after corrections have returned to neutral.
      assert.equal(
        await editor
          .getByRole('tabpanel', { name: 'Video', exact: true })
          .getByRole('region', { name: 'Video options', exact: true })
          .getByRole('button', { name: 'Reset', exact: true })
          .isDisabled(),
        false,
      );
      const persisted = await readVideo();
      assert.deepEqual([persisted.brightness, persisted.contrast], [0, 0]);
      assert.equal(persisted.grayscale, false);
      assert.equal(persisted.duotone, null);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'Video Format ribbon Reset clears shape formatting while preserving crop and media',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-video-ribbon-reset-'));
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
        w: inches(6),
        h: inches(3.375),
      });
      setShapeFill(video, { color: '#123456', opacity: 0.62 });
      setShapeStroke(video, { color: '#FEDCBA', widthEmu: 38100, opacity: 0.58 });
      setShapeImageBrightness(video, 0.3);
      setShapeImageContrast(video, -0.2);
      setShapeMediaPlayback(video, {
        autoplay: true,
        loop: true,
        volume: 0.73,
        muted: true,
        fullScreen: true,
        hideWhenStopped: true,
      });
      const source = join(dir, 'source.pptx');
      const sourceParts = unzipSync(await savePresentation(deck));
      const slidePath = 'ppt/slides/slide1.xml';
      const originalXml = strFromU8(sourceParts[slidePath]);
      const slideXml = originalXml
        .replace('<a:prstGeom prst="rect"', '<a:prstGeom prst="ellipse"')
        .replace(
          /(<a:blip[^>]*>)/,
          '$1<a:grayscl/><a:duotone><a:srgbClr val="FF0000"/><a:srgbClr val="0000FF"/></a:duotone>',
        )
        .replace(
          /(<p:blipFill>[\s\S]*?<\/a:blip>)/,
          '$1<a:srcRect l="10000" t="5000" r="15000" b="20000"/>',
        )
        .replace(
          /(<p:spPr>[\s\S]*?<\/a:ln>)/,
          '$1<a:effectLst><a:outerShdw blurRad="63500" dist="127000" dir="5400000"><a:srgbClr val="000000"/></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="0"/></a:lightRig></a:scene3d>',
        );
      assert.notEqual(slideXml, originalXml);
      sourceParts[slidePath] = strToU8(slideXml);
      await writeFile(source, zipSync(sourceParts));
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
        const bytes = new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer());
        const parts = unzipSync(bytes);
        const xml = strFromU8(parts[slidePath]);
        const savedDeck = await loadPresentation(bytes);
        const shape = getSlideShapes(getSlides(savedDeck)[0])[0];
        const clip = getShapeMedia(shape);
        assert.ok(clip && clip.kind === 'video');
        return {
          bounds: getShapeBounds(shape),
          crop: getShapeImageCrop(shape),
          brightness: getShapeImageBrightness(shape) ?? 0,
          contrast: getShapeImageContrast(shape) ?? 0,
          grayscale: isShapeImageGrayscale(shape),
          duotone: getShapeImageDuotone(savedDeck, shape),
          playback: getShapeMediaPlayback(shape),
          fill: getShapeFill(shape),
          stroke: getShapeStroke(shape),
          media: Buffer.from(clip.bytes),
          poster: Buffer.from(getShapeImageBytes(shape) ?? []),
          preset: getShapePreset(shape),
          hasEffects: xml.includes('<a:effectLst>'),
          hasScene3d: xml.includes('<a:scene3d>'),
        };
      };

      await saved();
      await editor.locator('.hit').first().click();
      await editor.getByRole('tab', { name: 'Video Format', exact: true }).click();
      const panel = editor.locator('#ribbon-panel');
      const reset = panel.getByRole('button', { name: 'Reset', exact: true });
      assert.equal(await reset.count(), 1);

      const before = await readVideo();
      assert.equal(before.preset, 'ellipse');
      assert.deepEqual([before.brightness, before.contrast], [0.3, -0.2]);
      assert.equal(before.grayscale, true);
      assert.deepEqual(before.duotone, { firstColor: '#FF0000', secondColor: '#0000FF' });
      assert.deepEqual(before.playback, {
        autoplay: true,
        loop: true,
        volume: 0.73,
        muted: true,
        fullScreen: true,
        hideWhenStopped: true,
      });
      assert.deepEqual(before.crop, { left: 0.1, top: 0.05, right: 0.15, bottom: 0.2 });
      assert.equal(before.hasEffects, true);
      assert.equal(before.hasScene3d, true);
      await reset.click();
      await saved();
      const cleared = await readVideo();
      assert.equal(cleared.preset, 'rect');
      assert.deepEqual([cleared.brightness, cleared.contrast], [0, 0]);
      assert.equal(cleared.grayscale, false);
      assert.equal(cleared.duotone, null);
      assert.deepEqual(cleared.playback, before.playback);
      assert.deepEqual(cleared.crop, before.crop);
      assert.deepEqual(cleared.fill, { kind: 'inherit' });
      assert.deepEqual(cleared.stroke, { kind: 'inherit' });
      assert.equal(cleared.hasEffects, false);
      assert.equal(cleared.hasScene3d, false);
      assert.deepEqual(cleared.bounds, before.bounds);
      assert.deepEqual(cleared.media, before.media);
      assert.deepEqual(cleared.poster, before.poster);
      await page.screenshot({ path: '/tmp/pptx-video-ribbon-reset.png' });

      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      const undone = await readVideo();
      assert.equal(undone.preset, 'ellipse');
      assert.deepEqual([undone.brightness, undone.contrast], [0.3, -0.2]);
      assert.equal(undone.grayscale, true);
      assert.deepEqual(undone.duotone, { firstColor: '#FF0000', secondColor: '#0000FF' });
      assert.deepEqual(undone.playback, before.playback);
      assert.deepEqual(undone.crop, before.crop);
      assert.deepEqual(undone.fill, before.fill);
      assert.deepEqual(undone.stroke, before.stroke);
      assert.equal(undone.hasEffects, true);
      assert.equal(undone.hasScene3d, true);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      const redone = await readVideo();
      assert.equal(redone.preset, 'rect');
      assert.deepEqual([redone.brightness, redone.contrast], [0, 0]);
      assert.equal(redone.grayscale, false);
      assert.equal(redone.duotone, null);
      assert.deepEqual(redone.playback, before.playback);
      assert.equal(redone.hasEffects, false);
      assert.equal(redone.hasScene3d, false);
      assert.deepEqual(redone.media, before.media);
      assert.deepEqual(redone.poster, before.poster);

      await page.reload();
      await saved();
      const persisted = await readVideo();
      assert.equal(persisted.preset, 'rect');
      assert.deepEqual([persisted.brightness, persisted.contrast], [0, 0]);
      assert.equal(persisted.grayscale, false);
      assert.equal(persisted.duotone, null);
      assert.deepEqual(persisted.playback, before.playback);
      assert.deepEqual(persisted.crop, before.crop);
      assert.equal(persisted.hasEffects, false);
      assert.equal(persisted.hasScene3d, false);
      assert.deepEqual(persisted.media, before.media);
      assert.deepEqual(persisted.poster, before.poster);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

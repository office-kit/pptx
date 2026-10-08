import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  addBlankSlide,
  addSlideMedia,
  addSlideTextBox,
  createPresentation,
  getShapeImageBrightness,
  getShapeImageContrast,
  getSlideShapes,
  getSlides,
  loadPresentation,
  inches,
  savePresentation,
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
  'Video Format Corrections applies 25 presets, saves brightness and contrast, and undoes atomically',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-video-corrections-'));
    let browser;
    let preview;
    try {
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 900 } });
      const media = Uint8Array.from(await makeVideo(page));
      const deck = createPresentation();
      const slide = addBlankSlide(deck);
      addSlideMedia(slide, {
        kind: 'video',
        data: media,
        x: inches(1),
        y: inches(1),
        w: inches(6),
        h: inches(3.375),
      });
      addSlideTextBox(slide, {
        x: inches(8),
        y: inches(1),
        w: inches(2),
        h: inches(1),
        text: 'Selection test',
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
      await saved();
      await editor.locator('.hit').first().click();
      await editor.getByRole('tab', { name: 'Video Format', exact: true }).click();
      const panel = editor.locator('#ribbon-panel');
      const correctionsButton = panel.getByRole('button', { name: 'Corrections', exact: true });
      await correctionsButton.click();
      const corrections = editor.getByRole('menu', { name: 'Corrections', exact: true });
      await corrections.waitFor();
      assert.equal(await corrections.getByRole('menuitemradio').count(), 25);
      assert.equal(
        await corrections.getByRole('menuitemradio').nth(1).getAttribute('aria-label'),
        'Brightness: -20% Contrast: -40%',
      );
      await page.screenshot({ path: '/tmp/pptx-video-corrections-gallery.png' });

      const readAdjustments = async () => {
        const savedDeck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        const shape = getSlideShapes(getSlides(savedDeck)[0])[0];
        return [getShapeImageBrightness(shape) ?? 0, getShapeImageContrast(shape) ?? 0];
      };
      assert.deepEqual(await readAdjustments(), [0, 0]);

      const plus20 = corrections.getByRole('menuitemradio', {
        name: 'Brightness: +20% Contrast: +20%',
        exact: true,
      });
      assert.equal(await plus20.count(), 1);
      await plus20.click();
      await saved();
      assert.deepEqual(await readAdjustments(), [0.2, 0.2]);

      // The live canvas preview must use the same DrawingML transfer function
      // as the exported picture, including the reference desktop app's brightness/contrast order.
      const video = editor.locator('.media-preview video');
      const mediaPreview = editor.locator('.media-preview:has(video)');
      assert.match(await video.evaluate((node) => getComputedStyle(node).filter), /url\(/);
      const transfer = mediaPreview.locator('feComponentTransfer feFuncR');
      await transfer.first().waitFor({ state: 'attached' });
      const transferValues = await transfer.evaluateAll((nodes) =>
        nodes.map((node) => ({
          slope: node.getAttribute('slope'),
          intercept: node.getAttribute('intercept'),
        })),
      );
      assert.ok(
        transferValues.some(
          ({ slope, intercept }) =>
            Math.abs(Number(slope) - 1.2475633528) < 1e-9 &&
            Math.abs(Number(intercept) - 0.1004892405) < 1e-9,
        ),
        `missing +20/+20 transfer function: ${JSON.stringify(transferValues)}`,
      );

      // A gallery choice is one history action in the reference desktop app: one Undo removes
      // both DrawingML lum transforms together.
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await readAdjustments(), [0, 0]);

      // Undo can release the canvas selection. Re-select before reopening the
      // shape-specific Video Format controls.
      await editor.locator('.hit').first().click();
      await editor.getByRole('tab', { name: 'Video Format', exact: true }).click();
      await correctionsButton.click();
      const reopened = editor.getByRole('menu', { name: 'Corrections', exact: true });
      await reopened
        .getByRole('menuitemradio', { name: 'Brightness: +20% Contrast: +20%', exact: true })
        .click();
      await saved();
      await correctionsButton.click();
      const normal = reopened.getByRole('menuitemradio', {
        name: 'Brightness: 0% (Normal) Contrast: 0% (Normal)',
        exact: true,
      });
      assert.equal(await normal.count(), 1);
      await reopened.getByRole('menuitemradio').nth(24).focus();
      await reopened.getByRole('menuitemradio').nth(24).press('ArrowUp');
      assert.equal(
        await reopened
          .getByRole('menuitemradio')
          .nth(19)
          .evaluate((node) => node === node.ownerDocument.activeElement),
        true,
      );
      await normal.click();
      await saved();
      assert.deepEqual(await readAdjustments(), [0, 0]);

      await correctionsButton.click();
      await reopened.press('Escape');
      await reopened.waitFor({ state: 'hidden' });
      assert.equal(
        await correctionsButton.evaluate((node) => node === node.ownerDocument.activeElement),
        true,
      );

      await correctionsButton.click();
      await reopened
        .getByRole('menuitem', { name: 'Movie Correction Options...', exact: true })
        .click();
      const format = editor.getByRole('tablist', { name: 'Format Video', exact: true });
      await format.waitFor();
      assert.equal(
        await format.getByRole('tab', { name: 'Video', exact: true }).getAttribute('aria-selected'),
        'true',
      );
      const detail = editor.getByRole('tabpanel', { name: 'Video', exact: true });
      const brightness = detail.getByRole('spinbutton', { name: 'Brightness', exact: true });
      await detail.getByRole('button', { name: 'Corrections presets', exact: true }).click();
      assert.equal(
        await editor
          .getByRole('menu', { name: 'Corrections', exact: true })
          .getByRole('menuitemradio')
          .count(),
        25,
      );
      await editor.getByRole('menu', { name: 'Corrections', exact: true }).press('Escape');
      await brightness.fill('101');
      await brightness.press('Tab');
      assert.equal(await brightness.evaluate((input) => input.validity.rangeOverflow), true);
      assert.deepEqual(await readAdjustments(), [0, 0]);

      // Inspect the preview before release so onchange-only rendering cannot pass.
      const brightnessSlider = detail.getByRole('slider', { name: 'Brightness', exact: true });
      const sliderBounds = await brightnessSlider.boundingBox();
      assert.ok(sliderBounds, 'brightness slider must be measurable');
      const sliderStart = sliderBounds.x + sliderBounds.width / 2;
      const sliderTarget = sliderBounds.x + sliderBounds.width * 0.675;
      const liveTransfer = editor
        .locator('.media-preview:has(video) feComponentTransfer feFuncR')
        .first();
      await page.mouse.move(sliderStart, sliderBounds.y + sliderBounds.height / 2);
      await page.mouse.down();
      await page.mouse.move(sliderTarget, sliderBounds.y + sliderBounds.height / 2, { steps: 8 });
      const liveValue = Number(await brightnessSlider.inputValue());
      assert.ok(liveValue >= 30, `brightness should update during drag, got ${liveValue}`);
      await liveTransfer.waitFor({ state: 'attached' });
      const liveIntercept = Number(await liveTransfer.getAttribute('intercept'));
      assert.ok(
        liveIntercept > 0.2,
        `live transfer should update during drag, got ${liveIntercept}`,
      );
      await page.mouse.up();
      await saved();
      const draggedAdjustments = await readAdjustments();
      assert.ok(
        draggedAdjustments[0] >= 0.3,
        `unexpected dragged brightness: ${draggedAdjustments[0]}`,
      );
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await readAdjustments(), [0, 0]);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual(await readAdjustments(), draggedAdjustments);

      await brightness.fill('35');
      await brightness.press('Tab');
      await saved();
      assert.deepEqual(await readAdjustments(), [0.35, 0]);
      assert.equal(
        await detail.getByRole('slider', { name: 'Brightness', exact: true }).inputValue(),
        '35',
      );
      const contrast = detail.getByRole('slider', { name: 'Contrast', exact: true });
      await contrast.focus();
      await contrast.press('ArrowRight');
      await saved();
      assert.deepEqual(await readAdjustments(), [0.35, 0.01]);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await readAdjustments(), [0.35, 0]);
      await editor.locator('.hit').first().click();
      await editor.getByRole('tab', { name: 'Video Format', exact: true }).click();

      await format.getByRole('tab', { name: 'Video', exact: true }).click();
      await format.getByRole('tab', { name: 'Video', exact: true }).press('Home');
      assert.equal(
        await format
          .getByRole('tab', { name: 'Fill & Line', exact: true })
          .getAttribute('aria-selected'),
        'true',
      );
      await format.getByRole('tab', { name: 'Fill & Line', exact: true }).press('End');
      assert.equal(
        await format.getByRole('tab', { name: 'Video', exact: true }).getAttribute('aria-selected'),
        'true',
      );
      const detailBounds = await detail.boundingBox();
      for (const field of [
        brightness,
        contrast,
        detail.getByRole('button', { name: 'Corrections presets', exact: true }),
      ]) {
        const bounds = await field.boundingBox();
        assert.ok(
          bounds.x >= detailBounds.x &&
            bounds.x + bounds.width <= detailBounds.x + detailBounds.width,
          'video controls must fit the pane without horizontal clipping',
        );
      }
      await page.screenshot({ path: '/tmp/pptx-video-format-pane.png' });
      await editor.locator('.hit').nth(1).click();
      const shapeFormat = editor.getByRole('tablist', { name: 'Format Shape', exact: true });
      await shapeFormat.waitFor();
      assert.equal(await shapeFormat.getByRole('tab', { selected: true }).count(), 1);
      assert.equal(await shapeFormat.getByRole('tab', { name: 'Video', exact: true }).count(), 0);
      await editor.locator('.hit').first().click();
      await editor.getByRole('tab', { name: 'Video Format', exact: true }).click();

      // The reference desktop app exposes the same gallery after switching UI language.
      await editor.locator('.lang select').selectOption('ja');
      const jaCorrectionsButton = editor
        .locator('#ribbon-panel')
        .getByRole('button', { name: '修整', exact: true });
      await jaCorrectionsButton.click();
      const jaCorrections = editor.getByRole('menu', { name: '修整', exact: true });
      await jaCorrections.waitFor();
      assert.equal(await jaCorrections.getByRole('menuitemradio').count(), 25);
      assert.equal(
        await jaCorrections
          .getByRole('menuitemradio', { name: /明るさ: \+20%.*コントラスト: \+20%/ })
          .count(),
        1,
      );
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

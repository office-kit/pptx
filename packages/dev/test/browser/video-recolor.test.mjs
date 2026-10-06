import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { strFromU8, unzipSync } from 'fflate';
import { chromium } from 'playwright';
import {
  addBlankSlide,
  addSlideMedia,
  createPresentation,
  getShapeImageBrightness,
  getShapeImageContrast,
  getShapeImageDuotone,
  getShapeMedia,
  getShapeStroke,
  getSlideShapes,
  getSlides,
  isShapeImageGrayscale,
  inches,
  loadPresentation,
  savePresentation,
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
  'Video recolor gallery saves native color transforms and supports Undo and reload',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-video-recolor-'));
    let browser;
    let preview;
    try {
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 900 } });
      page.on('pageerror', (error) => console.error(error));
      const media = Uint8Array.from(await makeVideo(page));
      const deck = createPresentation();
      const video = addSlideMedia(addBlankSlide(deck), {
        kind: 'video',
        data: media,
        x: inches(1),
        y: inches(1),
        w: inches(6),
        h: inches(3.375),
      });
      setShapeStroke(video, { color: '#FEDCBA', widthEmu: 38100 });
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
      const read = async () => {
        const bytes = new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer());
        const pres = await loadPresentation(bytes);
        const shape = getSlideShapes(getSlides(pres)[0])[0];
        return { xml: strFromU8(unzipSync(bytes)['ppt/slides/slide1.xml']), shape, pres };
      };
      await saved();
      await editor.locator('.hit').first().click();
      await editor.getByRole('tab', { name: 'Video Format', exact: true }).click();
      const color = editor
        .locator('#ribbon-panel')
        .getByRole('button', { name: 'Color', exact: true });
      await color.click({ timeout: 3000 });
      const gallery = editor.getByRole('menu', { name: 'Recolor', exact: true });
      assert.equal(await gallery.getByRole('menuitemradio').count(), 21);
      const thresholdSamples = await Promise.all(
        [4, 5, 6].map((index) =>
          gallery.getByRole('menuitemradio').nth(index).locator('.sample').screenshot(),
        ),
      );
      assert.notDeepEqual(thresholdSamples[0], thresholdSamples[1]);
      assert.notDeepEqual(thresholdSamples[1], thresholdSamples[2]);
      await gallery.getByRole('menuitemradio', { name: 'Grayscale', exact: true }).click();
      await saved();
      assert.equal(isShapeImageGrayscale((await read()).shape), true);
      assert.equal(await editor.locator('.media-preview feColorMatrix').count(), 1);
      assert.match(
        await editor.locator('.media-preview video').getAttribute('style'),
        /filter: url/,
      );
      await color.click();
      await gallery.getByRole('menuitemradio', { name: 'Washout', exact: true }).click();
      await saved();
      let current = await read();
      assert.equal(isShapeImageGrayscale(current.shape), false);
      assert.equal(getShapeImageBrightness(current.shape), 0.7);
      assert.equal(getShapeImageContrast(current.shape), -0.7);
      const washout = await editor
        .locator('.media-preview feComponentTransfer')
        .first()
        .evaluate((node) => {
          const channel = node.querySelector('feFuncR');
          const slope = Number(channel.getAttribute('slope'));
          const intercept = Number(channel.getAttribute('intercept'));
          return {
            dark: Math.round(38 * slope + 255 * intercept),
            space: getComputedStyle(node).colorInterpolationFilters,
          };
        });
      assert.equal(
        washout.dark,
        217,
        'Washout must retain the dark poster detail seen in PowerPoint',
      );
      assert.equal(washout.space.toLowerCase(), 'srgb');
      await color.click();
      assert.equal(
        await gallery
          .getByRole('menuitemradio', { name: 'Washout', exact: true })
          .getAttribute('aria-checked'),
        'true',
      );
      await gallery.getByRole('menuitemradio', { name: 'Sepia', exact: true }).click();
      await saved();
      current = await read();
      assert.match(
        current.xml,
        /<a:srgbClr val="D9C3A5"><a:tint val="50000"\/><a:satMod val="180000"\/><\/a:srgbClr>/,
      );
      assert.equal(getShapeImageBrightness(current.shape) ?? 0, 0);
      await color.click();
      assert.equal(
        await gallery
          .getByRole('menuitemradio', { name: 'Sepia', exact: true })
          .getAttribute('aria-checked'),
        'true',
      );
      await page.keyboard.press('Escape');
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      current = await read();
      assert.equal(getShapeImageBrightness(current.shape), 0.7);
      assert.equal(getShapeImageContrast(current.shape), -0.7);
      assert.equal(getShapeImageDuotone(current.pres, current.shape), null);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      await color.click();
      await gallery
        .getByRole('menuitemradio', { name: 'Black and White: 50%', exact: true })
        .click();
      await saved();
      assert.match((await read()).xml, /<a:biLevel thresh="50000"\/>/);
      assert.equal(await editor.locator('.media-preview feFuncR[type="discrete"]').count(), 1);
      await color.click();
      await gallery.getByRole('menuitemradio').nth(8).click();
      await saved();
      current = await read();
      assert.match(
        current.xml,
        /<a:schemeClr val="accent1"><a:tint val="45000"\/><a:satMod val="400000"\/><\/a:schemeClr>/,
      );
      assert.doesNotMatch(current.xml, /<a:biLevel/);
      assert.deepEqual(Buffer.from(getShapeMedia(current.shape).bytes), Buffer.from(media));
      assert.deepEqual(getShapeStroke(current.shape), getShapeStroke(video));
      await page.reload();
      await saved();
      await editor.locator('.hit').first().click();
      await editor.getByRole('tab', { name: 'Video Format', exact: true }).click();
      await color.click();
      assert.equal(
        await gallery.getByRole('menuitemradio').nth(8).getAttribute('aria-checked'),
        'true',
      );
      await gallery.getByRole('menuitemradio').first().focus();
      await page.keyboard.press('ArrowDown');
      assert.equal(
        await gallery
          .getByRole('menuitemradio')
          .nth(7)
          .evaluate((el) => el === document.activeElement),
        true,
      );
      await page.keyboard.press('Escape');
      assert.equal(await color.getAttribute('aria-expanded'), 'false');
      await color.click();
      await gallery.getByRole('menuitemradio').nth(15).click();
      await saved();
      assert.match(
        (await read()).xml,
        /<a:schemeClr val="accent1"><a:shade val="45000"\/><a:satMod val="135000"\/><\/a:schemeClr>/,
      );
      await color.click();
      assert.equal(
        await gallery.getByRole('menuitemradio').nth(15).getAttribute('aria-checked'),
        'true',
      );
      await gallery.getByRole('menuitemradio', { name: 'No Recolor', exact: true }).click();
      await saved();
      current = await read();
      assert.equal(getShapeImageDuotone(current.pres, current.shape), null);
      await color.click();
      await gallery.getByRole('menuitem', { name: 'Movie Color Options...', exact: true }).click();
      await editor.getByRole('tabpanel', { name: 'Video', exact: true }).waitFor();
      const pane = editor.getByRole('tabpanel', { name: 'Video', exact: true });
      await pane.getByRole('button', { name: 'Recolor presets', exact: true }).click();
      await gallery.getByRole('button', { name: 'More Variations...', exact: true }).click();
      const variations = editor.getByRole('menu', { name: 'More Variations...', exact: true });
      await variations.getByRole('menuitemradio', { name: 'Accent 2', exact: true }).click();
      await saved();
      assert.match(
        (await read()).xml,
        /<a:schemeClr val="accent2"><a:tint val="45000"\/><a:satMod val="400000"\/><\/a:schemeClr>/,
      );
      await pane.getByRole('button', { name: 'Recolor presets', exact: true }).click();
      assert.equal(
        await gallery.getByRole('menuitemradio').nth(9).getAttribute('aria-checked'),
        'true',
      );
      await gallery.getByRole('button', { name: 'More Variations...', exact: true }).click();
      assert.equal(await variations.getByRole('menuitemradio').count(), 70);
      await variations
        .getByRole('menuitemradio', { name: 'Accent 1, Lighter 80%', exact: true })
        .click();
      await saved();
      assert.match(
        (await read()).xml,
        /<a:schemeClr val="accent1"><a:lumMod val="20000"\/><a:lumOff val="80000"\/><a:tint val="45000"\/><a:satMod val="400000"\/><\/a:schemeClr>/,
      );
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.match(
        (await read()).xml,
        /<a:schemeClr val="accent2"><a:tint val="45000"\/><a:satMod val="400000"\/><\/a:schemeClr>/,
      );
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      assert.match((await read()).xml, /<a:lumMod val="20000"\/><a:lumOff val="80000"\/>/);
      await editor.locator('.hit').first().click();
      await editor.getByRole('tab', { name: 'Video Format', exact: true }).click();
      await color.click();
      await gallery.getByRole('button', { name: 'More Variations...', exact: true }).click();
      await variations
        .getByRole('menuitemradio', { name: 'Accent 1, Darker 25%', exact: true })
        .click();
      await saved();
      assert.match(
        (await read()).xml,
        /<a:schemeClr val="accent1"><a:lumMod val="75000"\/><a:tint val="45000"\/><a:satMod val="400000"\/><\/a:schemeClr>/,
      );
      await color.click();
      await gallery.getByRole('button', { name: 'More Variations...', exact: true }).click();
      await page.screenshot({ path: '/tmp/pptx-video-recolor-pane.png' });
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

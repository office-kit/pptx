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
  getShapeMediaPlayback,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
  savePresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';
const wav = (durationMs = 5033) => {
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
      Math.round(
        Math.sin((index * 2 * Math.PI * 440) / sampleRate) * (index < samples / 2 ? 16000 : 4000),
      ),
      true,
    );
  }
  return bytes;
};

test(
  'trim dialog previews, cancels, saves and undoes the trimmed interval',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-trim-editor-'));
    let browser, preview;
    try {
      const deck = createPresentation();
      const slide = addBlankSlide(deck);
      addSlideMedia(slide, {
        kind: 'audio',
        data: wav(),
        x: inches(1),
        y: inches(1),
        w: inches(3),
        h: inches(1),
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
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.hit').first().click();
      await editor.getByRole('tab', { name: 'Playback', exact: true }).click();
      const read = async () => {
        const saved = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getShapeMediaPlayback(getSlideShapes(getSlides(saved)[0])[0]);
      };
      const open = async () => {
        await editor.getByRole('button', { name: 'Trim Audio', exact: true }).click();
        const dialog = editor.getByRole('dialog', { name: 'Trim Audio', exact: true });
        await dialog.getByRole('button', { name: 'Trim', exact: true }).waitFor();
        await dialog
          .locator('audio')
          .evaluate((element) =>
            element.readyState >= 1
              ? undefined
              : new Promise((resolve) =>
                  element.addEventListener('loadedmetadata', resolve, { once: true }),
                ),
          );
        return dialog;
      };
      let dialog = await open();
      const waveform = dialog.getByRole('img', { name: 'Audio waveform', exact: true });
      await waveform.locator('path[d*="M511"]').waitFor();
      assert.ok((await waveform.boundingBox()).width > (await dialog.boundingBox()).width / 2);
      const trace = await waveform.locator('path').getAttribute('d');
      assert.notEqual(trace.slice(0, trace.indexOf('M1 ')), 'M0 24V24');
      const mediaDuration = Number(
        await dialog.getByLabel('Start Trim', { exact: true }).getAttribute('max'),
      );
      assert.ok(mediaDuration > 5000 && mediaDuration % 50 !== 0);
      assert.equal(
        Number(await dialog.getByLabel('End Trim', { exact: true }).inputValue()),
        mediaDuration,
      );
      await dialog.getByLabel('Start Trim', { exact: true }).fill('4999');
      assert.equal(
        Number(await dialog.getByLabel('Start Trim', { exact: true }).inputValue()),
        mediaDuration - 50,
      );
      await dialog.getByLabel('End Trim', { exact: true }).fill('100');
      assert.equal(
        Number(await dialog.getByLabel('End Trim', { exact: true }).inputValue()),
        mediaDuration,
      );
      await dialog.getByLabel('Start Trim', { exact: true }).fill('500');
      await dialog.getByLabel('End Trim', { exact: true }).fill('4000');
      const dragHandle = async (label, from, to, min, max, reversed = false) => {
        const control = dialog.getByLabel(label, { exact: true });
        const box = await control.boundingBox();
        assert.ok(box);
        const x = (value) =>
          box.x +
          7 +
          (reversed ? 1 - (value - min) / (max - min) : (value - min) / (max - min)) *
            (box.width - 14);
        await page.mouse.move(x(from), box.y + box.height / 2);
        await page.mouse.down();
        await page.mouse.move(x(to), box.y + box.height / 2, { steps: 5 });
        await page.mouse.up();
        assert.equal(Number(await control.inputValue()), to, label);
      };
      await dragHandle('Start Trim', 500, 1000, 0, mediaDuration);
      await dragHandle('End Trim', 4000, 4500, 0, mediaDuration);
      await dragHandle('Fade In', 0, 500, 0, 3500);
      await dragHandle('Fade Out', 0, 500, 0, 3500, true);
      const fadeOut = dialog.getByLabel('Fade Out', { exact: true });
      await fadeOut.press('ArrowLeft');
      assert.equal(await fadeOut.inputValue(), '550');
      await fadeOut.press('ArrowRight');
      assert.equal(await fadeOut.inputValue(), '500');
      await fadeOut.press('End');
      assert.equal(await fadeOut.inputValue(), '3500');
      await fadeOut.press('Home');
      assert.equal(await fadeOut.inputValue(), '0');

      const current = Number(
        await dialog.getByLabel('Current Position', { exact: true }).inputValue(),
      );
      await dragHandle('Current Position', current, 2000, 0, mediaDuration);
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.equal((await read()).trim, undefined);
      dialog = await open();
      await dialog.getByLabel('Start Trim', { exact: true }).fill('500');
      await dialog.getByLabel('End Trim', { exact: true }).fill('4000');
      await dialog.getByLabel('Fade In', { exact: true }).fill('250');
      await dialog.getByLabel('Start Trim', { exact: true }).press('ArrowRight');
      assert.equal(await dialog.getByLabel('Start Trim', { exact: true }).inputValue(), '550');
      await dialog.getByLabel('Start Trim', { exact: true }).press('ArrowLeft');
      await dialog.screenshot({ path: '/tmp/pptx-trim-timeline.png' });
      await dialog.getByRole('button', { name: 'Play', exact: true }).click();
      await dialog.getByRole('button', { name: 'Pause', exact: true }).waitFor();
      await dialog.getByRole('button', { name: 'Play', exact: true }).waitFor();
      assert.equal(await dialog.locator('audio').evaluate((element) => element.currentTime), 4);
      await dialog.getByRole('button', { name: 'Play', exact: true }).click();
      await dialog.getByRole('button', { name: 'Pause', exact: true }).waitFor();
      assert.ok(await dialog.locator('audio').evaluate((element) => element.currentTime < 1));
      await dialog.getByLabel('Start Trim', { exact: true }).fill('500.25');
      await dialog.getByLabel('Fade In', { exact: true }).fill('250.125');
      await dialog.getByRole('button', { name: 'Trim', exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.deepEqual((await read()).trim, { startMs: 500.25, endMs: mediaDuration - 4000 });
      assert.deepEqual((await read()).fade, { inMs: 250.125, outMs: 0 });
      dialog = await open();
      assert.equal(await dialog.getByLabel('Start Trim', { exact: true }).inputValue(), '500.25');
      assert.equal(await dialog.getByLabel('End Trim', { exact: true }).inputValue(), '4000');
      assert.equal(await dialog.getByLabel('Fade In', { exact: true }).inputValue(), '250.125');
      await dialog.press('Escape');
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal((await read()).trim, undefined);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

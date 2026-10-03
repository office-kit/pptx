import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeFill,
  getShapeImageFillLayout,
  getShapeImageOpacity,
  getShapeImageFillBytes,
  getSlideShapes,
  getSlides,
  loadPresentation,
} from '../../../../dist/index.js';
import { startPreview } from '../helpers/server.mjs';

test(
  'picture fill insertion, replacement and type switching preserve settings and saved history',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-picture-fill-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={3} height={2} fill="#00FF00">First</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1100 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const read = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        const shape = getSlideShapes(getSlides(pres)[0])[0];
        return {
          fill: getShapeFill(shape),
          layout: getShapeImageFillLayout(shape),
          opacity: getShapeImageOpacity(shape),
          bytes: getShapeImageFillBytes(shape),
        };
      };
      const image = {
        name: 'pixel.png',
        mimeType: 'image/png',
        buffer: Buffer.from(
          'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==',
          'base64',
        ),
      };
      const upload = async (trigger) => {
        const chooser = page.waitForEvent('filechooser');
        await trigger.click();
        const persisted = page.waitForResponse(
          (response) =>
            response.url().endsWith('/editor/document') &&
            response.request().method() === 'PUT' &&
            response.ok(),
        );
        await (await chooser).setFiles(image);
        await persisted;
        await editor
          .getByRole('radio', { name: 'Picture or texture fill', exact: true, checked: true })
          .waitFor();
        await saved();
      };
      const number = async (label, value) => {
        const input = editor.getByRole('spinbutton', { name: label, exact: true });
        await input.fill(String(value));
        await input.press('Tab');
        await saved();
      };
      await saved();
      await editor.locator('.hit').first().click();
      const original = (await read()).fill;
      const picture = editor.getByRole('radio', { name: 'Picture or texture fill', exact: true });
      await upload(picture);
      assert.equal(await picture.isChecked(), true);
      assert.equal((await read()).fill.kind, 'image');
      await picture.click();
      assert.equal(await picture.isChecked(), true);
      await number('Picture transparency', 35);
      await number('Offset left', 25);
      const tiled = editor.getByRole('checkbox', { name: 'Tile picture as texture', exact: true });
      await tiled.check();
      await saved();
      await number('Scale X', 60);
      await tiled.uncheck();
      await saved();
      assert.equal((await read()).layout.left, 0.25);
      await tiled.check();
      await saved();
      assert.equal((await read()).layout.scaleX, 0.6);
      await upload(editor.getByRole('button', { name: 'Insert...', exact: true }));
      let replacement = await read();
      assert.equal(replacement.opacity, 0.65);
      assert.equal(replacement.layout.mode, 'stretch');
      assert.equal(replacement.layout.left, 0.25);
      assert.deepEqual(replacement.bytes, new Uint8Array(image.buffer));
      const clipboard = editor.getByRole('button', { name: 'Clipboard', exact: true });
      await editor.locator('body').evaluate(() => {
        Object.defineProperty(navigator.clipboard, 'read', {
          configurable: true,
          value: async () => [],
        });
      });
      await clipboard.click();
      await editor
        .getByRole('alert')
        .filter({ hasText: 'The clipboard does not contain a picture.' })
        .waitFor();
      assert.deepEqual(await read(), replacement);
      await editor.locator('body').evaluate(() => {
        Object.defineProperty(navigator.clipboard, 'read', {
          configurable: true,
          value: async () => {
            throw new DOMException('Denied', 'NotAllowedError');
          },
        });
      });
      await clipboard.click();
      await editor.getByRole('alert').filter({ hasText: 'Clipboard access was denied' }).waitFor();
      assert.deepEqual(await read(), replacement);
      const clipboardBytes = await editor.locator('body').evaluate(async () => {
        const canvas = document.createElement('canvas');
        canvas.width = canvas.height = 2;
        const context = canvas.getContext('2d');
        context.fillStyle = '#0000ff';
        context.fillRect(0, 0, 2, 2);
        const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
        Object.defineProperty(navigator.clipboard, 'read', {
          configurable: true,
          value: async () => [new ClipboardItem({ 'image/png': blob })],
        });
        return [...new Uint8Array(await blob.arrayBuffer())];
      });
      const pasted = page.waitForResponse(
        (response) =>
          response.url().endsWith('/editor/document') &&
          response.request().method() === 'PUT' &&
          response.ok(),
      );
      await clipboard.click();
      await pasted;
      await saved();
      assert.deepEqual(await read(), { ...replacement, bytes: new Uint8Array(clipboardBytes) });
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), replacement);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      replacement = { ...replacement, bytes: new Uint8Array(clipboardBytes) };
      assert.deepEqual(await read(), replacement);
      await editor.getByRole('radio', { name: 'Solid fill', exact: true }).check();
      await saved();
      assert.deepEqual((await read()).fill, original);
      await picture.click();
      await saved();
      assert.equal(await picture.isChecked(), true);
      assert.deepEqual(await read(), replacement);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual((await read()).fill, original);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), replacement);
      await page.reload();
      await saved();
      await editor.locator('.hit').first().click();
      assert.equal(await picture.isChecked(), true);
      assert.equal(
        await editor
          .getByRole('spinbutton', { name: 'Picture transparency', exact: true })
          .inputValue(),
        '35',
      );
      assert.deepEqual(await read(), replacement);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

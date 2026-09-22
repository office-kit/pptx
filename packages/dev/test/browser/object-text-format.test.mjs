import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeRunFormat,
  getShapeText,
  getSlideShapes,
  getSlides,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'object text formatting applies to the whole selection in English and Japanese',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-object-text-format-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={2} height={1}>日本語</Text><Text x={4} y={1} width={2} height={1}>English</Text><Text x={1} y={3} width={2} height={1}>Untouched</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      let ja = false;
      const saved = () =>
        editor
          .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
          .waitFor();
      const read = async () => {
        const deck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getSlideShapes(getSlides(deck)[0]).map((shape) => ({
          text: getShapeText(shape),
          format: getShapeRunFormat(shape, 0, 0),
        }));
      };
      await saved();
      const original = await read();
      await editor.locator('.hit').nth(0).click();
      const bar = editor.locator('.bespoke .text-format-bar');
      await bar.getByRole('button', { name: 'Bold', exact: true }).click();
      await saved();
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      assert.equal(
        await bar.getByRole('button', { name: 'Bold', exact: true }).getAttribute('aria-pressed'),
        'false',
      );
      await bar.getByRole('button', { name: 'Bold', exact: true }).click();
      await saved();
      assert.deepEqual(
        (await read()).map((value) => value.format.bold),
        [true, true, original[2].format.bold],
      );
      await bar.getByRole('spinbutton', { name: 'Font size', exact: true }).fill('32');
      await bar.getByRole('spinbutton', { name: 'Font size', exact: true }).press('Tab');
      await saved();
      const sized = await read();
      assert.deepEqual(
        sized.slice(0, 2).map((value) => value.format.size),
        [32, 32],
      );
      assert.deepEqual(sized[2], original[2]);
      assert.deepEqual(
        sized.map((value) => value.text),
        original.map((value) => value.text),
      );
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(
        (await read()).map((value) => value.format.size),
        original.map((value) => value.format.size),
      );
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), sized);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await bar.getByRole('button', { name: '斜体', exact: true }).click();
      await saved();
      const expected = await read();
      assert.deepEqual(
        expected.slice(0, 2).map((value) => value.format.italic),
        [true, true],
      );
      assert.deepEqual(expected[2], original[2]);
      await bar.getByRole('button', { name: '文字の書式を解除', exact: true }).click();
      await saved();
      const cleared = await read();
      assert.deepEqual(
        cleared
          .slice(0, 2)
          .map((value) => [value.format.bold, value.format.italic, value.format.size]),
        [
          [undefined, undefined, undefined],
          [undefined, undefined, undefined],
        ],
      );
      assert.deepEqual(cleared[2], original[2]);
      assert.deepEqual(
        cleared.map((value) => value.text),
        original.map((value) => value.text),
      );
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), expected);
      await page.reload();
      await saved();
      assert.deepEqual(await read(), expected);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

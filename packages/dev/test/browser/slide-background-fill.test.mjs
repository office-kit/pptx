import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeFill,
  getSlideShapes,
  getSlides,
  loadPresentation,
} from '../../../../dist/index.js';
import { startPreview } from '../helpers/server.mjs';

test(
  'slide background fill supports multiple selections, undo and saved reloads',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-background-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={3} height={2} fill="#00FF00">First</Text><Text x={6} y={1} width={3} height={2} fill="#FF0000">Second</Text></Slide></Presentation>`,
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
        return getSlideShapes(getSlides(pres)[0]).map((shape) => getShapeFill(shape));
      };
      const undo = async () => {
        await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
        await saved();
      };
      const redo = async () => {
        await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
        await saved();
      };
      await saved();
      await editor.locator('.hit').first().click();
      const originals = await read();
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      await editor.getByRole('radio', { name: 'Slide background fill', exact: true }).check();
      await saved();
      assert.deepEqual(await read(), [{ kind: 'background' }, { kind: 'background' }]);
      await undo();
      assert.deepEqual(await read(), originals);
      await redo();
      assert.deepEqual(await read(), [{ kind: 'background' }, { kind: 'background' }]);
      await page.reload();
      await saved();
      await editor.locator('.hit').first().click();
      assert.equal(
        await editor.getByRole('radio', { name: 'Slide background fill', exact: true }).isChecked(),
        true,
      );
      await editor.getByRole('radio', { name: 'Solid fill', exact: true }).check();
      await saved();
      assert.equal((await read())[0].kind, 'solid');
      assert.equal((await read())[1].kind, 'background');
      await undo();
      await editor.locator('.lang select').selectOption('ja');
      assert.equal(
        await editor
          .getByRole('radio', { name: '塗りつぶし（スライドの背景）', exact: true })
          .isChecked(),
        true,
      );
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeGradientFill,
  getSlideShapes,
  getSlides,
  loadPresentation,
} from '../../../../dist/index.js';
import { startPreview } from '../helpers/server.mjs';

test(
  'gradient stop transparency updates the editing preview and survives undo and reload',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-gradient-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={3} height={2} fill={{stops:[{offset:0,color:'accent1',brightness:0.95},{offset:1,color:'accent1',brightness:0.7}],angleDeg:90,scaled:false}}>Gradient</Text><Text x={6} y={1} width={3} height={2} fill="#00FF00">Solid</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const gradient = async (index = 0) => {
        const deck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getShapeGradientFill(getSlideShapes(getSlides(deck)[0])[index]);
      };
      await saved();
      await editor.locator('.hit').nth(0).click();
      const track = editor.getByRole('group', { name: 'Gradient stops', exact: true });
      const paint = () => track.evaluate((element) => getComputedStyle(element).backgroundImage);
      const original = await paint();
      const transparency = editor.getByRole('spinbutton', {
        name: 'Gradient stop transparency',
        exact: true,
      });
      await transparency.fill('60');
      await transparency.press('Tab');
      await saved();
      assert.equal((await gradient()).stops[0].opacity, 0.4);
      const translucent = await paint();
      assert.notEqual(translucent, original, 'stop transparency must affect the editing track');
      assert.match(translucent, /rgba\([^)]*, 0\.4\) 0%/);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(await paint(), original);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.equal(await paint(), translucent);
      await page.reload();
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      await saved();
      await editor.locator('.hit').nth(0).click();
      assert.equal(await paint(), translucent);
      await transparency.fill('100');
      await transparency.press('Tab');
      await saved();
      assert.match(await paint(), /rgba\([^)]*, 0\) 0%/);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

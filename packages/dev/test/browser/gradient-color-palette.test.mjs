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
  'gradient color palette preserves theme references, undo and reload',
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
      const color = editor.getByRole('button', { name: 'Gradient stop color', exact: true });
      await color.click();
      const menu = editor.getByRole('menu', { name: 'Gradient stop color', exact: true });
      assert.equal(await menu.getByRole('menuitemradio').count(), 20);
      const initial = await gradient();
      await menu.press('Escape');
      assert.deepEqual(await gradient(), initial);
      await color.click();
      await menu.getByRole('menuitemradio', { name: 'Accent 2', exact: true }).click();
      await saved();
      const changed = await gradient();
      assert.equal(changed.stops[0].color, 'scheme:accent2');
      assert.equal(changed.stops[0].brightness ?? 0, 0);
      assert.deepEqual(changed.stops.slice(1), initial.stops.slice(1));
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await gradient(), initial);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual(await gradient(), changed);
      await page.reload();
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      await saved();
      await editor.locator('.hit').nth(0).click();
      await color.click();
      assert.equal(
        await menu
          .getByRole('menuitemradio', { name: 'Accent 2', exact: true })
          .getAttribute('aria-checked'),
        'true',
      );
      await menu.getByRole('menuitemradio', { name: 'Red', exact: true }).click();
      await saved();
      assert.equal((await gradient()).stops[0].color.toUpperCase(), '#FF0000');
      await color.click();
      await page.keyboard.press('Home');
      await page.keyboard.press('ArrowRight');
      await page.keyboard.press('Enter');
      await saved();
      assert.equal((await gradient()).stops[0].color, 'scheme:tx1');
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

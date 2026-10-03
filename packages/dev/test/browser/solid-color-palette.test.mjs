import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeFill,
  getShapeStroke,
  getSlideShapes,
  getSlides,
  loadPresentation,
} from '../../../../dist/index.js';
import { startPreview } from '../helpers/server.mjs';

test(
  'solid fill and outline palettes preserve theme references and history',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-solid-palette-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={3} height={2} fill="#123456">First</Text><Text x={6} y={1} width={3} height={2} fill="#00FF00">Solid</Text></Slide></Presentation>`,
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
      const paints = async (index = 0) => {
        const deck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        const shape = getSlideShapes(getSlides(deck)[0])[index];
        return { fill: getShapeFill(shape), stroke: getShapeStroke(shape) };
      };
      await saved();
      await editor.locator('.hit').nth(0).click();
      const color = editor.getByRole('button', { name: 'Fill', exact: true });
      await color.click();
      const menu = editor.getByRole('menu', { name: 'Fill', exact: true });
      assert.equal(await menu.getByRole('menuitemradio').count(), 20);
      const initial = await paints();
      await menu.press('Escape');
      assert.deepEqual(await paints(), initial);
      await color.click();
      await menu.getByRole('menuitemradio', { name: 'Accent 2', exact: true }).click();
      await saved();
      const changed = await paints();
      assert.equal(changed.fill.color, 'scheme:accent2');
      assert.deepEqual(changed.stroke, initial.stroke);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await paints(), initial);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual(await paints(), changed);
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
      assert.equal((await paints()).fill.color.toUpperCase(), '#FF0000');
      await color.click();
      await page.keyboard.press('Home');
      await page.keyboard.press('ArrowRight');
      await page.keyboard.press('Enter');
      await saved();
      assert.equal((await paints()).fill.color, 'scheme:tx1');
      await editor.getByRole('button', { name: 'Outline', exact: true }).click();
      await editor
        .getByRole('menu', { name: 'Outline', exact: true })
        .getByRole('menuitemradio', { name: 'Accent 3', exact: true })
        .click();
      await saved();
      const outlined = await paints();
      assert.equal(outlined.stroke.color, 'scheme:accent3');
      assert.equal(outlined.fill.color, 'scheme:tx1');
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual((await paints()).stroke, initial.stroke);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

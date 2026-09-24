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
  'gradient stops edit brightness, opacity, position and count with undo and persistence',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-gradient-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={3} height={2} fill={{stops:[{offset:0,color:'accent1',brightness:0.95},{offset:1,color:'accent1',brightness:0.7}],angleDeg:90,scaled:true}}>Gradient</Text></Slide></Presentation>`,
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
      const gradient = async () => {
        const deck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getShapeGradientFill(getSlideShapes(getSlides(deck)[0])[0]);
      };
      await saved();
      await editor.locator('.hit').click();
      const brightness = editor.getByRole('spinbutton', {
        name: 'Gradient stop brightness',
        exact: true,
      });
      const opacity = editor.getByRole('spinbutton', {
        name: 'Gradient stop transparency',
        exact: true,
      });
      const position = editor.getByRole('spinbutton', {
        name: 'Gradient stop position',
        exact: true,
      });
      assert.equal(await brightness.inputValue(), '95');
      const paintedStop = editor.locator('.paint linearGradient stop').first();
      const initialColor = await paintedStop.getAttribute('stop-color');
      assert.equal(
        await editor
          .getByRole('button', { name: 'Remove gradient stop', exact: true })
          .isDisabled(),
        true,
      );
      await brightness.fill('-25');
      await brightness.press('Tab');
      await saved();
      assert.equal((await gradient()).stops[0].brightness, -0.25);
      assert.notEqual(await paintedStop.getAttribute('stop-color'), initialColor);
      await opacity.fill('60');
      await opacity.press('Tab');
      await saved();
      assert.equal((await gradient()).stops[0].opacity, 0.4);
      assert.equal(await paintedStop.getAttribute('stop-opacity'), '0.4');
      await editor.getByRole('button', { name: 'Gradient stop 2', exact: true }).click();
      assert.equal(await brightness.inputValue(), '70');
      await position.fill('80');
      await position.press('Tab');
      await saved();
      assert.equal((await gradient()).stops[1].offset, 0.8);
      assert.equal(await position.inputValue(), '80');
      await editor.getByRole('button', { name: 'Add gradient stop', exact: true }).click();
      await saved();
      assert.equal((await gradient()).stops.length, 3);
      assert.equal(await position.inputValue(), '90');
      assert.equal(await brightness.inputValue(), '0');
      assert.match((await gradient()).stops[2].color, /^#[0-9A-F]{6}$/i);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await gradient()).stops.length, 2);
      await position.fill('100');
      await position.press('Tab');
      await saved();
      assert.equal((await gradient()).stops[1].offset, 1);
      await editor.getByRole('button', { name: 'Add gradient stop', exact: true }).click();
      await saved();
      assert.equal(await position.inputValue(), '50');
      assert.equal(await brightness.inputValue(), '0');
      assert.equal(await opacity.inputValue(), '30');
      await editor.getByRole('button', { name: 'Remove gradient stop', exact: true }).click();
      await saved();
      assert.equal((await gradient()).stops.length, 2);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await gradient()).stops.length, 3);
      await page.reload();
      await saved();
      await editor.locator('.hit').click();
      assert.equal(await brightness.inputValue(), '-25');
      assert.equal(await opacity.inputValue(), '60');
      await brightness.fill('101');
      await brightness.press('Tab');
      assert.equal(await brightness.inputValue(), '-25');
      const handle = editor.getByRole('button', { name: 'Gradient stop 1', exact: true });
      const bounds = await handle.boundingBox();
      const track = await editor
        .getByRole('group', { name: 'Gradient stops', exact: true })
        .boundingBox();
      assert.ok(bounds && track);
      await page.mouse.move(bounds.x + bounds.width / 2, bounds.y + bounds.height / 2);
      await page.mouse.down();
      await page.mouse.move(track.x + track.width / 4, bounds.y + bounds.height / 2, { steps: 5 });
      assert.equal(await position.inputValue(), '25');
      await page.mouse.up();
      await saved();
      assert.equal((await gradient()).stops[0].offset, 0.25);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await gradient()).stops[0].offset, 0);
      await page.mouse.move(bounds.x + bounds.width / 2, bounds.y + bounds.height / 2);
      await page.mouse.down();
      await page.mouse.move(track.x + track.width / 3, bounds.y + bounds.height / 2, { steps: 5 });
      await page.keyboard.press('Escape');
      await page.mouse.up();
      assert.equal(await position.inputValue(), '0');
      assert.equal((await gradient()).stops[0].offset, 0);
      const type = editor.getByRole('combobox', { name: 'Gradient type', exact: true });
      await type.selectOption('circle');
      await saved();
      assert.equal((await gradient()).path, 'circle');
      assert.equal(
        await editor.getByRole('spinbutton', { name: 'Gradient angle', exact: true }).isDisabled(),
        true,
      );
      await page.screenshot({ path: '/tmp/pptx-pr287-gradient-panel.png', fullPage: true });
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

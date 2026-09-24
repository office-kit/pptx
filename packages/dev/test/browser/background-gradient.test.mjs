import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlideBackgroundGradientFill,
  getSlides,
  loadPresentation,
} from '../../../../dist/index.js';
import { startPreview } from '../helpers/server.mjs';

test(
  'background gradients edit selected slides, undo and persist after reload',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-background-gradient-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation>{['A','B','C'].map(text => <Slide><Text x={1} y={1} width={7} height={1}>{text}</Text></Slide>)}</Presentation>`,
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
      const gradients = async () =>
        getSlides(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
        ).map(getSlideBackgroundGradientFill);
      const thumbs = editor.locator('.thumb-row');
      await saved();
      await thumbs.nth(0).click();
      await thumbs.nth(1).click({ modifiers: ['Shift'] });
      const pane = editor.getByRole('region', { name: 'Slide options', exact: true });
      await pane.getByRole('radio', { name: 'Gradient fill', exact: true }).check();
      await saved();
      let values = await gradients();
      assert.equal(values[0].angleDeg, 0);
      assert.deepEqual(values[0], values[1]);
      assert.equal(values[2], null);
      assert.equal(
        await pane.getByRole('checkbox', { name: 'Rotate with shape', exact: true }).isEnabled(),
        false,
      );
      assert.equal(
        await pane
          .getByRole('checkbox', { name: 'Rotate with shape', exact: true })
          .evaluate((node) => node.indeterminate),
        true,
      );
      const brightness = pane.getByRole('spinbutton', {
        name: 'Gradient stop brightness',
        exact: true,
      });
      await brightness.fill('25');
      await brightness.press('Tab');
      await saved();
      values = await gradients();
      assert.equal(values[0].stops[0].brightness, 0.25);
      assert.deepEqual(values[0], values[1]);
      const opacity = pane.getByRole('spinbutton', {
        name: 'Gradient stop transparency',
        exact: true,
      });
      await opacity.fill('40');
      await opacity.press('Tab');
      await saved();
      assert.equal((await gradients())[0].stops[0].opacity, 0.6);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await gradients(), values);
      await pane.getByRole('button', { name: 'Add gradient stop', exact: true }).click();
      await saved();
      assert.equal((await gradients())[0].stops.length, 3);
      await pane.getByLabel('Gradient type', { exact: true }).selectOption('rect');
      await saved();
      values = await gradients();
      assert.equal(values[0].path, 'rect');
      assert.deepEqual(values[0], values[1]);
      assert.equal(values[2], null);
      await page.screenshot({ path: '/tmp/pptx-background-gradient-panel.png' });
      await page.reload();
      await saved();
      assert.deepEqual(await gradients(), values);
      await thumbs.nth(0).click();
      await pane.getByRole('button', { name: 'Reset background', exact: true }).click();
      await saved();
      const reset = await gradients();
      assert.equal(reset[0], null);
      assert.deepEqual(reset[1], values[1]);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await gradients(), values);
      await editor.locator('.lang select').selectOption('ja');
      assert.equal(
        await editor
          .getByRole('radio', { name: '塗りつぶし（グラデーション）', exact: true })
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

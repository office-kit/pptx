import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getGridSpacing,
  getSnapToGrid,
  getDrawingGuides,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';

test(
  'grid options cancel, application visibility, saved settings and guide dragging',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-grid-guides-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={3} height={1}>Grid test</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const context = await browser.newContext({ viewport: { width: 1500, height: 1000 } });
      const page = await context.newPage();
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const revision = (await waitForState(preview.url, () => true)).revision;
      const open = async () => {
        await editor.getByRole('button', { name: 'View', exact: true }).click();
        await editor.getByRole('menuitem', { name: 'Grid and Guides', exact: true }).click();
        await editor.getByRole('menuitem', { name: 'Grid Options...', exact: true }).click();
      };
      const dialog = editor.getByRole('dialog', { name: 'Grid and Guides' });
      await open();
      await page.screenshot({ path: '/tmp/pptx-grid-options-current.png' });
      await dialog.getByLabel('Display grid on screen', { exact: true }).check();
      await dialog.getByRole('button', { name: 'Set as Default' }).click();
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.equal(await editor.locator('.grid-dots').count(), 0);
      await open();
      assert.equal(
        await dialog.getByLabel('Display grid on screen', { exact: true }).isChecked(),
        false,
      );
      await dialog.getByLabel('Display grid on screen', { exact: true }).check();
      await dialog.getByLabel('Display drawing guides on screen', { exact: true }).check();
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await editor.locator('.grid-dots').waitFor();
      assert.equal(await editor.locator('.drawing-guide').count(), 2);
      assert.equal((await waitForState(preview.url, () => true)).revision, revision);
      await open();
      await dialog.getByLabel('Snap objects to grid', { exact: true }).check();
      await dialog.getByRole('combobox').selectOption('90000');
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await waitForState(preview.url, (state) => state.revision !== revision);
      const deck = async () =>
        loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
      assert.deepEqual(getGridSpacing(await deck()), { x: 90000, y: 90000 });
      assert.equal(getSnapToGrid(await deck()), true);
      const guide = editor.getByRole('button', { name: 'Vertical guide', exact: true });
      const bounds = await guide.boundingBox();
      const currentRevision = (await waitForState(preview.url, () => true)).revision;
      await page.mouse.move(bounds.x + bounds.width / 2, bounds.y + 25);
      await page.mouse.down();
      await page.mouse.move(bounds.x + 45, bounds.y + 25, { steps: 5 });
      await page.mouse.up();
      await waitForState(preview.url, (state) => state.revision !== currentRevision);
      assert.equal(getDrawingGuides(await deck()).length, 2);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal(getDrawingGuides(await deck()), null);
      await guide.click({ button: 'right', position: { x: 3, y: 25 } });
      await page.keyboard.press('Delete');
      assert.equal(await editor.locator('.nav [data-slide-index]').count(), 1);
      assert.equal(await editor.locator('.drawing-guide').count(), 2);
      await editor.getByRole('menuitem', { name: 'Add Horizontal Guide', exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal(getDrawingGuides(await deck()).length, 3);
      await guide.click({ button: 'right', position: { x: 3, y: 25 } });
      await editor.getByRole('menuitem', { name: 'Color', exact: true }).click();
      await editor.getByRole('menuitem', { name: 'Red', exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal(
        getDrawingGuides(await deck()).find((item) => item.axis === 'x').color,
        '#c43c3c',
      );
      const second = await page.context().newPage();
      await second.goto(preview.url);
      await second.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const otherEditor = second.frameLocator('#editor-frame');
      await otherEditor.getByRole('button', { name: 'View', exact: true }).waitFor();
      await otherEditor.locator('.grid-dots').waitFor();
      await open();
      await dialog.getByLabel('Display grid on screen', { exact: true }).uncheck();
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await otherEditor.locator('.grid-dots').waitFor({ state: 'detached' });
      await second.close();
      await open();
      await dialog.getByRole('combobox').selectOption('custom');
      await dialog.getByRole('spinbutton').fill('0.50001');
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.deepEqual(getGridSpacing(await deck()), { x: 180004, y: 180004 });
      while (await editor.locator('.drawing-guide').count()) {
        await editor.locator('.drawing-guide').first().focus();
        await page.keyboard.press('Delete');
      }
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.deepEqual(getDrawingGuides(await deck()), []);
      await editor.locator('.stage').click({ button: 'right', position: { x: 20, y: 20 } });
      await editor.getByRole('menuitem', { name: 'Add Vertical Guide', exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal(getDrawingGuides(await deck()).length, 1);
      assert.equal(await editor.locator('.drawing-guide').count(), 1);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { startPreview, waitForState } from '../helpers/server.mjs';

test(
  'slide sorter selection, keyboard navigation, undo and independent view zoom',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-view-modes-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation>{Array.from({length:8},(_,i)=><Slide><Text x={1} y={1} width={3} height={1}>Slide {i+1}</Text></Slide>)}</Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const revision = (await waitForState(preview.url, () => true)).revision;
      await editor.getByTitle('Reset to 100%', { exact: true }).click();
      const normalZoom = '100%';
      await editor.getByRole('button', { name: 'View', exact: true }).click();
      await page.keyboard.press('Delete');
      assert.equal(await editor.locator('.nav [data-slide-index]').count(), 8);
      await editor.getByRole('menuitemradio', { name: 'Slide Sorter', exact: true }).click();
      const sorter = editor.locator('.nav.sorter');
      await sorter.waitFor();
      assert.equal(
        await editor
          .getByRole('button', { name: 'Slide Sorter', exact: true })
          .getAttribute('aria-pressed'),
        'true',
      );
      const first = sorter.locator('[data-slide-index="0"]');
      await first.click();
      await first.press('ArrowRight');
      assert.equal(
        await sorter
          .locator('[data-slide-index="1"]')
          .evaluate((node) => node === document.activeElement),
        true,
      );
      await sorter.locator('[data-slide-index="1"]').press('Shift+ArrowRight');
      assert.equal(await sorter.locator('[aria-pressed="true"]').count(), 2);
      await editor.getByTitle('Zoom in (Ctrl+=)', { exact: true }).click();
      assert.equal(await editor.getByTitle('Reset to 100%', { exact: true }).innerText(), '120%');
      await editor.getByRole('button', { name: 'Normal', exact: true }).click();
      assert.equal(
        await editor.getByTitle('Reset to 100%', { exact: true }).innerText(),
        normalZoom,
      );
      await editor.getByRole('button', { name: 'Slide Sorter', exact: true }).click();
      assert.equal(await editor.getByTitle('Reset to 100%', { exact: true }).innerText(), '120%');
      assert.equal((await waitForState(preview.url, () => true)).revision, revision);
      await sorter.locator('[data-slide-index="1"]').focus();
      await page.keyboard.press('Delete');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal(await sorter.locator('[data-slide-index]').count(), 6);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal(await sorter.locator('[data-slide-index]').count(), 8);
      await sorter.locator('[data-slide-index="3"]').dblclick();
      await sorter.waitFor({ state: 'detached' });
      assert.equal(
        await editor.locator('.nav [data-slide-index="3"]').getAttribute('aria-current'),
        'true',
      );
      await editor.getByRole('button', { name: 'View', exact: true }).focus();
      await page.keyboard.press('Meta+2');
      await sorter.waitFor();
      await page.keyboard.press('Meta+1');
      await sorter.waitFor({ state: 'detached' });
      const openZoom = async () => {
        await editor.getByRole('button', { name: 'View', exact: true }).click();
        await editor.getByRole('menuitem', { name: 'Zoom', exact: true }).click();
        await editor.getByRole('menuitem', { name: 'Zoom...', exact: true }).click();
      };
      await openZoom();
      const zoomDialog = editor.getByRole('dialog', { name: 'Zoom', exact: true });
      await zoomDialog.getByRole('spinbutton').fill('175');
      await zoomDialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.equal(
        await editor.getByTitle('Reset to 100%', { exact: true }).innerText(),
        normalZoom,
      );
      await openZoom();
      await zoomDialog.getByRole('spinbutton').fill('175');
      await zoomDialog.getByRole('button', { name: 'OK', exact: true }).click();
      assert.equal(await editor.getByTitle('Reset to 100%', { exact: true }).innerText(), '175%');
      await openZoom();
      await zoomDialog.getByRole('radio', { name: 'Fit', exact: true }).check();
      await zoomDialog.getByRole('button', { name: 'OK', exact: true }).click();
      assert.notEqual(
        await editor.getByTitle('Reset to 100%', { exact: true }).innerText(),
        '175%',
      );
      await page.screenshot({ path: '/tmp/pptx-view-modes.png' });
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

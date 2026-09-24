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
      await editor.getByTitle('Zoom...', { exact: true }).click();
      await editor
        .getByRole('dialog', { name: 'Zoom', exact: true })
        .getByRole('radio', { name: '100%', exact: true })
        .check();
      await editor
        .getByRole('dialog', { name: 'Zoom', exact: true })
        .getByRole('button', { name: 'OK', exact: true })
        .click();
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
      assert.equal(await editor.getByTitle('Zoom...', { exact: true }).innerText(), '110%');
      await editor.getByRole('button', { name: 'Normal', exact: true }).click();
      assert.equal(await editor.getByTitle('Zoom...', { exact: true }).innerText(), normalZoom);
      await editor.getByRole('button', { name: 'Slide Sorter', exact: true }).click();
      assert.equal(await editor.getByTitle('Zoom...', { exact: true }).innerText(), '110%');
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
      assert.equal(await editor.getByTitle('Zoom...', { exact: true }).innerText(), normalZoom);
      await openZoom();
      await zoomDialog.getByRole('spinbutton').fill('175');
      await zoomDialog.getByRole('button', { name: 'OK', exact: true }).click();
      assert.equal(await editor.getByTitle('Zoom...', { exact: true }).innerText(), '175%');
      await openZoom();
      await zoomDialog.getByRole('radio', { name: 'Fit', exact: true }).check();
      await zoomDialog.getByRole('button', { name: 'OK', exact: true }).click();
      assert.notEqual(await editor.getByTitle('Zoom...', { exact: true }).innerText(), '175%');
      const slider = editor.getByRole('slider', { name: 'Zoom percentage' });
      const beforeZoomRevision = (await waitForState(preview.url, () => true)).revision;
      await slider.focus();
      await slider.press('End');
      assert.equal(await editor.getByTitle('Zoom...', { exact: true }).innerText(), '400%');
      await slider.press('Home');
      assert.equal(await editor.getByTitle('Zoom...', { exact: true }).innerText(), '10%');
      await slider.press('ArrowRight');
      assert.equal(await editor.getByTitle('Zoom...', { exact: true }).innerText(), '11%');
      await slider.press('Shift+ArrowRight');
      assert.equal(await editor.getByTitle('Zoom...', { exact: true }).innerText(), '21%');
      assert.equal((await waitForState(preview.url, () => true)).revision, beforeZoomRevision);
      const viewTab = editor.getByRole('tab', { name: 'View', exact: true });
      await viewTab.click();
      const viewPanel = editor.getByRole('tabpanel', { name: 'View', exact: true });
      await viewPanel.getByRole('checkbox', { name: 'Thumbnails', exact: true }).uncheck();
      await editor.locator('.nav').waitFor({ state: 'detached' });
      await viewPanel.getByRole('checkbox', { name: 'Thumbnails', exact: true }).check();
      await editor.locator('.nav').waitFor();
      await viewPanel.getByRole('checkbox', { name: 'Guides', exact: true }).check();
      assert.equal(await editor.locator('.drawing-guide').count(), 2);
      await viewPanel.getByRole('button', { name: 'Slide Sorter', exact: true }).click();
      await editor.locator('.nav.sorter').waitFor();
      assert.equal(
        await viewPanel.getByRole('checkbox', { name: 'Thumbnails', exact: true }).isDisabled(),
        true,
      );
      await viewPanel.getByRole('button', { name: 'Normal', exact: true }).click();
      await editor.locator('.nav.sorter').waitFor({ state: 'detached' });
      await viewPanel.getByRole('button', { name: 'Zoom', exact: true }).click();
      await zoomDialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      await viewTab.focus();
      await viewTab.press('Home');
      assert.equal(
        await editor.getByRole('tab', { name: 'Home', exact: true }).getAttribute('aria-selected'),
        'true',
      );
      await editor.getByRole('tab', { name: 'Home', exact: true }).press('End');
      assert.equal(await viewTab.getAttribute('aria-selected'), 'true');
      assert.equal((await waitForState(preview.url, () => true)).revision, beforeZoomRevision);
      await page.screenshot({ path: '/tmp/pptx-view-modes.png' });
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

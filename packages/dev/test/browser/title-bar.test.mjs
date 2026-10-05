import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getSlides, loadPresentation } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const DECK = `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={4} height={1}>Hello</Text></Slide></Presentation>`;

test(
  'title bar AutoSave pauses saving, and the tab row has Comments and Share like PowerPoint',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-title-bar-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(file, DECK);
    let preview;
    let browser;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1512, height: 900 } });
      await page.goto(preview.url + '/editor');
      await page.getByText('Saved to this project', { exact: true }).waitFor();
      const savedSlides = async () =>
        getSlides(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
        ).length;

      // AutoSave off: edits stay unsaved until ⌘S.
      const autoSave = page.getByRole('switch', { name: 'AutoSave' });
      assert.equal(await autoSave.isChecked(), true);
      await autoSave.uncheck();
      await page
        .locator('#ribbon-panel')
        .getByRole('button', { name: 'New Slide', exact: true })
        .click();
      await page.getByText('Unsaved changes', { exact: true }).waitFor();
      await page.waitForTimeout(1500);
      assert.equal(await savedSlides(), 1);
      await autoSave.check();
      await page.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal(await savedSlides(), 2);

      // Comments toggles the comments pane.
      const comments = page
        .locator('.actions')
        .getByRole('button', { name: 'Comments', exact: true });
      await comments.click();
      assert.equal(await comments.getAttribute('aria-pressed'), 'true');
      // The pane is docked and non-modal: the ribbon and thumbnails stay usable,
      // and it follows the selected slide.
      const pane = page.getByRole('dialog', { name: 'Comments', exact: true });
      await page.getByRole('tab', { name: 'View', exact: true }).click();
      await page.getByRole('tab', { name: 'Home', exact: true }).click();
      await page.locator('.thumb-row').nth(1).click();
      assert.equal(await pane.getByRole('combobox', { name: 'Review slide' }).inputValue(), '1');
      await pane.getByRole('button', { name: 'Close', exact: true }).click();
      assert.equal(await pane.count(), 0);

      // Share ▸ Send a Copy downloads the deck; cloud sharing is unavailable.
      await page.getByRole('button', { name: 'Share', exact: true }).click();
      const menu = page.getByRole('menu', { name: 'Share' });
      assert.equal(await menu.getByRole('menuitem', { name: 'Copy Link' }).isDisabled(), true);
      const download = page.waitForEvent('download');
      await menu.getByRole('menuitem', { name: 'Send a Copy (PowerPoint Presentation)' }).click();
      assert.match((await download).suggestedFilename(), /\.pptx$/);

      // The window chrome follows the system's dark appearance, as PowerPoint does.
      const dark = await browser.newPage({
        viewport: { width: 1200, height: 800 },
        colorScheme: 'dark',
      });
      await dark.goto(preview.url + '/editor');
      await dark.getByText('Saved to this project', { exact: true }).waitFor();
      const ribbon = await dark
        .locator('.ribbon')
        .first()
        .evaluate((node) => getComputedStyle(node).backgroundColor);
      assert.equal(ribbon, 'rgb(42, 42, 42)');
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

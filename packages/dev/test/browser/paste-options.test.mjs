import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getShapeText, getSlides, getSlideShapes, loadPresentation } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const DECK = `import {Presentation,Slide} from '@office-kit/pptx-dsl';export default <Presentation><Slide /></Presentation>`;

test(
  'Paste ▾ offers Keep Text Only and Paste Special like the reference desktop app',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-paste-options-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(file, DECK);
    let preview;
    let browser;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const context = await browser.newContext({ viewport: { width: 1512, height: 900 } });
      await context.grantPermissions(['clipboard-read', 'clipboard-write'], {
        origin: preview.url,
      });
      const page = await context.newPage();
      await page.goto(preview.url + '/editor');
      await page.getByText('Saved to this project', { exact: true }).waitFor();
      await page.evaluate(() => navigator.clipboard.writeText('Plain words'));
      const panel = page.locator('#ribbon-panel');
      await panel.getByRole('button', { name: 'Paste options', exact: true }).click();
      const menu = panel.getByRole('menu', { name: 'Paste options', exact: true });
      assert.equal(
        await menu.getByRole('menuitem', { name: 'Paste Special...' }).isDisabled(),
        true,
      );
      await menu.getByRole('menuitem', { name: 'Keep Text Only', exact: true }).click();
      await page
        .getByText('Unsaved changes', { exact: true })
        .or(page.getByText('Saving…'))
        .first()
        .waitFor();
      await page.getByText('Saved to this project', { exact: true }).waitFor();
      const pres = await loadPresentation(
        new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
      );
      assert.ok(
        getSlideShapes(getSlides(pres)[0]).some((shape) => getShapeText(shape) === 'Plain words'),
      );
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

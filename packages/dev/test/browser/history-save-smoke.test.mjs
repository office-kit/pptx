import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeBoundsResolved,
  getSlideShapes,
  getSlides,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';

test(
  'a committed drag autosaves and undoes, and the served deck bytes match after a reload',
  { timeout: 60000 },
  async (t) => {
    const dir = await mkdtemp(join(tmpdir(), 'office-history-save-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={6} height={1}>Smoke</Text></Slide></Presentation>`,
    );
    let preview, browser;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      const consoleErrors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      page.on('console', (message) => {
        if (message.type() === 'error') consoleErrors.push(message.text());
      });
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const x = async () => {
        const deck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getShapeBoundsResolved(deck, getSlideShapes(getSlides(deck)[0])[0]).x;
      };
      const before = await x();

      const box = await editor.locator('.hit').first().boundingBox();
      assert.ok(box);
      const revision = (await waitForState(preview.url, () => true)).revision;
      await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
      await page.mouse.down();
      await page.mouse.move(box.x + box.width / 2 + 70, box.y + box.height / 2, { steps: 5 });
      await page.mouse.up();
      await waitForState(preview.url, (state) => state.revision !== revision);
      assert.ok((await x()) > before);

      const moved = (await waitForState(preview.url, () => true)).revision;
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await waitForState(preview.url, (state) => state.revision !== moved);
      assert.equal(await x(), before);

      await page.reload();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal(await x(), before);
      assert.deepEqual(errors, []);
      // Reported, not asserted: unrelated resource errors would make this flaky.
      t.diagnostic(`console errors: ${JSON.stringify(consoleErrors)}`);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

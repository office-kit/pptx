import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeBoundsResolved,
  getSlideShapes,
  getSlideText,
  getSlides,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';

const firstShapeX = (deck) => getShapeBoundsResolved(deck, getSlideShapes(getSlides(deck)[0])[0]).x;

test(
  'the Download button saves the undo/redo result as a file that reopens in the editor',
  { timeout: 60000 },
  async (t) => {
    const dir = await mkdtemp(join(tmpdir(), 'office-download-button-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={6} height={1}>Download sample</Text></Slide></Presentation>`,
    );
    let preview, browser;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({
        viewport: { width: 1500, height: 1000 },
        acceptDownloads: true,
      });
      const errors = [];
      const consoleErrors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      page.on('console', (message) => {
        if (message.type() === 'error') consoleErrors.push(message.text());
      });
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      // The served deck only records each autosaved step for comparison; the
      // assertions below read the file the Download button produces.
      const servedX = async () =>
        firstShapeX(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
        );
      const settle = async (action) => {
        const revision = (await waitForState(preview.url, () => true)).revision;
        await action();
        await waitForState(preview.url, (state) => state.revision !== revision);
        return servedX();
      };
      const drag = async (dx) => {
        const box = await editor.locator('.hit').first().boundingBox();
        assert.ok(box);
        await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
        await page.mouse.down();
        await page.mouse.move(box.x + box.width / 2 + dx, box.y + box.height / 2, { steps: 5 });
        await page.mouse.up();
      };
      const undo = () => editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      const redo = () => editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();

      const original = await servedX();
      const first = await settle(() => drag(60));
      const second = await settle(() => drag(60));
      assert.ok(original < first && first < second);
      assert.equal(await settle(undo), first);
      assert.equal(await settle(redo), second);
      assert.equal(await settle(undo), first);

      // Download sits under the Quick Access Toolbar's ⋯, as PowerPoint keeps
      // extra commands there.
      await editor.getByRole('button', { name: 'More Commands', exact: true }).click();
      const [download] = await Promise.all([
        page.waitForEvent('download'),
        editor.getByRole('menuitem', { name: 'Download', exact: true }).click(),
      ]);
      assert.match(download.suggestedFilename(), /\.pptx$/);
      const saved = join(dir, download.suggestedFilename());
      await download.saveAs(saved);
      const downloaded = await loadPresentation(new Uint8Array(await readFile(saved)));
      assert.equal(firstShapeX(downloaded), first);
      assert.equal(getSlideText(getSlides(downloaded)[0]), 'Download sample');

      await editor.locator('.topbar input[type="file"]').setInputFiles(saved);
      await editor.getByText(`Opened ${download.suggestedFilename()}`).waitFor();
      assert.equal(await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).isDisabled(), true);
      assert.match(await editor.locator('.paint').textContent(), /Download sample/);
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

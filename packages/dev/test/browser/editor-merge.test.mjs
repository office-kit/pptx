import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { strFromU8, unzipSync } from 'fflate';
import { getShapeText, getSlideShapes, getSlides, loadPresentation } from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';

test(
  'source rewrites merge with editor edits and only real collisions ask for a choice',
  { timeout: 90000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-editor-merge-browser-'));
    const file = join(dir, 'deck.tsx');
    const source = (a, b) =>
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={6} height={1}>${a}</Text><Text x={1} y={3} width={6} height={1}>${b}</Text></Slide></Presentation>`;
    await writeFile(file, source('First', 'Second'));
    let preview, browser, page;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const keep = editor.getByRole('button', { name: 'Keep my edits', exact: true });
      const edit = async (text) => {
        await editor.locator('.hit').first().dblclick();
        await editor.locator('.canvas-shell .inline-edit').fill(text);
        await editor.locator('.canvas-shell .inline-edit').press('Control+Enter');
      };
      const shows = async (...texts) => {
        for (const text of texts) await editor.locator('.paint', { hasText: text }).waitFor();
      };
      const exported = async () =>
        getSlideShapes(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        ).map(getShapeText);
      await saved();

      // A saved edit and a source change to another shape.
      await edit('Edited here');
      await saved();
      let { revision } = await waitForState(preview.url, (state) => state.hasEdits);
      await writeFile(file, source('First', 'Second from source'));
      let state = await waitForState(preview.url, (state) => state.revision !== revision);
      assert.equal(state.conflict, false);
      await shows('Edited here', 'Second from source');
      await saved();
      assert.equal(await keep.count(), 0);
      const sidecar = unzipSync(await readFile(join(dir, '.office-kit', 'deck.tsx.editor.zip')));
      assert.deepEqual(JSON.parse(strFromU8(sidecar['state.json'])), {
        version: 2,
        sourceHash: state.sourceHash,
      });
      await page.reload();
      await saved();
      await shows('Edited here', 'Second from source');

      // An edit still being saved when the source changes another shape.
      let release;
      const gate = new Promise((resolve) => (release = resolve));
      await page.route('**/editor/document', async (route) => {
        if (route.request().method() === 'PUT') await gate;
        await route.continue();
      });
      revision = state.revision;
      await edit('Edited again');
      await editor.getByText('Saving…', { exact: true }).waitFor();
      await writeFile(file, source('First', 'Second changed again'));
      await waitForState(preview.url, (state) => state.revision !== revision);
      release();
      await shows('Edited again', 'Second changed again');
      await saved();
      assert.equal(await keep.count(), 0);
      assert.deepEqual(await exported(), ['Edited again', 'Second changed again']);
      await page.unroute('**/editor/document');

      // The same shape changed in the editor and in the source.
      ({ revision } = await waitForState(preview.url, () => true));
      await writeFile(file, source('First from source', 'Second changed again'));
      await waitForState(preview.url, (state) => state.revision !== revision && state.conflict);
      await editor
        .getByText('Slide 1: TextBox 2 was changed both here and in the source.', { exact: true })
        .waitFor();
      await editor.locator('.lang select').selectOption('ja');
      await editor
        .getByText('スライド 1: TextBox 2 はこのエディターとソースの両方で変更されました。', {
          exact: true,
        })
        .waitFor();
      await editor.getByRole('button', { name: '現在の編集を維持', exact: true }).click();
      await waitForState(preview.url, (state) => !state.conflict);
      assert.deepEqual(await exported(), ['Edited again', 'Second changed again']);
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-editor-merge-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

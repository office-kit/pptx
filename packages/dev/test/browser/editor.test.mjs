import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeBoundsResolved,
  getSlides,
  getSlideShapes,
  getSlideText,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';

test(
  'the development preview edits, autosaves, reloads and resolves conflicts in both languages',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-editor-browser-'));
    const file = join(dir, 'deck.tsx');
    const source = (title) =>
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={6} height={1}>${title}</Text></Slide></Presentation>`;
    await writeFile(file, source('Source title'));
    let preview;
    let browser;
    let page;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.hit').first().dblclick();
      await editor.locator('.inline-edit').fill('日本語の編集 / Edited title');
      await page.route('**/editor/document', async (route) => {
        if (route.request().method() === 'PUT') {
          await route.fulfill({ status: 500, body: 'Simulated save failure' });
        } else await route.continue();
      });
      await editor.locator('.inline-edit').press('Control+s');
      await editor.getByRole('button', { name: 'Retry', exact: true }).waitFor();
      assert.equal((await waitForState(preview.url, () => true)).hasEdits, false);
      assert.match(await editor.locator('.paint').textContent(), /日本語の編集/);
      await page.unroute('**/editor/document');
      await editor.getByRole('button', { name: 'Retry', exact: true }).click();
      await waitForState(preview.url, (state) => state.hasEdits);
      const download = async () =>
        loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
      let presentation = await download();
      assert.equal(getSlideText(getSlides(presentation)[0]), '日本語の編集 / Edited title');
      const before = getShapeBoundsResolved(
        presentation,
        getSlideShapes(getSlides(presentation)[0])[0],
      );
      const box = await editor.locator('.hit').first().boundingBox();
      assert.ok(box);
      const revision = (await waitForState(preview.url, (state) => state.hasEdits)).revision;
      await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
      await page.mouse.down();
      await page.mouse.move(box.x + box.width / 2 + 70, box.y + box.height / 2 + 40, { steps: 5 });
      // A paused drag must not persist an intermediate position.
      await page.waitForTimeout(900);
      assert.equal((await waitForState(preview.url, () => true)).revision, revision);
      await page.mouse.up();
      await waitForState(preview.url, (state) => state.revision !== revision);
      presentation = await download();
      const after = getShapeBoundsResolved(
        presentation,
        getSlideShapes(getSlides(presentation)[0])[0],
      );
      assert.ok(after.x > before.x);
      assert.ok(after.y > before.y);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      presentation = await download();
      assert.equal(
        getShapeBoundsResolved(presentation, getSlideShapes(getSlides(presentation)[0])[0]).x,
        before.x,
      );
      await page.reload();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.match(await editor.locator('.paint').textContent(), /日本語の編集/);

      await editor.locator('.lang select').selectOption('ja');
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      await writeFile(file, source('Changed source'));
      await editor.getByRole('button', { name: '現在の編集を維持', exact: true }).waitFor();
      assert.match(await editor.locator('.paint').textContent(), /日本語の編集/);
      await editor.getByRole('button', { name: '現在の編集を維持', exact: true }).click();
      await waitForState(preview.url, (state) => !state.conflict);
      await page.screenshot({ path: '/tmp/pptx-pr287-editor-ja.png', fullPage: true });
      await page.reload();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      assert.equal(await editor.locator('.lang select').inputValue(), 'ja');
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-pr287-editor-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { transform } from 'esbuild';
import { chromium } from 'playwright';
import { getSlideText, getSlides, loadPresentation } from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';

test(
  'recovers failed saves after reload, protects source changes, and discards drafts in Japanese',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-recovery-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      const source = (text) =>
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={6} height={2}>${text}</Text></Slide></Presentation>`;
      await writeFile(file, source('Original'));
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      page.on('dialog', (dialog) => dialog.accept());
      const errors = [];
      page.on('pageerror', (e) => errors.push(e.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const blockSave = async (route) =>
        route.request().method() === 'PUT'
          ? route.fulfill({
              status: 503,
              contentType: 'application/json',
              body: JSON.stringify({ message: 'Simulated save failure' }),
            })
          : route.continue();
      await saved();
      const storeModule = await transform(
        await readFile(
          new URL('../../../../site/src/lib/editor/dev/draft-store.ts', import.meta.url),
          'utf8',
        ),
        { loader: 'ts', format: 'esm' },
      );
      const versions = await page.evaluate(async (source) => {
        const url = URL.createObjectURL(new Blob([source], { type: 'text/javascript' }));
        const { DraftStore } = await import(url);
        URL.revokeObjectURL(url);
        const store = new DraftStore();
        const draft = {
          id: 'version-test',
          projectId: 'version-test',
          fileName: 'test.pptx',
          bytes: new Uint8Array([2]),
          version: 2,
          updated: Date.now(),
        };
        await store.put(draft);
        await store.put({ ...draft, version: 1, bytes: new Uint8Array([1]) });
        await store.remove(draft.id, 1);
        const remaining = await store.list(draft.projectId);
        await store.remove(draft.id, 2);
        return {
          versions: remaining.map((d) => d.version),
          bytes: remaining.map((d) => Array.from(d.bytes)),
          count: (await store.list(draft.projectId)).length,
        };
      }, storeModule.code);
      assert.deepEqual(versions, { versions: [2], bytes: [[2]], count: 0 });
      await page.route('**/editor/document', blockSave);
      const edit = async (text) => {
        await editor.locator('.hit').first().dblclick();
        await editor.locator('.inline-edit').fill(text);
        await editor.locator('.inline-edit').press('Control+Enter');
      };
      await edit('Recovered 日本語');
      await editor.getByText('Simulated save failure', { exact: false }).waitFor();
      await page.evaluate(async () => {
        const db = await new Promise((resolve, reject) => {
          const request = indexedDB.open('office-pptx-editor-drafts', 1);
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
        await new Promise((resolve, reject) => {
          const tx = db.transaction('drafts', 'readwrite');
          const store = tx.objectStore('drafts');
          const request = store.getAll();
          request.onsuccess = () => {
            if (request.result.length !== 1) {
              tx.abort();
              return;
            }
            store.put({ ...request.result[0], id: 'foreign-draft', projectId: 'another-project' });
          };
          tx.oncomplete = resolve;
          tx.onabort = () => reject(new Error('Could not seed another project draft'));
        });
        db.close();
      });
      await page.reload();
      const recovery = editor.getByRole('dialog', { name: 'Recover unsaved changes', exact: true });
      await recovery.waitFor();
      await recovery.getByRole('button', { name: 'Restore changes', exact: true }).click();
      await editor.getByText('Simulated save failure', { exact: false }).waitFor();
      assert.match(await editor.locator('.paint').textContent(), /Recovered 日本語/);
      await page.unroute('**/editor/document', blockSave);
      await editor.getByRole('button', { name: 'Retry', exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      assert.equal(await recovery.count(), 0);
      const persistedText = async () =>
        getSlideText(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        );
      assert.equal(await persistedText(), 'Recovered 日本語');
      await page.route('**/editor/document', blockSave);
      await edit('Local conflict');
      await editor.getByText('Simulated save failure', { exact: false }).waitFor();
      const revision = (await waitForState(preview.url, () => true)).revision;
      await writeFile(file, source('New source'));
      await waitForState(preview.url, (s) => s.revision !== revision);
      await page.reload();
      await recovery.waitFor();
      await page.unroute('**/editor/document', blockSave);
      await recovery.getByRole('button', { name: 'Restore changes', exact: true }).click();
      await editor.getByRole('button', { name: 'Keep my edits', exact: true }).waitFor();
      assert.equal(await persistedText(), 'Recovered 日本語');
      await editor.getByRole('button', { name: 'Keep my edits', exact: true }).click();
      await saved();
      assert.equal(await persistedText(), 'Local conflict');
      await editor.locator('.lang select').selectOption('ja');
      await page.route('**/editor/document', blockSave);
      await edit('破棄する変更');
      await editor.getByText('Simulated save failure', { exact: false }).waitFor();
      await page.reload();
      await editor
        .getByRole('dialog', { name: '未保存の変更を復元', exact: true })
        .getByRole('button', { name: '復元用コピーを破棄', exact: true })
        .click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      assert.match(await editor.locator('.paint').textContent(), /Local conflict/);
      await page.reload();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      assert.equal(await editor.getByRole('dialog').count(), 0);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

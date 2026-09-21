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
  getSlideXmlString,
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
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={6} height={1} paragraphs={[{runs:[{text:${JSON.stringify(title.slice(0, 7))},format:{bold:true}},{text:${JSON.stringify(title.slice(7))},format:{italic:true}}]}]} /></Slide></Presentation>`;
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
      await editor.locator('.inline-edit').fill('Source headline');
      await editor.locator('.inline-edit').press('Control+Enter');
      await waitForState(preview.url, (state) => state.hasEdits);
      const richDeck = await loadPresentation(
        new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
      );
      const richXml = getSlideXmlString(getSlides(richDeck)[0]);
      assert.match(richXml, /<a:rPr[^>]*b="1"[^>]*>[\s\S]*?<a:t>Source /);
      assert.match(richXml, /<a:rPr[^>]*i="1"/);
      const savedRevision = (await waitForState(preview.url, () => true)).revision;
      await editor.locator('.hit').first().dblclick();
      await editor.locator('.inline-edit').fill('日本語の編集 / Edited title');
      await page.route('**/editor/document', async (route) => {
        if (route.request().method() === 'PUT') {
          await route.fulfill({ status: 500, body: 'Simulated save failure' });
        } else await route.continue();
      });
      await editor.locator('.inline-edit').press('Control+s');
      await editor.getByRole('button', { name: 'Retry', exact: true }).waitFor();
      assert.equal((await waitForState(preview.url, () => true)).revision, savedRevision);
      assert.match(await editor.locator('.paint').textContent(), /日本語の編集/);
      await page.unroute('**/editor/document');
      await editor.getByRole('button', { name: 'Retry', exact: true }).click();
      await waitForState(preview.url, (state) => state.revision !== savedRevision);
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

      // Clipboard snapshots must outlive deletion, editing and history restores.
      await editor.locator('.hit').first().click();
      await page.keyboard.press('Control+x');
      await editor.locator('.hit').waitFor({ state: 'detached' });
      await page.keyboard.press('Control+v');
      await editor.locator('.hit').waitFor();
      assert.match(await editor.locator('.paint').textContent(), /日本語の編集/);
      await page.keyboard.press('Control+z');
      await editor.locator('.hit').waitFor({ state: 'detached' });
      await page.keyboard.press('Control+Shift+z');
      await editor.locator('.hit').waitFor();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      presentation = await download();
      assert.equal(getSlideText(getSlides(presentation)[0]), '日本語の編集 / Edited title');

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

test(
  'slide controls insert, duplicate, reorder and delete with keyboard history and persisted order',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-slides-browser-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={6} height={1}>First</Text></Slide><Slide><Text x={1} y={1} width={6} height={1}>Second</Text></Slide></Presentation>`,
    );
    let preview;
    let browser;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const rows = editor.locator('.thumb-row');
      await rows.first().click();
      await page.keyboard.press('Control+d');
      await editor.getByRole('button', { name: 'Slide 3', exact: true }).waitFor();
      assert.equal(
        await editor.locator('.thumb-row[aria-current="true"]').getAttribute('data-slide-index'),
        '1',
      );
      await page.keyboard.press('Alt+ArrowDown');
      assert.equal(
        await editor.locator('.thumb-row[aria-current="true"]').getAttribute('data-slide-index'),
        '2',
      );
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const titles = async () =>
        getSlides(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
        ).map(getSlideText);
      assert.deepEqual(await titles(), ['First', 'Second', 'First']);
      await page.keyboard.press('Delete');
      await editor
        .getByRole('button', { name: 'Slide 3', exact: true })
        .waitFor({ state: 'detached' });
      await page.keyboard.press('Control+z');
      await editor.getByRole('button', { name: 'Slide 3', exact: true }).waitFor();
      assert.equal(
        await editor.locator('.thumb-row[aria-current="true"]').getAttribute('data-slide-index'),
        '2',
      );
      await editor.locator('.lang select').selectOption('ja');
      await editor.getByRole('button', { name: 'スライドを上へ移動', exact: true }).click();
      await editor.getByRole('button', { name: 'スライドを削除', exact: true }).click();
      await editor.locator('.nav').getByTitle('新しいスライド', { exact: true }).click();
      await editor.getByRole('button', { name: 'スライド 3', exact: true }).waitFor();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      assert.deepEqual(await titles(), ['First', 'Second', '']);
      await page.reload();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      assert.equal(await rows.count(), 3);
      await page.screenshot({ path: '/tmp/pptx-pr287-slide-controls-ja.png', fullPage: true });
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'arrange controls align, distribute, group and ungroup selected objects and persist geometry',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-arrange-browser-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={1} height={0.5}>A</Text><Text x={4} y={2} width={2} height={0.5}>B</Text><Text x={9} y={3} width={3} height={0.5}>C</Text></Slide></Presentation>`,
    );
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
      await editor.locator('.hit').first().click();
      await page.keyboard.press('Control+a');
      const arrange = editor.getByRole('region', { name: 'Arrange', exact: true });
      await arrange.getByRole('button', { name: 'Align top', exact: true }).click();
      await arrange.getByRole('button', { name: 'Distribute horizontally', exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const readDeck = async () =>
        loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
      const deck = await readDeck();
      const boxes = getSlideShapes(getSlides(deck)[0]).map((shape) =>
        getShapeBoundsResolved(deck, shape),
      );
      assert.deepEqual(
        boxes.map((b) => b.x / 914400),
        [1, 4.5, 9],
      );
      assert.deepEqual(
        boxes.map((b) => b.y / 914400),
        [1, 1, 1],
      );
      await arrange.getByRole('button', { name: 'Group', exact: true }).click();
      await editor.locator('.hit').nth(1).waitFor({ state: 'detached' });
      assert.equal(await editor.locator('.hit.selected').count(), 1);
      await page.keyboard.press('ArrowDown');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const groupedDeck = await readDeck();
      assert.match(getSlideXmlString(getSlides(groupedDeck)[0]), /<p:grpSp>/);
      await editor.locator('.lang select').selectOption('ja');
      await editor
        .getByRole('region', { name: '配置', exact: true })
        .getByRole('button', { name: 'グループ解除', exact: true })
        .click();
      await editor.locator('.hit.selected').nth(2).waitFor();
      await page.keyboard.press('Control+z');
      await editor.locator('.hit').nth(1).waitFor({ state: 'detached' });
      await page.keyboard.press('Control+Shift+g');
      await editor.locator('.hit.selected').nth(2).waitFor();
      await page.screenshot({ path: '/tmp/pptx-pr287-arrange-ja.png', fullPage: true });
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      await page.reload();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      assert.equal(await editor.locator('.hit').count(), 3);
      const restored = await readDeck();
      assert.deepEqual(
        getSlideShapes(getSlides(restored)[0]).map((shape) =>
          getShapeBoundsResolved(restored, shape),
        ),
        boxes.map((b) => ({ ...b, y: b.y + 18288 })),
      );
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-pr287-arrange-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { unzipSync, strFromU8 } from 'fflate';
import {
  getShapeChartSpec,
  readPackagePart,
  listPackageParts,
  getTableStyleFlags,
  getTableCells,
  getTableCellSpan,
  getTableCellBorders,
  getTableCellAlignment,
  getTableCellAnchor,
  getTableCellText,
  getTableCellParagraphs,
  getTableCellFill,
  getTableColumnWidths,
  getShapeBoundsResolved,
  getShapeKind,
  getShapeParagraphElements,
  getShapeImageBytes,
  getShapeImageCrop,
  getShapeImageOpacity,
  getShapeImageBrightness,
  getShapeDescription,
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

test(
  'image upload, crop, appearance, replacement and alt text persist from the bilingual preview',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-image-browser-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide} from '@office-kit/pptx-dsl';export default <Presentation><Slide /></Presentation>`,
    );
    let preview, browser, page;
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
      const png = async (color) =>
        Buffer.from(
          await page.evaluate((fill) => {
            const canvas = document.createElement('canvas');
            canvas.width = 200;
            canvas.height = 100;
            const ctx = canvas.getContext('2d');
            ctx.fillStyle = fill;
            ctx.fillRect(0, 0, 200, 100);
            return canvas.toDataURL('image/png').split(',')[1];
          }, color),
          'base64',
        );
      const original = await png('#eb5757'),
        replacement = await png('#2d9cdb');
      await editor.getByRole('button', { name: 'Insert', exact: true }).click();
      await editor.getByTitle(/— addSlideImage$/).click();
      let dialog = editor.getByRole('dialog', { name: 'Insert image', exact: true });
      await dialog
        .getByLabel('Image file', { exact: true })
        .setInputFiles({ name: 'red.png', mimeType: 'image/png', buffer: original });
      await dialog.getByRole('button', { name: 'Insert image', exact: true }).click();
      await dialog.waitFor({ state: 'hidden' });
      await editor.locator('.image-controls').waitFor();
      const change = async (label, value) => {
        await editor.getByLabel(label, { exact: true }).fill(value);
        await editor.getByLabel(label, { exact: true }).press('Tab');
      };
      await change('Crop left (%)', '20');
      await change('Opacity (%)', '70');
      await change('Brightness (%)', '10');
      await change('Alternative text', '赤い画像 / Red image');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const download = async () =>
        loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
      let deck = await download();
      let picture = getSlideShapes(getSlides(deck)[0]).find((s) => getShapeKind(s) === 'picture');
      const bounds = getShapeBoundsResolved(deck, picture);
      assert.equal(bounds.w, bounds.h * 2);
      assert.equal(getShapeImageCrop(picture).left, 0.2);
      assert.equal(getShapeImageOpacity(picture), 0.7);
      assert.equal(getShapeImageBrightness(picture), 0.1);
      assert.equal(getShapeDescription(picture), '赤い画像 / Red image');
      assert.deepEqual(Buffer.from(getShapeImageBytes(picture)), original);
      await editor.locator('.lang select').selectOption('ja');
      await editor.getByRole('button', { name: '画像を差し替え', exact: true }).click();
      dialog = editor.getByRole('dialog', { name: '画像を差し替え', exact: true });
      await dialog
        .getByLabel('画像ファイル', { exact: true })
        .setInputFiles({ name: 'blue.png', mimeType: 'image/png', buffer: replacement });
      await dialog.getByRole('button', { name: '画像を差し替え', exact: true }).click();
      await dialog.waitFor({ state: 'hidden' });
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      deck = await download();
      picture = getSlideShapes(getSlides(deck)[0]).find((s) => getShapeKind(s) === 'picture');
      assert.deepEqual(Buffer.from(getShapeImageBytes(picture)), replacement);
      assert.deepEqual(getShapeBoundsResolved(deck, picture), bounds);
      assert.equal(getShapeImageCrop(picture).left, 0.2);
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      deck = await download();
      picture = getSlideShapes(getSlides(deck)[0]).find((s) => getShapeKind(s) === 'picture');
      assert.deepEqual(Buffer.from(getShapeImageBytes(picture)), original);
      await page.reload();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      await editor.locator('.hit').click();
      assert.equal(
        await editor.getByLabel('左のトリミング (%)', { exact: true }).inputValue(),
        '20',
      );
      assert.equal(
        await editor.getByLabel('代替テキスト', { exact: true }).inputValue(),
        '赤い画像 / Red image',
      );
      await page.screenshot({ path: '/tmp/pptx-pr287-image-ja.png', fullPage: true });
      await editor.getByRole('button', { name: 'トリミングをリセット', exact: true }).click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      deck = await download();
      picture = getSlideShapes(getSlides(deck)[0]).find((s) => getShapeKind(s) === 'picture');
      assert.equal(getShapeImageCrop(picture), null);
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-pr287-image-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'selected text formatting preserves surrounding runs and persists from the bilingual preview',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-text-format-browser-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={6} height={1} paragraphs={[{runs:[{text:'Hello ',format:{bold:true}},{text:'日本語 🌎',format:{italic:true}}]}]} /></Slide></Presentation>`,
    );
    let preview, browser, page;
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
      const input = editor.getByRole('textbox', { name: 'Edit text', exact: true });
      await input.fill('Hello 日本語 🌎!');
      await input.evaluate((node) => {
        node.setSelectionRange(6, 9);
        node.dispatchEvent(new Event('select', { bubbles: true }));
      });
      const bar = editor.getByRole('group', { name: 'Selected text formatting' });
      await bar.getByRole('button', { name: 'Bold', exact: true }).click();
      assert.equal(
        await bar.getByRole('button', { name: 'Bold', exact: true }).getAttribute('aria-pressed'),
        'true',
      );
      await bar.getByRole('button', { name: 'Underline', exact: true }).click();
      await bar.getByRole('textbox', { name: 'Font', exact: true }).fill('Yu Gothic');
      await bar.getByRole('textbox', { name: 'Font', exact: true }).press('Tab');
      const sizeInput = bar.getByRole('spinbutton', { name: 'Font size', exact: true });
      await sizeInput
        .evaluate(
          (node) =>
            new Promise((resolve) =>
              requestAnimationFrame(() =>
                requestAnimationFrame(() => resolve(node === document.activeElement)),
              ),
            ),
        )
        .then((focused) => assert.equal(focused, true));
      await sizeInput.fill('28');
      await bar.getByRole('spinbutton', { name: 'Font size', exact: true }).press('Tab');
      await bar.getByLabel('Text color', { exact: true }).fill('#ff0000');
      await page.screenshot({ path: '/tmp/pptx-pr287-text-format-toolbar.png', fullPage: true });
      await bar.getByRole('button', { name: 'Done', exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal(await editor.locator('.bespoke textarea').inputValue(), 'Hello 日本語 🌎!');
      const runs = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getShapeParagraphElements(getSlideShapes(getSlides(pres)[0])[0], 0);
      };
      let result = await runs();
      assert.deepEqual(
        result.map((r) => r.text),
        ['Hello ', '日本語', ' 🌎', '!'],
      );
      assert.equal(result[0].format.bold, true);
      assert.equal(result[0].format.color, undefined);
      assert.deepEqual(result[1].format, {
        italic: true,
        bold: true,
        underline: true,
        font: 'Yu Gothic',
        fontEastAsian: 'Yu Gothic',
        fontComplexScript: 'Yu Gothic',
        size: 28,
        color: '#FF0000',
      });
      assert.equal(result[2].format.italic, true);
      assert.equal(result[2].format.bold, undefined);
      await editor.locator('select').first().selectOption('ja');
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      assert.equal((await runs())[1].format.color, undefined);
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      await page.reload();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      await editor.locator('.hit').first().dblclick();
      await editor
        .getByRole('textbox', { name: 'テキストを編集', exact: true })
        .evaluate((node) => {
          node.setSelectionRange(6, 9);
          node.dispatchEvent(new Event('select', { bubbles: true }));
        });
      const japanese = editor.getByRole('group', { name: '選択した文字の書式' });
      assert.equal(
        await japanese
          .getByRole('button', { name: '太字', exact: true })
          .getAttribute('aria-pressed'),
        'true',
      );
      await japanese.getByRole('button', { name: '太字', exact: true }).click();
      await japanese.getByRole('button', { name: '完了', exact: true }).click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      result = await runs();
      assert.equal(result[1].format.bold, false);
      assert.equal(result[1].format.color, '#FF0000');
      assert.equal(result[0].format.bold, true);
      // Editing through the properties pane must keep the same mixed formatting.
      await editor.locator('.bespoke textarea').fill('Hello 日本語 🌎!?');
      await editor.locator('.bespoke textarea').press('Tab');
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      result = await runs();
      assert.equal(result.map((r) => r.text).join(''), 'Hello 日本語 🌎!?');
      assert.equal(result[1].format.color, '#FF0000');
      assert.equal(result[1].format.bold, false);
      assert.equal(result[0].format.bold, true);
      await page.screenshot({ path: '/tmp/pptx-pr287-text-format-ja.png', fullPage: true });
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-pr287-text-format-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'table cells, dimensions and rows can be edited and saved in both languages',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-table-browser-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Table} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Table x={1} y={1} width={6} height={2} rows={[[{paragraphs:[{runs:[{text:'Hello ',format:{bold:true}},{text:'日本語',format:{italic:true}}]}]},'B'],['C','D']]} /></Slide></Presentation>`,
    );
    let preview, browser, page;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (e) => errors.push(e.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.hit').first().click();
      const panel = editor.getByRole('region', { name: 'Table options', exact: true });
      const tableBox = await editor.locator('.hit').first().boundingBox();
      await editor
        .locator('.hit')
        .first()
        .dblclick({ position: { x: tableBox.width / 4, y: tableBox.height / 4 } });
      await editor.locator('.inline-edit').fill('Hello 日本語!');
      await editor.locator('.inline-edit').press('Control+Enter');
      assert.equal(
        await panel.getByLabel('Cell text', { exact: true }).inputValue(),
        'Hello 日本語!',
      );
      assert.equal(await editor.locator('.cell-selection').count(), 1);
      await editor
        .locator('.hit')
        .first()
        .dblclick({ position: { x: tableBox.width / 4, y: tableBox.height / 4 } });
      await editor.locator('.inline-edit').evaluate((node) => {
        node.focus();
        node.setSelectionRange(6, 9);
        node.dispatchEvent(new Event('select', { bubbles: true }));
      });
      await editor
        .locator('.canvas-shell .text-format-bar')
        .getByRole('button', { name: 'Underline', exact: true })
        .click();
      await editor
        .locator('.canvas-shell .text-format-bar')
        .getByRole('button', { name: 'Done', exact: true })
        .click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const rangeDeck = await loadPresentation(
        new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
      );
      const rangeCell = getTableCells(getSlideShapes(getSlides(rangeDeck)[0])[0])[0][0];
      const rangeRuns = getTableCellParagraphs(rangeCell)[0].elements;
      assert.equal(rangeRuns[0].format.bold, true);
      assert.equal(rangeRuns[0].format.underline, undefined);
      assert.equal(rangeRuns.find((run) => run.text === '日本語').format.underline, true);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();

      await panel.getByLabel('Cell fill', { exact: true }).fill('#aabbcc');
      await panel.getByLabel('Horizontal alignment', { exact: true }).selectOption('ctr');
      assert.equal(
        await panel.getByLabel('Horizontal alignment', { exact: true }).inputValue(),
        'ctr',
      );
      await panel.getByLabel('Column width (inches)', { exact: true }).fill('2.5');
      await panel.getByLabel('Column width (inches)', { exact: true }).press('Tab');
      await panel.getByRole('button', { name: 'Insert row below', exact: true }).click();
      await panel.getByRole('button', { name: 'Cell 2, 2', exact: true }).click();
      await panel.getByLabel('Cell text', { exact: true }).fill('追加');
      await panel.getByLabel('Cell text', { exact: true }).press('Tab');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const readTable = async () =>
        getSlideShapes(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        )[0];
      let table = await readTable();
      let cells = getTableCells(table);
      assert.equal(cells.length, 3);
      assert.equal(getTableCellText(cells[1][1]), '追加');
      assert.equal(getTableCellText(cells[0][0]), 'Hello 日本語!');
      assert.equal(getTableCellFill(cells[0][0]), '#AABBCC');
      assert.equal(getTableColumnWidths(table)[0], 2286000);
      const runs = getTableCellParagraphs(cells[0][0])[0].elements;
      assert.equal(runs[0].format.bold, true);
      assert.equal(runs[1].format.italic, true);
      await editor.locator('select').first().selectOption('ja');
      await editor.getByRole('button', { name: '列を削除', exact: true }).click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      assert.equal(getTableCells(await readTable())[0].length, 1);
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      assert.equal(getTableCells(await readTable())[0].length, 2);
      await page.reload();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      await editor.locator('.hit').first().click();
      await editor.getByRole('button', { name: 'セル 2, 2', exact: true }).click();
      assert.equal(await editor.getByLabel('セルのテキスト', { exact: true }).inputValue(), '追加');
      await editor.getByRole('button', { name: 'セル 1, 1', exact: true }).click();
      await editor
        .getByRole('button', { name: 'セル 2, 2', exact: true })
        .click({ modifiers: ['Shift'] });
      await editor.getByLabel('セルの塗りつぶし', { exact: true }).fill('#112233');
      await editor.getByLabel('水平方向の配置', { exact: true }).selectOption('r');
      await editor.getByLabel('垂直方向の配置', { exact: true }).selectOption('bottom');
      const tablePanel = editor.getByRole('region', { name: '表の設定', exact: true });
      await tablePanel.getByRole('button', { name: '太字', exact: true }).click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      cells = getTableCells(await readTable());
      for (const cell of cells.slice(0, 2).flat()) {
        assert.equal(getTableCellFill(cell), '#112233');
        assert.equal(getTableCellAlignment(cell), 'r');
        assert.equal(getTableCellAnchor(cell), 'bottom');
        for (const paragraph of getTableCellParagraphs(cell))
          for (const run of paragraph.elements) assert.equal(run.format?.bold, true);
      }
      assert.notEqual(getTableCellFill(cells[2][0]), '#112233');
      assert.notEqual(getTableCellAlignment(cells[2][0]), 'r');
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      cells = getTableCells(await readTable());
      assert.equal(getTableCellParagraphs(cells[0][0])[0].elements[0].format.bold, true);
      assert.notEqual(getTableCellParagraphs(cells[0][0])[0].elements[1].format.bold, true);
      assert.notEqual(getTableCellParagraphs(cells[0][1])[0].elements[0].format?.bold, true);
      await editor.getByRole('button', { name: 'セルを結合', exact: true }).click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      cells = getTableCells(await readTable());
      assert.equal(getTableCellSpan(cells[0][0]).gridSpan, 2);
      assert.equal(getTableCellSpan(cells[0][0]).rowSpan, 2);
      assert.equal(getTableCellText(cells[0][0]), 'Hello 日本語!\nB\n追加');
      await editor.getByRole('button', { name: 'セルを分割', exact: true }).click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      cells = getTableCells(await readTable());
      assert.equal(getTableCellSpan(cells[0][0]).gridSpan, 1);
      assert.equal(getTableCellText(cells[0][1]), '');
      assert.equal(getTableCellText(cells[0][0]), 'Hello 日本語!\nB\n追加');
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      assert.equal(getTableCellSpan(getTableCells(await readTable())[0][0]).gridSpan, 2);
      await page.reload();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      await editor.locator('.hit').first().click();
      assert.equal(await editor.getByRole('button', { name: 'セル 1, 2', exact: true }).count(), 0);
      await editor.locator('select').first().selectOption('en');
      await editor
        .getByRole('region', { name: 'Table options', exact: true })
        .getByRole('button', { name: 'Italic', exact: true })
        .click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      cells = getTableCells(await readTable());
      for (const cell of cells.slice(0, 2).flat())
        for (const paragraph of getTableCellParagraphs(cell))
          for (const run of paragraph.elements) assert.equal(run.format?.italic, true);
      assert.notEqual(getTableCellParagraphs(cells[2][0])[0].elements[0].format?.italic, true);
      const fontPanel = editor.getByRole('region', { name: 'Table options', exact: true });
      await fontPanel.getByLabel('Font', { exact: true }).fill('Arial');
      await fontPanel.getByLabel('Font', { exact: true }).press('Tab');
      await fontPanel.getByLabel('Font size', { exact: true }).fill('18.5');
      await fontPanel.getByLabel('Font size', { exact: true }).press('Tab');
      await fontPanel.getByLabel('Text color', { exact: true }).fill('#ffffff');
      await fontPanel.getByRole('button', { name: 'Underline', exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      cells = getTableCells(await readTable());
      for (const cell of cells.slice(0, 2).flat())
        for (const paragraph of getTableCellParagraphs(cell))
          for (const run of paragraph.elements) {
            assert.equal(run.format.font, 'Arial');
            assert.equal(run.format.fontEastAsian, 'Arial');
            assert.equal(run.format.fontComplexScript, 'Arial');
            assert.equal(run.format.size, 18.5);
            assert.equal(run.format.color, '#FFFFFF');
            assert.equal(run.format.underline, true);
          }
      assert.notEqual(getTableCellParagraphs(cells[2][0])[0].elements[0].format?.size, 18.5);
      await page.reload();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.hit').first().click();
      assert.equal(await fontPanel.getByLabel('Font size', { exact: true }).inputValue(), '18.5');
      assert.equal(
        await fontPanel.getByLabel('Text color', { exact: true }).inputValue(),
        '#ffffff',
      );

      await fontPanel
        .getByRole('button', { name: 'Cell 3, 2', exact: true })
        .click({ modifiers: ['Shift'] });
      assert.equal(await fontPanel.getByLabel('Font size', { exact: true }).inputValue(), '');
      assert.equal(await fontPanel.getByLabel('Font', { exact: true }).inputValue(), '');
      await editor.locator('select').first().selectOption('ja');
      const japaneseFontPanel = editor.getByRole('region', { name: '表の設定', exact: true });
      await japaneseFontPanel.getByLabel('文字サイズ', { exact: true }).fill('20');
      await japaneseFontPanel.getByLabel('文字サイズ', { exact: true }).press('Tab');
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      cells = getTableCells(await readTable());
      assert.equal(getTableCellParagraphs(cells[2][1])[0].elements[0].format.size, 20);
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      cells = getTableCells(await readTable());
      assert.equal(getTableCellParagraphs(cells[0][0])[0].elements[0].format.size, 18.5);
      assert.notEqual(getTableCellParagraphs(cells[2][1])[0].elements[0].format?.size, 20);
      assert.equal(
        await japaneseFontPanel.getByLabel('文字サイズ', { exact: true }).inputValue(),
        '',
      );
      await japaneseFontPanel.getByText('セルの罫線', { exact: true }).click();
      await japaneseFontPanel.getByLabel('罫線の色', { exact: true }).fill('#ee9900');
      await japaneseFontPanel.getByLabel('罫線の太さ（ポイント）', { exact: true }).fill('2');
      await japaneseFontPanel.getByLabel('罫線の種類', { exact: true }).selectOption('dash');
      await japaneseFontPanel.getByRole('button', { name: '罫線を適用', exact: true }).click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      const readBorders = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        const table = getSlideShapes(getSlides(pres)[0])[0];
        return getTableCells(table).map((row) =>
          row.map((cell) => getTableCellBorders(pres, cell)),
        );
      };
      let borders = await readBorders();
      for (const cell of borders.flat())
        for (const side of ['left', 'right', 'top', 'bottom'])
          assert.deepEqual(cell[side], { color: '#EE9900', widthEmu: 25400, dash: 'dash' });
      await editor.locator('select').first().selectOption('en');
      await fontPanel.getByLabel('Border color', { exact: true }).fill('#0066cc');
      await fontPanel.getByLabel('Border placement', { exact: true }).selectOption('outer');
      await fontPanel.getByRole('button', { name: 'Apply borders', exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      borders = await readBorders();
      assert.equal(borders[0][0].left.color, '#0066CC');
      assert.equal(borders[0][0].right.color, '#0066CC');
      assert.equal(borders[0][0].bottom.color, '#EE9900');
      assert.equal(borders[2][0].top.color, '#EE9900');
      assert.equal(borders[2][0].bottom.color, '#0066CC');
      await fontPanel.getByRole('button', { name: 'Reset borders', exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      for (const cell of (await readBorders()).flat())
        for (const side of ['left', 'right', 'top', 'bottom']) assert.equal(cell[side], null);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.deepEqual(await readBorders(), borders);
      await page.reload();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.deepEqual(await readBorders(), borders);
      await editor.locator('.hit').first().click();
      await editor.locator('select').first().selectOption('ja');
      await editor.getByText('セルの罫線', { exact: true }).click();
      const mergedBox = await editor.locator('.hit').first().boundingBox();
      await editor
        .locator('.hit')
        .first()
        .dblclick({ position: { x: mergedBox.width * 0.8, y: mergedBox.height * 0.2 } });
      const inlineCell = editor.locator('.inline-edit');
      assert.equal(await inlineCell.getAttribute('aria-label'), 'セルのテキスト');
      const mergedText = await inlineCell.inputValue();
      assert.equal(mergedText, 'Hello 日本語!\nB\n追加');
      await inlineCell.fill(mergedText + ' キャンバス');
      await inlineCell.press('Control+Enter');
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      assert.equal(
        getTableCellText(getTableCells(await readTable())[0][0]),
        mergedText + ' キャンバス',
      );
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      assert.equal(getTableCellText(getTableCells(await readTable())[0][0]), mergedText);
      await editor
        .locator('.hit')
        .first()
        .dblclick({ position: { x: mergedBox.width * 0.8, y: mergedBox.height * 0.2 } });
      await inlineCell.fill('取り消す編集');
      await inlineCell.press('Escape');
      assert.equal(
        await editor.getByLabel('セルのテキスト', { exact: true }).inputValue(),
        mergedText,
      );
      await page.screenshot({ path: '/tmp/pptx-pr287-table-ja.png', fullPage: true });
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-pr287-table-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'table insertion chooses dimensions and styles with bilingual dialogs, undo and saved cell editing',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-table-insert-browser-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide} from '@office-kit/pptx-dsl';export default <Presentation><Slide /></Presentation>`,
    );
    let preview, browser, page;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (e) => errors.push(e.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const readShapes = async () =>
        getSlideShapes(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        );
      await editor.locator('select').first().selectOption('ja');
      await editor.getByRole('button', { name: '挿入', exact: true }).click();
      await editor.locator('button[title$="— addSlideTable"]').click();
      const jpDialog = editor.getByRole('dialog', { name: '表を挿入', exact: true });
      await jpDialog.getByLabel('行数', { exact: true }).fill('0');
      assert.equal(
        await jpDialog.getByRole('button', { name: '表を挿入', exact: true }).isDisabled(),
        true,
      );
      await jpDialog.getByLabel('行数', { exact: true }).press('Escape');
      assert.equal(await editor.locator('.hit').count(), 0);
      await editor.locator('button[title$="— addSlideTable"]').click();
      await jpDialog.getByLabel('行数', { exact: true }).fill('3');
      await jpDialog.getByLabel('列数', { exact: true }).fill('2');
      await jpDialog.getByLabel('見出し行', { exact: true }).uncheck();
      await jpDialog.getByLabel('交互の行色', { exact: true }).uncheck();
      await page.screenshot({ path: '/tmp/pptx-pr287-table-insert-ja.png', fullPage: true });
      await jpDialog.getByRole('button', { name: '表を挿入', exact: true }).click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      let shapes = await readShapes();
      assert.equal(shapes.length, 1);
      assert.equal(getTableCells(shapes[0]).length, 3);
      assert.equal(getTableCells(shapes[0])[0].length, 2);
      assert.equal(getTableStyleFlags(shapes[0]).firstRow, false);
      assert.equal(getTableStyleFlags(shapes[0]).bandRow, false);
      assert.equal(
        await editor
          .getByRole('button', { name: 'セル 1, 1', exact: true })
          .getAttribute('aria-pressed'),
        'true',
      );
      await editor.locator('select').first().selectOption('en');
      await editor.locator('button[title$="— addSlideTable"]').click();
      const dialog = editor.getByRole('dialog', { name: 'Insert table', exact: true });
      await dialog.getByLabel('Number of rows', { exact: true }).fill('4');
      await dialog.getByLabel('Number of columns', { exact: true }).fill('3');
      await dialog.getByRole('button', { name: 'Insert table', exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      shapes = await readShapes();
      assert.equal(shapes.length, 2);
      assert.equal(getTableCells(shapes[1]).length, 4);
      assert.equal(getTableCells(shapes[1])[0].length, 3);
      assert.equal(getTableStyleFlags(shapes[1]).firstRow, true);
      assert.equal(getTableStyleFlags(shapes[1]).bandRow, true);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal((await readShapes()).length, 1);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.getByRole('button', { name: 'Cell 4, 3', exact: true }).click();
      await editor.getByLabel('Cell text', { exact: true }).fill('新しい表 / New table');
      await editor.getByLabel('Cell text', { exact: true }).press('Tab');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await page.reload();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      shapes = await readShapes();
      assert.equal(shapes.length, 2);
      assert.equal(getTableCellText(getTableCells(shapes[1])[3][2]), '新しい表 / New table');
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-pr287-table-insert-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'chart dialogs create and edit bilingual data, save workbooks and restore history',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-chart-browser-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide} from '@office-kit/pptx-dsl';export default <Presentation><Slide /></Presentation>`,
    );
    let preview, browser, page;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (e) => errors.push(e.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const download = async () =>
        loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
      const chart = (pres) => getShapeChartSpec(getSlideShapes(getSlides(pres)[0])[0]);
      await saved();
      await editor.locator('select').first().selectOption('ja');
      await editor.getByRole('button', { name: '挿入', exact: true }).click();
      await editor.locator('button[title$="— addSlideChart"]').click();
      const jp = editor.getByRole('dialog', { name: 'グラフを挿入', exact: true });
      await jp.getByLabel('グラフのタイトル', { exact: true }).fill('売上 / Revenue');
      await jp.getByLabel('系列名 1', { exact: true }).fill('日本');
      await jp.getByLabel('項目 1', { exact: true }).fill('春');
      await jp.getByLabel('値 1, 1', { exact: true }).fill('0');
      await jp.getByLabel('値 2, 1', { exact: true }).fill('-2.5');
      await jp.getByLabel('値 3, 1', { exact: true }).fill('');
      await jp.getByRole('button', { name: '項目を追加', exact: true }).click();
      await jp.getByLabel('項目 4', { exact: true }).fill('冬');
      await jp.getByRole('button', { name: '系列を追加', exact: true }).click();
      await jp.getByLabel('系列名 2', { exact: true }).fill('Global');
      await jp.getByLabel('値 4, 2', { exact: true }).fill('42.5');
      await page.screenshot({ path: '/tmp/pptx-pr287-chart-ja.png', fullPage: true });
      await jp.getByRole('button', { name: 'グラフを挿入', exact: true }).click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      let pres = await download();
      assert.equal(chart(pres).kind, 'column');
      assert.equal(chart(pres).title, '売上 / Revenue');
      assert.deepEqual(chart(pres).series[0].values, [0, -2.5, null, 0]);
      assert.deepEqual(chart(pres).series[1].values, [0, 0, 0, 42.5]);
      const bounds = getShapeBoundsResolved(pres, getSlideShapes(getSlides(pres)[0])[0]);
      await editor.locator('select').first().selectOption('en');
      await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
      const dialog = editor.getByRole('dialog', { name: 'Edit chart', exact: true });
      await dialog.getByLabel('Chart type', { exact: true }).selectOption('pie');
      assert.equal(
        await dialog.getByRole('button', { name: 'Apply changes', exact: true }).isDisabled(),
        true,
      );
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.equal(chart(await download()).kind, 'column');
      await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
      await dialog.getByLabel('Chart type', { exact: true }).selectOption('line');
      await dialog.getByLabel('Category 1', { exact: true }).fill('Spring');
      await dialog.getByLabel('Value 1, 1', { exact: true }).fill('123.75');
      await dialog.getByRole('button', { name: 'Remove category 2', exact: true }).click();
      await dialog.getByRole('button', { name: 'Remove series 2', exact: true }).click();
      await dialog.getByLabel('Series color 1', { exact: true }).fill('#d04030');
      await dialog.getByRole('button', { name: 'Apply changes', exact: true }).click();
      await saved();
      pres = await download();
      assert.equal(chart(pres).kind, 'line');
      assert.deepEqual(chart(pres).categories, ['Spring', '項目 3', '冬']);
      assert.deepEqual(chart(pres).series[0].values, [123.75, null, 0]);
      assert.equal(chart(pres).series.length, 1);
      assert.equal(chart(pres).series[0].color.toLowerCase(), '#d04030');
      assert.deepEqual(getShapeBoundsResolved(pres, getSlideShapes(getSlides(pres)[0])[0]), bounds);
      const workbookPath = listPackageParts(pres).find((part) => part.name.endsWith('.xlsx'))?.name;
      assert.ok(workbookPath);
      const workbook = unzipSync(readPackagePart(pres, workbookPath));
      const workbookXml = Object.entries(workbook)
        .filter(([name]) => name.endsWith('.xml'))
        .map(([, bytes]) => strFromU8(bytes))
        .join('\n');
      assert.match(workbookXml, /Spring/);
      assert.match(workbookXml, /123.75/);
      assert.doesNotMatch(workbookXml, /Global|42.5/);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(chart(await download()).kind, 'column');
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.equal(chart(await download()).kind, 'line');
      for (const kind of ['bar', 'area', 'pie', 'doughnut', 'radar']) {
        await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
        await dialog.getByLabel('Chart type', { exact: true }).selectOption(kind);
        await dialog.getByRole('button', { name: 'Apply changes', exact: true }).click();
        await saved();
        assert.equal(chart(await download()).kind, kind);
        assert.match(await editor.locator('.paint').textContent(), /売上/);
      }
      await page.reload();
      await saved();
      pres = await download();
      assert.equal(chart(pres).kind, 'radar');
      assert.deepEqual(chart(pres).series[0].values, [123.75, null, 0]);
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-pr287-chart-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

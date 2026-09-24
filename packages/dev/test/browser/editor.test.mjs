import { installRichTextSelection } from '../helpers/rich-text.mjs';
import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { unzipSync, strFromU8 } from 'fflate';
import {
  getSlideNotes,
  getSlideTransition,
  getSlideSize,
  getSlideBackground,
  getSlideBackgroundImageBytes,
  getSlideLayout,
  getSlideLayoutName,
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
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.hit').first().dblclick();
      await editor.locator('.inline-edit').focus();
      await editor.locator('.inline-edit').evaluate((node) => window.selectEditorText(node, 7, 12));
      await page.keyboard.insertText('headline');
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
      await installRichTextSelection(page);
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
      await installRichTextSelection(page);
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
      await arrange.getByRole('combobox', { name: 'Alignment reference' }).selectOption('slide');
      await arrange.getByRole('button', { name: 'Align center', exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const centered = await readDeck();
      for (const shape of getSlideShapes(getSlides(centered)[0])) {
        const b = getShapeBoundsResolved(centered, shape);
        assert.ok(Math.abs(b.x + b.w / 2 - getSlideSize(centered).width / 2) <= 0.5);
      }
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.lang select').selectOption('ja');
      const japaneseArrange = editor.getByRole('region', { name: '配置', exact: true });
      assert.equal(
        await japaneseArrange.getByRole('combobox', { name: '整列の基準' }).inputValue(),
        'slide',
      );
      await japaneseArrange.getByRole('button', { name: '下揃え', exact: true }).click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      const bottomAligned = await readDeck();
      for (const shape of getSlideShapes(getSlides(bottomAligned)[0])) {
        const b = getShapeBoundsResolved(bottomAligned, shape);
        assert.equal(b.y + b.h, getSlideSize(bottomAligned).height);
      }
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      await japaneseArrange.getByRole('combobox', { name: '整列の基準' }).selectOption('selection');
      await editor.locator('.lang select').selectOption('en');
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
      await installRichTextSelection(page);
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
      await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
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
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.hit').first().dblclick();
      const input = editor.getByRole('textbox', { name: 'Edit text', exact: true });
      await input.focus();
      await input.evaluate((node) =>
        window.selectEditorText(
          node,
          (node.textContent ?? node.value).length,
          (node.textContent ?? node.value).length,
        ),
      );
      await page.keyboard.insertText('!');
      await input.evaluate((node) => {
        window.selectEditorText(node, 6, 9);
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
          window.selectEditorText(node, 6, 9);
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
      await installRichTextSelection(page);
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
      await editor.locator('.inline-edit').focus();
      await editor
        .locator('.inline-edit')
        .evaluate((node) =>
          window.selectEditorText(
            node,
            (node.textContent ?? node.value).length,
            (node.textContent ?? node.value).length,
          ),
        );
      await page.keyboard.insertText('!');
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
        window.selectEditorText(node, 6, 9);
        node.dispatchEvent(new Event('select', { bubbles: true }));
      });
      await editor.locator('.inline-edit').press('Control+u');
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

      await fontPanel.getByRole('button', { name: 'Cell 1, 1', exact: true }).click();
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
      const mergedText = await inlineCell.textContent();
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
      await inlineCell.fill('Escapeで確定する編集');
      await inlineCell.press('Escape');
      assert.equal(
        await editor.getByLabel('セルのテキスト', { exact: true }).inputValue(),
        'Escapeで確定する編集',
      );
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      assert.equal(
        getTableCellText(getTableCells(await readTable())[0][0]),
        'Escapeで確定する編集',
      );
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      assert.equal(getTableCellText(getTableCells(await readTable())[0][0]), mergedText);
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
      await installRichTextSelection(page);
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
      await editor.getByRole('tab', { name: '挿入', exact: true }).click();
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
      await installRichTextSelection(page);
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
      await editor.getByRole('tab', { name: '挿入', exact: true }).click();
      await editor.locator('button[title$="— addSlideChart"]').click();
      const jp = editor.getByRole('dialog', { name: 'グラフを挿入', exact: true });
      await jp.getByLabel('グラフのタイトル', { exact: true }).fill('売上 / Revenue');
      await jp.getByLabel('凡例', { exact: true }).selectOption('b');
      await jp.getByLabel('値を表示', { exact: true }).check();
      await jp.getByLabel('項目名を表示', { exact: true }).check();
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
      assert.equal(chart(pres).legend.position, 'b');
      assert.equal(chart(pres).dataLabels.showValue, true);
      assert.equal(chart(pres).dataLabels.showCategory, true);
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
      await dialog.getByLabel('Legend', { exact: true }).selectOption('r');
      await dialog.getByLabel('Show categories', { exact: true }).uncheck();
      await dialog.getByLabel('Show series names', { exact: true }).check();
      await dialog.getByLabel('Category 1', { exact: true }).fill('Spring');
      await dialog.getByLabel('Value 1, 1', { exact: true }).fill('123.75');
      await dialog.getByRole('button', { name: 'Remove category 2', exact: true }).click();
      await dialog.getByRole('button', { name: 'Remove series 2', exact: true }).click();
      await dialog.getByLabel('Series color 1', { exact: true }).fill('#d04030');
      await dialog.getByRole('button', { name: 'Apply changes', exact: true }).click();
      await saved();
      pres = await download();
      assert.equal(chart(pres).kind, 'line');
      assert.equal(chart(pres).legend.position, 'r');
      assert.equal(chart(pres).dataLabels.showCategory, false);
      assert.equal(chart(pres).dataLabels.showSeriesName, true);
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
        if (kind === 'pie') await dialog.getByLabel('Show percentages', { exact: true }).check();
        if (kind === 'radar') {
          await dialog.getByLabel('Show percentages', { exact: true }).uncheck();
          await dialog.getByLabel('Legend', { exact: true }).selectOption({ label: 'None' });
        }
        await dialog.getByRole('button', { name: 'Apply changes', exact: true }).click();
        await saved();
        const current = chart(await download());
        assert.equal(current.kind, kind);
        if (kind === 'pie') assert.equal(current.dataLabels.showPercent, true);
        if (kind === 'radar') {
          assert.equal(current.legend?.position ?? null, null);
          assert.equal(current.dataLabels.showPercent, false);
        }
        assert.match(await editor.locator('.paint').textContent(), /売上/);
      }
      await page.reload();
      await saved();
      pres = await download();
      assert.equal(chart(pres).kind, 'radar');
      assert.equal(chart(pres).legend?.position ?? null, null);
      assert.equal(chart(pres).dataLabels.showValue, true);
      assert.equal(chart(pres).dataLabels.showSeriesName, true);
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

test(
  'slide options edit layouts and backgrounds in both languages with saved history',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-slide-options-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={4} height={1}>Keep this text</Text></Slide><Slide /></Presentation>`,
    );
    let preview, browser, page;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const slides = async () =>
        getSlides(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
        );
      await saved();
      assert.equal(getSlideText((await slides())[0]), 'Keep this text');
      const pane = editor.getByRole('region', { name: 'Slide options', exact: true });
      await pane
        .getByLabel('Slide layout', { exact: true })
        .selectOption({ label: 'Title and Content' });
      await saved();
      assert.equal(getSlideLayoutName(getSlideLayout((await slides())[0])), 'Title and Content');
      await pane.getByLabel('Background color', { exact: true }).fill('#d8ebff');
      await saved();
      assert.deepEqual(getSlideBackground((await slides())[0]), {
        kind: 'solid',
        color: '#D8EBFF',
      });
      assert.equal(getSlideBackground((await slides())[1]).kind, 'inherit');
      await pane.getByRole('button', { name: 'Reset background', exact: true }).click();
      await saved();
      assert.equal(getSlideBackground((await slides())[0]).kind, 'inherit');
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(getSlideBackground((await slides())[0]).kind, 'solid');
      await editor.locator('select').first().selectOption('ja');
      const jp = editor.getByRole('region', { name: 'スライドの設定', exact: true });
      const bytes = Buffer.from(
        await page.evaluate(() => {
          const canvas = document.createElement('canvas');
          canvas.width = 20;
          canvas.height = 20;
          const ctx = canvas.getContext('2d');
          ctx.fillStyle = '#d04030';
          ctx.fillRect(0, 0, 20, 20);
          return canvas.toDataURL('image/png').split(',')[1];
        }),
        'base64',
      );
      await jp
        .getByLabel('背景画像', { exact: true })
        .setInputFiles({ name: 'background.png', mimeType: 'image/png', buffer: bytes });
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      assert.deepEqual(Buffer.from(getSlideBackgroundImageBytes((await slides())[0])), bytes);
      await jp
        .getByLabel('スライドのレイアウト', { exact: true })
        .selectOption({ label: 'タイトルスライド' });
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      await page.screenshot({ path: '/tmp/pptx-pr287-slide-options-ja.png', fullPage: true });
      await page.reload();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      const result = await slides();
      assert.equal(getSlideLayoutName(getSlideLayout(result[0])), 'Title Slide');
      assert.deepEqual(Buffer.from(getSlideBackgroundImageBytes(result[0])), bytes);
      assert.equal(getSlideText(result[0]), 'Keep this text');
      assert.equal(getSlideBackground(result[1]).kind, 'inherit');
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-pr287-slide-options-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'page setup saves standard and custom slide sizes in both languages with undo',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-page-setup-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text,Table} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={4} height={1}>Size test</Text></Slide><Slide><Table x={1} y={1} width={4} height={2} rows={[["Default table text"]]} /></Slide></Presentation>`,
    );
    let preview, browser, page;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const read = async () =>
        loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
      const tableFontSize = (deck) =>
        getTableCellParagraphs(getTableCells(getSlideShapes(getSlides(deck)[1])[0])[0][0])[0]
          .elements[0].format?.size;
      await saved();
      assert.equal(tableFontSize(await read()), undefined);
      const original = await read();
      const bounds = getShapeBoundsResolved(original, getSlideShapes(getSlides(original)[0])[0]);
      await editor.getByRole('button', { name: 'Page setup', exact: true }).click();
      let dialog = editor.getByRole('dialog', { name: 'Page setup', exact: true });
      await dialog.getByLabel('Slide size', { exact: true }).selectOption('standard');
      await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
      await saved();
      assert.deepEqual(getSlideSize(await read()), {
        width: 9144000,
        height: 6858000,
        type: 'screen4x3',
      });
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(getSlideSize(await read()), getSlideSize(original));
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.equal(getSlideSize(await read()).width, 9144000);
      await editor.getByRole('button', { name: 'Page setup', exact: true }).click();
      await dialog.getByLabel('Slide size', { exact: true }).selectOption('wide10');
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.equal(getSlideSize(await read()).width, 9144000);
      for (const [id, expected] of [
        ['wide10', { width: 12192000, height: 7620000, type: 'screen16x10' }],
        ['wide', { width: 12192000, height: 6858000, type: 'screen16x9' }],
        ['standard', { width: 9144000, height: 6858000, type: 'screen4x3' }],
      ]) {
        await editor.getByRole('button', { name: 'Page setup', exact: true }).click();
        await dialog.getByLabel('Slide size', { exact: true }).selectOption(id);
        await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
        await saved();
        assert.deepEqual(getSlideSize(await read()), expected);
        const stage = await editor.locator('.stage').boundingBox();
        assert.ok(Math.abs(stage.width / stage.height - expected.width / expected.height) < 0.001);
      }
      await editor.locator('select').first().selectOption('ja');
      await editor.getByRole('button', { name: 'ページ設定', exact: true }).click();
      dialog = editor.getByRole('dialog', { name: 'ページ設定', exact: true });
      await dialog.getByLabel('単位', { exact: true }).selectOption('cm');
      assert.equal(
        Number(await dialog.getByLabel('ページの幅', { exact: true }).inputValue()),
        25.4,
      );
      await dialog.getByLabel('ページの幅', { exact: true }).fill('0');
      assert.equal(
        await dialog.getByRole('button', { name: '適用', exact: true }).isDisabled(),
        true,
      );
      assert.equal(
        await dialog.getByRole('alert').textContent(),
        '幅と高さは2.54〜142.24 cmで入力してください。',
      );
      for (const value of ['0.1', '2.53']) {
        await dialog.getByLabel('ページの幅', { exact: true }).fill(value);
        assert.equal(
          await dialog.getByRole('button', { name: '適用', exact: true }).isDisabled(),
          true,
        );
      }
      await dialog.getByLabel('ページの幅', { exact: true }).fill('2.54');
      assert.equal(
        await dialog.getByRole('button', { name: '適用', exact: true }).isDisabled(),
        false,
      );
      await dialog.getByLabel('ページの幅', { exact: true }).fill('200');
      assert.equal(
        await dialog.getByRole('button', { name: '適用', exact: true }).isDisabled(),
        true,
      );
      await dialog.getByLabel('ページの幅', { exact: true }).fill('');
      assert.equal(
        await dialog.getByRole('button', { name: '適用', exact: true }).isDisabled(),
        true,
      );
      await dialog.getByLabel('ページの幅', { exact: true }).fill('21');
      await dialog.getByLabel('ページの高さ', { exact: true }).fill('29.7');
      assert.equal(
        await dialog.getByLabel('スライドのサイズ', { exact: true }).inputValue(),
        'custom',
      );
      await page.screenshot({ path: '/tmp/pptx-pr287-page-setup-ja.png', fullPage: true });
      await dialog.getByRole('button', { name: '適用', exact: true }).click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      await page.reload();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      const result = await read();
      assert.deepEqual(getSlideSize(result), { width: 7560000, height: 10692000 });
      assert.equal(getSlides(result).length, 2);
      assert.equal(getSlideText(getSlides(result)[0]), 'Size test');
      assert.deepEqual(
        getShapeBoundsResolved(result, getSlideShapes(getSlides(result)[0])[0]),
        bounds,
      );
      await editor.getByRole('button', { name: 'ページ設定', exact: true }).click();
      await dialog.getByLabel('単位', { exact: true }).selectOption('cm');
      assert.equal(Number(await dialog.getByLabel('ページの幅', { exact: true }).inputValue()), 21);
      assert.equal(
        Number(await dialog.getByLabel('ページの高さ', { exact: true }).inputValue()),
        29.7,
      );
      await dialog.getByRole('button', { name: 'キャンセル', exact: true }).click();
      for (const language of ['ja', 'en']) {
        await editor.locator('select').first().selectOption(language);
        const ja = language === 'ja';
        const targetInches = ja ? 10 : 12;
        const target = targetInches * 914400;
        const beforeFit = await read();
        const beforeSize = getSlideSize(beforeFit);
        const beforeBounds = getShapeBoundsResolved(
          beforeFit,
          getSlideShapes(getSlides(beforeFit)[0])[0],
        );
        await editor
          .getByRole('button', { name: ja ? 'ページ設定' : 'Page setup', exact: true })
          .click();
        const fitDialog = editor.getByRole('dialog', {
          name: ja ? 'ページ設定' : 'Page setup',
          exact: true,
        });
        await fitDialog
          .getByLabel(ja ? 'ページの幅' : 'Page width', { exact: true })
          .fill(String(targetInches));
        await fitDialog
          .getByLabel(ja ? 'ページの高さ' : 'Page height', { exact: true })
          .fill(String(targetInches));
        await fitDialog
          .getByLabel(ja ? 'サイズ変更時の処理' : 'When resizing', { exact: true })
          .selectOption('fit');
        if (ja) await page.screenshot({ path: '/tmp/pptx-pr287-page-fit-ja.png', fullPage: true });
        await fitDialog.getByRole('button', { name: ja ? '適用' : 'Apply', exact: true }).click();
        const waitSaved = () =>
          editor
            .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
            .waitFor();
        await waitSaved();
        const fitted = await read();
        const scale = Math.min(target / beforeSize.width, target / beforeSize.height);
        const authoredSizes = (deck) =>
          [...getSlideXmlString(getSlides(deck)[0]).matchAll(/\bsz="(\d+)"/g)].map((match) =>
            Number(match[1]),
          );
        assert.deepEqual(
          authoredSizes(fitted),
          authoredSizes(beforeFit).map((size) => Math.round(size * scale)),
        );
        const expectedTableSize = Math.round((tableFontSize(beforeFit) ?? 18) * 100 * scale) / 100;
        assert.equal(tableFontSize(fitted), expectedTableSize);
        const expected = {
          x: Math.round(beforeBounds.x * scale + (target - beforeSize.width * scale) / 2),
          y: Math.round(beforeBounds.y * scale + (target - beforeSize.height * scale) / 2),
          w: Math.round(beforeBounds.w * scale),
          h: Math.round(beforeBounds.h * scale),
        };
        assert.deepEqual(
          getShapeBoundsResolved(fitted, getSlideShapes(getSlides(fitted)[0])[0]),
          expected,
        );
        await editor
          .getByTitle(ja ? '元に戻す (Ctrl+Z)' : 'Undo (Ctrl+Z)', { exact: true })
          .click();
        await waitSaved();
        assert.deepEqual(getSlideSize(await read()), beforeSize);
        assert.equal(tableFontSize(await read()), tableFontSize(beforeFit));
        await editor
          .getByTitle(ja ? 'やり直し (Ctrl+Y)' : 'Redo (Ctrl+Y)', { exact: true })
          .click();
        await waitSaved();
        await page.reload();
        await waitSaved();
        const reloaded = await read();
        assert.equal(tableFontSize(reloaded), expectedTableSize);
        assert.deepEqual(
          getShapeBoundsResolved(reloaded, getSlideShapes(getSlides(reloaded)[0])[0]),
          expected,
        );
      }
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-pr287-page-setup-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'speaker notes and slide transitions edit selected slides and persist bilingual history',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-notes-transitions-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide} from '@office-kit/pptx-dsl';export default <Presentation><Slide /><Slide /></Presentation>`,
    );
    let preview, browser, page;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      let locale = 'en';
      const saved = () =>
        editor
          .getByText(locale === 'en' ? 'Saved to this project' : 'このプロジェクトに保存済み', {
            exact: true,
          })
          .waitFor();
      const slides = async () =>
        getSlides(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
        );
      await saved();
      await editor.getByRole('button', { name: 'Speaker notes', exact: true }).click();
      let notesInput = editor.getByLabel('Notes content', { exact: true });
      const notes = 'Opening remarks\n日本語の説明 🎉\n\nFinal point';
      await notesInput.fill(notes);
      await notesInput.press('Tab');
      await saved();
      assert.equal(getSlideNotes((await slides())[0]), notes);
      assert.equal(getSlideNotes((await slides())[1]), null);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(getSlideNotes((await slides())[0]), null);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      await editor.getByRole('button', { name: 'Speaker notes', exact: true }).click();
      assert.equal(await notesInput.inputValue(), notes);
      await notesInput.fill('Undo this edit');
      await notesInput.press('Meta+z');
      await saved();
      assert.equal(getSlideNotes((await slides())[0]), notes);
      await editor.locator('.thumb-row').nth(1).click();
      await editor.locator('.lang select').selectOption('ja');
      locale = 'ja';
      await editor.getByRole('button', { name: '発表者ノート', exact: true }).click();
      notesInput = editor.getByLabel('ノートの内容', { exact: true });
      assert.equal(await notesInput.inputValue(), '');
      await notesInput.fill('次のスライドの説明');
      await page.screenshot({ path: '/tmp/pptx-pr287-notes-ja.png', fullPage: true });
      await notesInput.press('Tab');
      await saved();
      await editor.getByRole('button', { name: 'スライドの画面切り替え', exact: true }).click();
      let dialog = editor.getByRole('dialog', { name: 'スライドの画面切り替え', exact: true });
      await dialog.getByLabel('切り替え効果', { exact: true }).selectOption('push');
      await dialog.getByLabel('切り替え速度', { exact: true }).selectOption('fast');
      await dialog.getByLabel('切り替え方向', { exact: true }).selectOption('r');
      await dialog.getByLabel('クリックで次に進む', { exact: true }).uncheck();
      await dialog.getByLabel('自動で次に進む', { exact: true }).check();
      await dialog.getByLabel('次に進むまでの秒数', { exact: true }).fill('-1');
      assert.equal(
        await dialog.getByRole('button', { name: '適用', exact: true }).isDisabled(),
        true,
      );
      await dialog.getByLabel('次に進むまでの秒数', { exact: true }).fill('2.5');
      await dialog.getByLabel('すべてのスライドに適用', { exact: true }).check();
      await page.screenshot({ path: '/tmp/pptx-pr287-transition-ja.png', fullPage: true });
      await dialog.getByRole('button', { name: '適用', exact: true }).click();
      await saved();
      const push = {
        effect: 'push',
        speed: 'fast',
        direction: 'r',
        advanceOnClick: false,
        advanceAfterMs: 2500,
      };
      for (const slide of await slides()) assert.deepEqual(getSlideTransition(slide), push);
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      for (const slide of await slides()) assert.equal(getSlideTransition(slide), null);
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      await editor.locator('.thumb-row').first().click();
      await editor.locator('.lang select').selectOption('en');
      locale = 'en';
      await editor.getByRole('button', { name: 'Slide transition', exact: true }).click();
      dialog = editor.getByRole('dialog', { name: 'Slide transition', exact: true });
      assert.equal(
        await dialog.getByLabel('Advance after (seconds)', { exact: true }).inputValue(),
        '2.5',
      );
      await dialog.getByLabel('Transition effect', { exact: true }).selectOption('fade');
      await dialog.getByLabel('Through black', { exact: true }).check();
      await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
      await saved();
      assert.deepEqual(getSlideTransition((await slides())[0]), {
        effect: 'fade',
        speed: 'fast',
        thruBlack: true,
        advanceOnClick: false,
        advanceAfterMs: 2500,
      });
      assert.deepEqual(getSlideTransition((await slides())[1]), push);
      await editor.getByRole('button', { name: 'Slide transition', exact: true }).click();
      await dialog.getByLabel('Transition effect', { exact: true }).selectOption('split');
      await dialog.getByLabel('Transition direction', { exact: true }).selectOption('out');
      await dialog.getByLabel('Split orientation', { exact: true }).selectOption('vert');
      await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
      await saved();
      assert.deepEqual(getSlideTransition((await slides())[0]), {
        effect: 'split',
        speed: 'fast',
        direction: 'out',
        orientation: 'vert',
        advanceOnClick: false,
        advanceAfterMs: 2500,
      });
      await editor.getByRole('button', { name: 'Slide transition', exact: true }).click();
      await dialog.getByLabel('Transition effect', { exact: true }).selectOption('wheel');
      await dialog.getByLabel('Wheel spokes', { exact: true }).fill('1.5');
      assert.equal(
        await dialog.getByRole('button', { name: 'Apply', exact: true }).isDisabled(),
        true,
      );
      await dialog.getByLabel('Wheel spokes', { exact: true }).fill('8');
      await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
      await saved();
      assert.equal(getSlideTransition((await slides())[0]).spokes, 8);
      await editor.locator('.lang select').selectOption('ja');
      locale = 'ja';
      await editor.getByRole('button', { name: 'スライドの画面切り替え', exact: true }).click();
      const wheelDialog = editor.getByRole('dialog', {
        name: 'スライドの画面切り替え',
        exact: true,
      });
      assert.equal(
        await wheelDialog.getByLabel('ホイールの本数', { exact: true }).inputValue(),
        '8',
      );
      await wheelDialog.getByLabel('ホイールの本数', { exact: true }).fill('3');
      await wheelDialog.getByRole('button', { name: '適用', exact: true }).click();
      await saved();
      assert.equal(getSlideTransition((await slides())[0]).spokes, 3);
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(getSlideTransition((await slides())[0]).spokes, 8);
      await editor.locator('.lang select').selectOption('en');
      locale = 'en';
      await editor.getByRole('button', { name: 'Slide transition', exact: true }).click();
      await dialog.getByLabel('Transition effect', { exact: true }).selectOption('none');
      await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
      await saved();
      assert.deepEqual(getSlideTransition((await slides())[0]), {
        effect: 'none',
        advanceOnClick: false,
        advanceAfterMs: 2500,
      });
      await editor.getByRole('button', { name: 'Slide transition', exact: true }).click();
      assert.equal(
        await dialog.getByLabel('Advance automatically', { exact: true }).isChecked(),
        true,
      );
      await dialog.getByLabel('Advance automatically', { exact: true }).uncheck();
      await dialog.getByLabel('Advance on click', { exact: true }).check();
      await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      const result = await slides();
      assert.equal(getSlideTransition(result[0]), null);
      assert.deepEqual(getSlideTransition(result[1]), push);
      assert.equal(getSlideNotes(result[0]), notes);
      assert.equal(getSlideNotes(result[1]), '次のスライドの説明');
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({
        path: '/tmp/pptx-pr287-notes-transition-failure.png',
        fullPage: true,
      });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

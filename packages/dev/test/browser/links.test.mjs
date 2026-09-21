import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeClickAction,
  getSlideIndex,
  getShapeKind,
  getSlides,
  getSlideShapes,
  getShapeHyperlink,
  getShapeHyperlinkTooltip,
  getShapeParagraphElements,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'link dialog adds, edits, cancels and removes links with bilingual batch undo and reload',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-links-'));
    let preview, browser, page;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={4} height={1} bold>First link</Text><Text x={1} y={3} width={4} height={1}>Second link</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (e) => errors.push(e.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      let ja = false;
      const saved = () =>
        editor
          .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
          .waitFor();
      const read = async () =>
        getSlideShapes(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        );
      const open = async () => {
        await editor.getByRole('button', { name: ja ? '挿入' : 'Insert', exact: true }).click();
        await editor.locator('button[title$="— setShapeHyperlink"]').click();
        return editor.getByRole('dialog', { name: ja ? 'リンクを編集' : 'Edit link', exact: true });
      };
      await saved();
      await editor.locator('.hit').nth(0).click();
      let dialog = await open();
      await dialog.getByLabel('Link address', { exact: true }).fill('https://example.com/first');
      await dialog.getByLabel('Link description', { exact: true }).fill('First destination');
      await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
      await saved();
      assert.equal(getShapeHyperlink((await read())[0]), 'https://example.com/first');
      dialog = await open();
      assert.equal(
        await dialog.getByLabel('Link address', { exact: true }).inputValue(),
        'https://example.com/first',
      );
      assert.equal(
        await dialog.getByLabel('Link description', { exact: true }).inputValue(),
        'First destination',
      );
      await dialog
        .getByLabel('Link address', { exact: true })
        .fill('https://example.com/cancelled');
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.equal(getShapeHyperlink((await read())[0]), 'https://example.com/first');
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      await editor.locator('select').first().selectOption('ja');
      ja = true;
      dialog = await open();
      await dialog
        .getByText('選択した図形には異なるリンクが設定されています。', { exact: true })
        .waitFor();
      await dialog.getByLabel('リンク先', { exact: true }).fill('https://example.com/shared');
      await dialog.getByLabel('リンクの説明', { exact: true }).fill('共通の資料');
      await page.screenshot({ path: '/tmp/pptx-pr287-link-ja.png', fullPage: true });
      await dialog.getByRole('button', { name: '適用', exact: true }).click();
      await saved();
      assert.deepEqual((await read()).map(getShapeHyperlink), [
        'https://example.com/shared',
        'https://example.com/shared',
      ]);
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual((await read()).map(getShapeHyperlink), ['https://example.com/first', null]);
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      assert.deepEqual((await read()).map(getShapeHyperlinkTooltip), ['共通の資料', '共通の資料']);
      assert.equal(getShapeParagraphElements((await read())[0], 0)[0].format.bold, true);
      await editor.locator('.hit').nth(0).click();
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      dialog = await open();
      await dialog.getByRole('button', { name: 'リンクを解除', exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      assert.deepEqual((await read()).map(getShapeHyperlink), [null, null]);
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-pr287-link-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'image and text links switch between web and slide destinations with undo and reload',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-object-links-'));
    let preview, browser, page;
    const errors = [];
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={4} height={1} bold>Link text</Text></Slide><Slide><Text x={1} y={1} width={4} height={1}>Destination</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      let ja = false;
      const saved = () =>
        editor
          .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
          .waitFor();
      const read = async () =>
        loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
      const open = async () => {
        await editor.getByRole('button', { name: ja ? '挿入' : 'Insert', exact: true }).click();
        await editor.locator('button[title$="— setShapeHyperlink"]').click();
        return editor.getByRole('dialog', { name: ja ? 'リンクを編集' : 'Edit link', exact: true });
      };
      await saved();
      await editor.getByRole('button', { name: 'Insert', exact: true }).click();
      await editor.getByTitle(/— addSlideImage$/).click();
      let dialog = editor.getByRole('dialog');
      const png = Buffer.from(
        await page.evaluate(() => {
          const canvas = document.createElement('canvas');
          canvas.width = 100;
          canvas.height = 60;
          const ctx = canvas.getContext('2d');
          ctx.fillStyle = '#387cdc';
          ctx.fillRect(0, 0, 100, 60);
          return canvas.toDataURL('image/png').split(',')[1];
        }),
        'base64',
      );
      await dialog
        .getByLabel('Image file', { exact: true })
        .setInputFiles({ name: 'link.png', mimeType: 'image/png', buffer: png });
      await dialog.getByRole('button', { name: 'Insert image', exact: true }).click();
      await saved();
      dialog = await open();
      await dialog.getByLabel('Link address', { exact: true }).fill('https://example.com/image');
      await dialog.getByLabel('Link description', { exact: true }).fill('Image reference');
      await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
      await saved();
      let pres = await read();
      const picture = (p) =>
        getSlideShapes(getSlides(p)[0]).find((s) => getShapeKind(s) === 'picture');
      assert.deepEqual(getShapeClickAction(picture(pres)), {
        kind: 'url',
        url: 'https://example.com/image',
      });
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      dialog = await open();
      assert.equal(
        await dialog.getByLabel('リンク先', { exact: true }).inputValue(),
        'https://example.com/image',
      );
      assert.equal(
        await dialog.getByLabel('リンクの説明', { exact: true }).inputValue(),
        'Image reference',
      );
      await dialog.getByLabel('リンクの説明', { exact: true }).fill('画像の参考資料');
      await dialog.getByLabel('リンク先の種類', { exact: true }).selectOption('slide');
      await dialog.getByLabel('移動先のスライド', { exact: true }).selectOption('1');
      await page.screenshot({ path: '/tmp/pptx-pr287-object-link-ja.png', fullPage: true });
      await dialog.getByRole('button', { name: '適用', exact: true }).click();
      await saved();
      pres = await read();
      assert.equal(getSlideIndex(pres, getShapeClickAction(picture(pres)).slide), 1);
      const viewer = await browser.newPage();
      viewer.on('pageerror', (error) => errors.push(error.message));
      await viewer.goto(preview.url);
      await viewer.getByRole('button', { name: 'Preview', exact: true }).click();
      const imageLink = viewer.locator('#slide a').filter({ has: viewer.locator('image') });
      assert.equal(await imageLink.locator('title').textContent(), '画像の参考資料');
      await imageLink.click();
      assert.equal(await viewer.locator('#count').textContent(), 'Slide 2 of 2');
      await viewer.close();

      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(getShapeClickAction(picture(await read())).url, 'https://example.com/image');
      assert.equal(getShapeHyperlinkTooltip(picture(await read())), 'Image reference');
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      await editor.locator('.hit').nth(1).click();
      dialog = await open();
      assert.equal(
        await dialog.getByLabel('リンク先の種類', { exact: true }).inputValue(),
        'slide',
      );
      assert.equal(await dialog.getByLabel('移動先のスライド', { exact: true }).inputValue(), '1');
      assert.equal(
        await dialog.getByLabel('リンクの説明', { exact: true }).inputValue(),
        '画像の参考資料',
      );
      await dialog.getByLabel('リンクの説明', { exact: true }).fill('');
      await dialog.getByRole('button', { name: '適用', exact: true }).click();
      await saved();
      assert.equal(getShapeHyperlinkTooltip(picture(await read())), null);
      dialog = await open();
      await dialog.getByRole('button', { name: 'リンクを解除', exact: true }).click();
      await saved();
      assert.equal(getShapeClickAction(picture(await read())), null);
      await editor
        .locator('.hit')
        .nth(0)
        .click({ position: { x: 5, y: 5 } });
      dialog = await open();
      await dialog.getByLabel('リンク先の種類', { exact: true }).selectOption('slide');
      await dialog.getByLabel('移動先のスライド', { exact: true }).selectOption('1');
      await dialog.getByRole('button', { name: '適用', exact: true }).click();
      await saved();
      dialog = await open();
      await dialog.getByLabel('リンク先の種類', { exact: true }).selectOption('url');
      await dialog.getByLabel('リンク先', { exact: true }).fill('https://example.com/text');
      await dialog.getByRole('button', { name: '適用', exact: true }).click();
      await saved();
      pres = await read();
      const text = getSlideShapes(getSlides(pres)[0])[0];
      assert.equal(getShapeClickAction(text), null);
      assert.equal(getShapeHyperlink(text), 'https://example.com/text');
      assert.equal(getShapeParagraphElements(text, 0)[0].format.bold, true);
      for (const kind of ['nextSlide', 'prevSlide', 'firstSlide', 'lastSlide']) {
        dialog = await open();
        await dialog.getByLabel('リンク先の種類', { exact: true }).selectOption(kind);
        await dialog.getByRole('button', { name: '適用', exact: true }).click();
        await saved();
        const updated = getSlideShapes(getSlides(await read())[0])[0];
        assert.deepEqual(getShapeClickAction(updated), { kind });
        assert.equal(getShapeHyperlink(updated), null);
        assert.equal(getShapeParagraphElements(updated, 0)[0].format.bold, true);
      }
      await page.reload();
      await saved();
      await editor
        .locator('.hit')
        .nth(0)
        .click({ position: { x: 5, y: 5 } });
      await editor.locator('.lang select').selectOption('en');
      ja = false;
      dialog = await open();
      assert.equal(
        await dialog.getByLabel('Link destination', { exact: true }).inputValue(),
        'lastSlide',
      );
      await dialog.getByLabel('Link destination', { exact: true }).selectOption('nextSlide');
      await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
      await saved();
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(
        getShapeClickAction(getSlideShapes(getSlides(await read())[0])[0]).kind,
        'lastSlide',
      );
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-pr287-object-link-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

for (const target of ['shape', 'cell'])
  test(
    `selected ${target} text links preserve surrounding text with Japanese undo and reload`,
    { timeout: 60000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-range-links-'));
      let preview, browser;
      try {
        const file = join(dir, 'deck.tsx');
        await writeFile(
          file,
          target === 'cell'
            ? `import {Presentation,Slide,Table,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Table x={1} y={1} width={8} height={2} rows={[["Before 日本語 After", "Untouched"]]} /></Slide><Slide><Text x={1} y={1} width={5} height={1}>Target</Text></Slide></Presentation>`
            : `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={6} height={1} bold>Before 日本語 After</Text></Slide><Slide><Text x={1} y={1} width={5} height={1}>Target</Text></Slide></Presentation>`,
        );
        preview = await startPreview(file);
        browser = await chromium.launch({ headless: true });
        const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
        const errors = [];
        page.on('pageerror', (e) => errors.push(e.message));
        await page.goto(preview.url);
        await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
        const editor = page.frameLocator('#editor-frame');
        let ja = false;
        const saved = () =>
          editor
            .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
            .waitFor();
        const read = async () =>
          getSlideShapes(
            getSlides(
              await loadPresentation(
                new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
              ),
            )[0],
          )[0];
        const open = async (keyboard = false) => {
          const hit = editor.locator('.hit').first();
          const bounds = await hit.boundingBox();
          await hit.dblclick(
            target === 'cell' ? { position: { x: bounds.width / 4, y: bounds.height / 2 } } : {},
          );
          const input = editor.locator('textarea.inline-edit');
          await input.evaluate((el) => {
            el.focus();
            el.setSelectionRange(7, 10);
            el.dispatchEvent(new Event('select', { bubbles: true }));
          });
          if (keyboard) await input.press('Control+k');
          else
            await editor
              .locator('.canvas-shell .text-format-bar')
              .getByRole('button', { name: ja ? 'リンクを編集' : 'Edit link', exact: true })
              .click();
          return editor.getByRole('dialog', {
            name: ja ? 'リンクを編集' : 'Edit link',
            exact: true,
          });
        };
        await saved();
        let dialog = await open();
        await dialog.getByText('Applies to the selected text.', { exact: true }).waitFor();
        await dialog
          .getByLabel('Link address', { exact: true })
          .fill('https://example.com/japanese');
        await dialog.getByLabel('Link description', { exact: true }).fill('日本語の資料');
        await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
        await saved();
        const api = await import('@office-kit/pptx');
        const paragraphs = (shape) =>
          target === 'cell'
            ? api.getTableCellParagraphs(api.getTableCell(shape, 0, 0))[0].elements
            : getShapeParagraphElements(shape, 0);
        const getAction = (shape, _p, r) =>
          target === 'cell'
            ? (paragraphs(shape)[r].clickAction ?? null)
            : api.getShapeRunClickAction(shape, 0, r);
        const getLink = (shape, p, r) => {
          const action = getAction(shape, p, r);
          return action?.kind === 'url' ? action.url : null;
        };
        const getTip = (shape, _p, r) =>
          target === 'cell'
            ? (paragraphs(shape)[r].tooltip ?? null)
            : api.getShapeRunHyperlinkTooltip(shape, 0, r);
        const checkFormat = (shape) => {
          if (target === 'shape') assert.ok(paragraphs(shape).every((e) => e.format.bold));
          else assert.equal(api.getTableCellText(api.getTableCell(shape, 0, 1)), 'Untouched');
        };
        let shape = await read();
        assert.deepEqual(
          paragraphs(shape).map((e) => e.text),
          ['Before ', '日本語', ' After'],
        );
        assert.deepEqual(
          [0, 1, 2].map((r) => getLink(shape, 0, r)),
          [null, 'https://example.com/japanese', null],
        );
        checkFormat(shape);
        await editor.locator('select').first().selectOption('ja');
        ja = true;
        dialog = await open(true);
        assert.equal(
          await dialog.getByLabel('リンク先', { exact: true }).inputValue(),
          'https://example.com/japanese',
        );
        await dialog.getByLabel('リンク先', { exact: true }).fill('https://example.com/cancelled');
        await dialog.getByRole('button', { name: 'キャンセル', exact: true }).click();
        assert.equal(getLink(await read(), 0, 1), 'https://example.com/japanese');
        dialog = await open();
        await dialog.getByRole('button', { name: 'リンクを解除', exact: true }).click();
        await saved();
        assert.equal(getLink(await read(), 0, 1), null);
        await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
        await saved();
        await page.reload();
        await saved();
        shape = await read();
        assert.equal(getLink(shape, 0, 1), 'https://example.com/japanese');
        assert.equal(getTip(shape, 0, 1), '日本語の資料');
        dialog = await open(true);
        await dialog.getByLabel('リンク先の種類', { exact: true }).selectOption('slide');
        await dialog.getByLabel('移動先のスライド', { exact: true }).selectOption('1');
        if (target === 'cell')
          await page.screenshot({ path: '/tmp/pptx-pr287-cell-link-ja.png', fullPage: true });
        await dialog.getByRole('button', { name: '適用', exact: true }).click();
        await saved();
        assert.equal(getAction(await read(), 0, 1)?.kind, 'slide');
        const viewer = await browser.newPage();
        viewer.on('pageerror', (error) => errors.push(error.message));
        await viewer.goto(preview.url);
        await viewer.getByRole('button', { name: 'Preview', exact: true }).click();
        const link = viewer.locator('#slide a').filter({ hasText: '日本語' });
        assert.equal(await link.count(), 1);
        await link.click();
        assert.equal(await viewer.locator('#count').textContent(), 'Slide 2 of 2');
        await viewer.close();
        await page.reload();
        await saved();
        dialog = await open();
        assert.equal(
          await dialog.getByLabel('リンク先の種類', { exact: true }).inputValue(),
          'slide',
        );
        assert.equal(
          await dialog.getByLabel('移動先のスライド', { exact: true }).inputValue(),
          '1',
        );
        await dialog.getByRole('button', { name: 'キャンセル', exact: true }).click();
        await editor.locator('select').first().selectOption('en');
        ja = false;
        for (const kind of ['nextSlide', 'prevSlide', 'firstSlide', 'lastSlide']) {
          dialog = await open(true);
          await dialog.getByLabel('Link destination', { exact: true }).selectOption(kind);
          await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
          await saved();
          shape = await read();
          assert.deepEqual(getAction(shape, 0, 1), { kind });
          assert.equal(getShapeClickAction(shape), null);
          assert.equal(getAction(shape, 0, 0), null);
          assert.equal(getAction(shape, 0, 2), null);
          checkFormat(shape);
        }
        dialog = await open();
        assert.equal(
          await dialog.getByLabel('Link destination', { exact: true }).inputValue(),
          'lastSlide',
        );
        await dialog.getByRole('button', { name: 'Remove link', exact: true }).click();
        await saved();
        assert.equal(getAction(await read(), 0, 1), null);
        assert.deepEqual(errors, []);
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );

test(
  'selected table cells share bilingual links with atomic undo and reload',
  { timeout: 60000 },
  async () => {
    const api = await import('@office-kit/pptx');
    const dir = await mkdtemp(join(tmpdir(), 'office-cell-batch-links-'));
    let preview, browser, page;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Table} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Table x={1} y={1} width={8} height={2} rows={[["First 日本語","Second","Untouched"]]} /></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (e) => errors.push(e.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      let ja = false;
      const saved = () =>
        editor
          .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
          .waitFor();
      const read = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        const table = getSlideShapes(getSlides(pres)[0])[0];
        assert.equal(getShapeClickAction(table), null);
        return api
          .getTableCells(table)[0]
          .map((cell) => api.getTableCellParagraphs(cell)[0].elements[0]);
      };
      const cell = (c) =>
        editor.getByRole('button', { name: `${ja ? 'セル' : 'Cell'} 1, ${c}`, exact: true });
      const open = async (ribbon = false) => {
        if (ribbon) {
          await editor.getByRole('button', { name: ja ? '挿入' : 'Insert', exact: true }).click();
          await editor.locator('button[title$="— setShapeHyperlink"]').click();
        } else
          await editor
            .getByRole('button', { name: ja ? 'リンクを編集' : 'Edit link', exact: true })
            .click();
        const dialog = editor.getByRole('dialog');
        await dialog
          .getByText(
            ja
              ? '選択したセル内のすべての文字に適用します。'
              : 'Applies to all text in the selected cells.',
            { exact: true },
          )
          .waitFor();
        return dialog;
      };
      await saved();
      await editor.locator('.hit').first().click();
      // The table panel starts with its first cell even before explicit cell selection.
      let dialog = await open();
      await dialog.getByLabel('Link address', { exact: true }).fill('https://example.com/first');
      await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
      await saved();
      assert.deepEqual(
        (await read()).map((run) => run.clickAction?.url ?? null),
        ['https://example.com/first', null, null],
      );
      await cell(1).click();
      await cell(2).click({ modifiers: ['Shift'] });
      await editor.locator('select').first().selectOption('ja');
      ja = true;
      dialog = await open(true);
      await dialog
        .getByText('選択したテキストには異なるリンクが設定されています。', { exact: true })
        .waitFor();
      await dialog.getByLabel('リンク先', { exact: true }).fill('https://example.com/shared');
      await dialog.getByLabel('リンクの説明', { exact: true }).fill('共通の資料');
      await dialog.getByRole('button', { name: '適用', exact: true }).click();
      await saved();
      const shared = ['https://example.com/shared', 'https://example.com/shared', null];
      assert.deepEqual(
        (await read()).map((run) => run.clickAction?.url ?? null),
        shared,
      );
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(
        (await read()).map((run) => run.clickAction?.url ?? null),
        ['https://example.com/first', null, null],
      );
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      assert.deepEqual(
        (await read()).map((run) => run.tooltip ?? null),
        ['共通の資料', '共通の資料', null],
      );
      await editor.locator('.hit').first().click();
      await cell(1).click();
      await cell(2).click({ modifiers: ['Shift'] });
      dialog = await open();
      assert.equal(await dialog.getByLabel('リンク先', { exact: true }).inputValue(), shared[0]);
      await dialog.getByRole('button', { name: 'リンクを解除', exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      const runs = await read();
      assert.deepEqual(
        runs.map((run) => run.clickAction ?? null),
        [null, null, null],
      );
      assert.deepEqual(
        runs.map((run) => run.text),
        ['First 日本語', 'Second', 'Untouched'],
      );
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({
        path: '/tmp/pptx-pr287-cell-batch-link-failure.png',
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

test(
  'deleting a linked slide clears its links and undo restores destinations',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-delete-linked-slide-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={4} height={1}>Source</Text></Slide><Slide><Text x={1} y={1} width={4} height={1}>Target</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (e) => errors.push(e.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      const read = async () =>
        loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
      await editor.locator('select').first().selectOption('ja');
      await saved();
      await editor.locator('.hit').first().click();
      await editor.getByRole('button', { name: '挿入', exact: true }).click();
      await editor.locator('button[title$="— setShapeHyperlink"]').click();
      const dialog = editor.getByRole('dialog');
      await dialog.getByLabel('リンク先の種類', { exact: true }).selectOption('slide');
      await dialog.getByLabel('移動先のスライド', { exact: true }).selectOption('1');
      await dialog.getByRole('button', { name: '適用', exact: true }).click();
      await saved();
      await editor.locator('.thumb-row').nth(1).click();
      await editor.locator('.thumb-row').nth(1).press('Delete');
      await saved();
      let pres = await read();
      assert.equal(getSlides(pres).length, 1);
      assert.equal(getShapeClickAction(getSlideShapes(getSlides(pres)[0])[0]), null);
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      pres = await read();
      assert.equal(getSlides(pres).length, 2);
      const restored = getShapeClickAction(getSlideShapes(getSlides(pres)[0])[0]);
      assert.equal(restored?.kind, 'slide');
      assert.equal(getSlideIndex(pres, restored.slide), 1);
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      pres = await read();
      assert.equal(getSlides(pres).length, 1);
      assert.equal(getShapeClickAction(getSlideShapes(getSlides(pres)[0])[0]), null);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

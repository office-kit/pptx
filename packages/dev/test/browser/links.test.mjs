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
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={4} height={1} bold>Link text</Text></Slide><Slide><Text x={1} y={1} width={4} height={1}>Destination</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
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
      await dialog.getByLabel('リンク先の種類', { exact: true }).selectOption('slide');
      await dialog.getByLabel('移動先のスライド', { exact: true }).selectOption('1');
      await page.screenshot({ path: '/tmp/pptx-pr287-object-link-ja.png', fullPage: true });
      await dialog.getByRole('button', { name: '適用', exact: true }).click();
      await saved();
      pres = await read();
      assert.equal(getSlideIndex(pres, getShapeClickAction(picture(pres)).slide), 1);
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(getShapeClickAction(picture(await read())).url, 'https://example.com/image');
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

test(
  'selected text links preserve surrounding text with Japanese undo and reload',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-range-links-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={6} height={1} bold>Before 日本語 After</Text></Slide><Slide><Text x={1} y={1} width={5} height={1}>Target</Text></Slide></Presentation>`,
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
        await editor.locator('.hit').first().dblclick();
        const input = editor.locator('textarea.inline-edit');
        await input.evaluate((el) => {
          el.focus();
          el.setSelectionRange(7, 10);
          el.dispatchEvent(new Event('select', { bubbles: true }));
        });
        if (keyboard) await input.press('Control+k');
        else
          await editor
            .locator('.text-format-bar')
            .getByRole('button', { name: ja ? 'リンクを編集' : 'Edit link', exact: true })
            .click();
        return editor.getByRole('dialog', { name: ja ? 'リンクを編集' : 'Edit link', exact: true });
      };
      await saved();
      let dialog = await open();
      await dialog.getByText('Applies to the selected text.', { exact: true }).waitFor();
      await dialog.getByLabel('Link address', { exact: true }).fill('https://example.com/japanese');
      await dialog.getByLabel('Link description', { exact: true }).fill('日本語の資料');
      await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
      await saved();
      const { getShapeRunHyperlink, getShapeRunHyperlinkTooltip, getShapeRunClickAction } =
        await import('@office-kit/pptx');
      let shape = await read();
      assert.deepEqual(
        getShapeParagraphElements(shape, 0).map((e) => e.text),
        ['Before ', '日本語', ' After'],
      );
      assert.deepEqual(
        [0, 1, 2].map((r) => getShapeRunHyperlink(shape, 0, r)),
        [null, 'https://example.com/japanese', null],
      );
      assert.ok(getShapeParagraphElements(shape, 0).every((e) => e.format.bold));
      await editor.locator('select').first().selectOption('ja');
      ja = true;
      dialog = await open(true);
      assert.equal(
        await dialog.getByLabel('リンク先', { exact: true }).inputValue(),
        'https://example.com/japanese',
      );
      await dialog.getByLabel('リンク先', { exact: true }).fill('https://example.com/cancelled');
      await dialog.getByRole('button', { name: 'キャンセル', exact: true }).click();
      assert.equal(getShapeRunHyperlink(await read(), 0, 1), 'https://example.com/japanese');
      dialog = await open();
      await dialog.getByRole('button', { name: 'リンクを解除', exact: true }).click();
      await saved();
      assert.equal(getShapeRunHyperlink(await read(), 0, 1), null);
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      shape = await read();
      assert.equal(getShapeRunHyperlink(shape, 0, 1), 'https://example.com/japanese');
      assert.equal(getShapeRunHyperlinkTooltip(shape, 0, 1), '日本語の資料');
      dialog = await open(true);
      await dialog.getByLabel('リンク先の種類', { exact: true }).selectOption('slide');
      await dialog.getByLabel('移動先のスライド', { exact: true }).selectOption('1');
      await dialog.getByRole('button', { name: '適用', exact: true }).click();
      await saved();
      assert.equal(getShapeRunClickAction(await read(), 0, 1)?.kind, 'slide');
      await page.reload();
      await saved();
      dialog = await open();
      assert.equal(
        await dialog.getByLabel('リンク先の種類', { exact: true }).inputValue(),
        'slide',
      );
      assert.equal(await dialog.getByLabel('移動先のスライド', { exact: true }).inputValue(), '1');
      await dialog.getByRole('button', { name: 'キャンセル', exact: true }).click();
      await editor.locator('select').first().selectOption('en');
      ja = false;
      for (const kind of ['nextSlide', 'prevSlide', 'firstSlide', 'lastSlide']) {
        dialog = await open(true);
        await dialog.getByLabel('Link destination', { exact: true }).selectOption(kind);
        await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
        await saved();
        shape = await read();
        assert.deepEqual(getShapeRunClickAction(shape, 0, 1), { kind });
        assert.equal(getShapeClickAction(shape), null);
        assert.equal(getShapeRunClickAction(shape, 0, 0), null);
        assert.equal(getShapeRunClickAction(shape, 0, 2), null);
        assert.ok(getShapeParagraphElements(shape, 0).every((e) => e.format.bold));
      }
      dialog = await open();
      assert.equal(
        await dialog.getByLabel('Link destination', { exact: true }).inputValue(),
        'lastSlide',
      );
      await dialog.getByRole('button', { name: 'Remove link', exact: true }).click();
      await saved();
      assert.equal(getShapeRunClickAction(await read(), 0, 1), null);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

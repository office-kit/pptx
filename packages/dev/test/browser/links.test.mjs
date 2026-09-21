import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
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

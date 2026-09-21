import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlides,
  getSlideShapes,
  getShapeText,
  getShapeParagraphElements,
  getTableCells,
  getTableCellText,
  getTableCellParagraphs,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'bilingual find and replace navigates formatted text and table cells, saves and undoes',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-find-browser-'));
    let preview, browser, page;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text,Table} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={7} height={1} paragraphs={[{runs:[{text:'Hel',format:{bold:true}},{text:'lo 日本語 HELLO',format:{italic:true}}]}]} /></Slide><Slide><Table x={1} y={1} width={6} height={2} rows={[[{paragraphs:[{runs:[{text:'hel',format:{bold:true}},{text:'lo 日本語',format:{italic:true}}]}]},'Keep']]} /></Slide><Slide><Text x={1} y={1} width={7} height={1} >hello 日本語</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (e) => errors.push(e.message));
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
      const tableText = async () =>
        getTableCellText(getTableCells(getSlideShapes((await slides())[1])[0])[0][0]);
      await saved();
      await editor.getByRole('button', { name: 'Replace', exact: true }).click();
      let dialog = editor.getByRole('dialog', { name: 'Find and replace', exact: true });
      await dialog.getByLabel('Find text', { exact: true }).fill('Hello');
      await dialog.getByText('Match 1 / 4', { exact: true }).waitFor();
      await dialog.getByLabel('Match case', { exact: true }).check();
      await dialog.getByText('Match 1 / 1', { exact: true }).waitFor();
      await dialog.getByLabel('Match case', { exact: true }).uncheck();
      await dialog.getByLabel('Replace with', { exact: true }).fill('Hello!');
      await dialog.getByRole('button', { name: 'Replace match', exact: true }).click();
      await dialog.getByText('Match 2 / 4', { exact: true }).waitFor();
      await saved();
      assert.equal(getShapeText(getSlideShapes((await slides())[0])[0]), 'Hello! 日本語 HELLO');
      await dialog.getByRole('button', { name: 'Close', exact: true }).click();
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      await editor.getByRole('button', { name: 'Replace', exact: true }).click();
      await dialog.getByLabel('Find text', { exact: true }).fill('Hello');
      await dialog.getByRole('button', { name: 'Next match', exact: true }).click();
      await dialog.getByText('Match 2 / 4', { exact: true }).waitFor();
      await dialog.getByLabel('Find text', { exact: true }).press('Enter');
      await dialog.getByText('Slide 2 · Cell 1, 1', { exact: true }).waitFor();
      await dialog.getByLabel('Replace with', { exact: true }).fill('$&置換');
      await dialog.getByRole('button', { name: 'Replace match', exact: true }).click();
      await saved();
      assert.equal(await tableText(), '$&置換 日本語');
      const elements = getTableCellParagraphs(
        getTableCells(getSlideShapes((await slides())[1])[0])[0][0],
      )[0].elements;
      assert.equal(elements[0].text, '$&置換');
      assert.equal(elements[0].format.bold, true);
      assert.equal(elements[1].text, ' 日本語');
      assert.equal(elements[1].format.italic, true);
      await dialog.getByRole('button', { name: 'Close', exact: true }).click();
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(await tableText(), 'hello 日本語');
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.equal(await tableText(), '$&置換 日本語');
      await editor.locator('.thumb-row').nth(1).click();
      await editor.locator('.thumb-row').nth(1).press('Control+f');
      await dialog.getByLabel('Find text', { exact: true }).fill('日本語');
      await dialog.getByLabel('Only the starting slide', { exact: true }).check();
      await dialog.getByText('Match 1 / 1', { exact: true }).waitFor();
      await dialog.getByLabel('Replace with', { exact: true }).fill('Table');
      await dialog.getByRole('button', { name: 'Replace all', exact: true }).click();
      await saved();
      assert.equal(await tableText(), '$&置換 Table');
      assert.equal(getShapeText(getSlideShapes((await slides())[0])[0]), 'Hello 日本語 HELLO');
      await dialog.getByRole('button', { name: 'Close', exact: true }).click();
      await editor.locator('.lang select').selectOption('ja');
      locale = 'ja';
      await editor.locator('.thumb-row').first().click();
      await editor.locator('.thumb-row').first().press('Control+h');
      dialog = editor.getByRole('dialog', { name: '検索と置換', exact: true });
      await dialog.getByLabel('検索する文字列', { exact: true }).fill('日本語');
      await dialog.getByText('一致 1 / 2', { exact: true }).waitFor();
      await dialog.getByRole('button', { name: '前の一致箇所', exact: true }).click();
      await dialog.getByText('一致 2 / 2', { exact: true }).waitFor();
      await dialog.getByLabel('置換後の文字列', { exact: true }).fill('$1🗾');
      await page.screenshot({ path: '/tmp/pptx-pr287-find-ja.png', fullPage: true });
      await dialog.getByRole('button', { name: 'すべて置換', exact: true }).click();
      await dialog.getByText('一致する文字列はありません', { exact: true }).waitFor();
      await saved();
      await dialog.getByLabel('検索する文字列', { exact: true }).press('Escape');
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      const undone = await slides();
      assert.equal(getShapeText(getSlideShapes(undone[0])[0]), 'Hello 日本語 HELLO');
      assert.equal(getShapeText(getSlideShapes(undone[2])[0]), 'hello 日本語');
      assert.equal(await tableText(), '$&置換 Table');
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      const result = await slides();
      assert.equal(getShapeText(getSlideShapes(result[0])[0]), 'Hello $1🗾 HELLO');
      assert.equal(getShapeText(getSlideShapes(result[2])[0]), 'hello $1🗾');
      assert.equal(await tableText(), '$&置換 Table');
      assert.equal(getShapeParagraphElements(getSlideShapes(result[0])[0], 0)[0].format.bold, true);
      assert.equal(
        getShapeParagraphElements(getSlideShapes(result[0])[0], 0)[1].format.italic,
        true,
      );
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-pr287-find-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlides,
  getSlideComments,
  getCommentAuthor,
  getCommentText,
  getCommentDate,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'bilingual comments add, edit, cancel, delete, undo and reload',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-comments-'));
    let preview, browser, page;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide} from '@office-kit/pptx-dsl';export default <Presentation><Slide/><Slide/></Presentation>`,
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
        getSlides(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
        ).map(getSlideComments);
      const open = async () => {
        await editor.getByRole('button', { name: ja ? '挿入' : 'Insert', exact: true }).click();
        await editor.locator('button[title$="— addSlideComment"]').click();
        return editor.getByRole('dialog', { name: ja ? 'コメント' : 'Comments', exact: true });
      };
      await saved();
      let dialog = await open();
      assert.equal(
        await dialog.getByRole('button', { name: 'Apply', exact: true }).isDisabled(),
        true,
      );
      await dialog.getByLabel('Author name', { exact: true }).fill('Reviewer');
      await dialog.getByLabel('Comment text', { exact: true }).fill('First comment');
      await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
      await saved();
      let comments = await read();
      assert.equal(getCommentText(comments[0][0]), 'First comment');
      assert.equal(comments[1].length, 0);
      const date = getCommentDate(comments[0][0]);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      dialog = await open();
      await dialog.getByLabel('コメントの内容', { exact: true }).fill('キャンセル');
      await dialog.getByRole('button', { name: 'キャンセル', exact: true }).click();
      assert.equal(getCommentText((await read())[0][0]), 'First comment');
      dialog = await open();
      await dialog
        .getByLabel('コメントの内容', { exact: true })
        .fill('修正済み <script>test</script>');
      await dialog.getByRole('button', { name: 'コメントを追加', exact: true }).click();
      await dialog.getByLabel('作成者名', { exact: true }).fill('山田');
      await dialog
        .getByLabel('コメントの内容', { exact: true })
        .nth(1)
        .fill('日本語のコメント\nSecond line');
      await page.screenshot({ path: '/tmp/pptx-pr287-comments-ja.png', fullPage: true });
      await dialog.getByRole('button', { name: '適用', exact: true }).click();
      await saved();
      comments = await read();
      assert.equal(comments[0].length, 2);
      assert.equal(getCommentDate(comments[0][0]), date);
      assert.equal(getCommentAuthor(comments[0][0]).name, 'Reviewer');
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual((await read())[0].map(getCommentText), ['First comment']);
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      dialog = await open();
      assert.equal(
        await dialog.getByLabel('コメントの内容', { exact: true }).nth(0).inputValue(),
        '修正済み <script>test</script>',
      );
      await dialog.getByRole('button', { name: 'コメントを削除', exact: true }).nth(1).click();
      await dialog.getByRole('button', { name: 'コメントを削除', exact: true }).click();
      await dialog.getByText('このスライドにコメントはありません。', { exact: true }).waitFor();
      await dialog.getByRole('button', { name: '適用', exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      assert.deepEqual(
        (await read()).map((list) => list.length),
        [0, 0],
      );
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-pr287-comments-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

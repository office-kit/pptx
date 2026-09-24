import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getShapeText, getSlideShapes, getSlides, loadPresentation } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'text content editing requires one text shape while bulk formatting stays available',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-text-content-selection-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text,Line} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={2} height={1}>日本語</Text><Text x={4} y={1} width={2} height={1}>English</Text><Line x1={1} y1={3} x2={4} y2={4}/></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const panel = editor.locator('.bespoke');
      let ja = false;
      const saved = () =>
        editor
          .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
          .waitFor();
      const read = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getSlideShapes(getSlides(pres)[0]).slice(0, 2).map(getShapeText);
      };
      await saved();
      await editor.locator('.hit').nth(0).click();
      await editor.getByRole('tab', { name: 'Size & Properties', exact: true }).click();
      assert.equal(
        await panel.getByRole('textbox', { name: 'Text', exact: true }).inputValue(),
        '日本語',
      );
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      await panel
        .getByText('Select one text shape to edit its content.', { exact: true })
        .waitFor();
      assert.equal(await panel.locator('textarea').count(), 0);
      assert.equal(await panel.locator('.text-format-bar').count(), 1);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await panel
        .getByText('文章を編集するには、文字を入力できる図形を1つ選択してください。', {
          exact: true,
        })
        .waitFor();
      await editor
        .locator('.hit')
        .nth(0)
        .click({ modifiers: ['Shift'] });
      const content = panel.getByRole('textbox', { name: 'テキスト', exact: true });
      assert.equal(await content.inputValue(), 'English');
      await content.fill('English updated\n日本語も編集');
      await content.press('Tab');
      await saved();
      assert.deepEqual(await read(), ['日本語', 'English updated\n日本語も編集']);
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), ['日本語', 'English']);
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      await editor.locator('.hit').nth(2).click();
      assert.equal(await panel.locator('textarea').count(), 0);
      assert.equal(await panel.locator('.text-format-bar').count(), 0);
      await page.reload();
      await saved();
      assert.deepEqual(await read(), ['日本語', 'English updated\n日本語も編集']);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

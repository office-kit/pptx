import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlides,
  getSlideShapes,
  getParagraphBullet,
  getParagraphLevel,
  getParagraphAlignment,
  getParagraphLineSpacing,
  getParagraphSpacing,
  getShapeParagraphElements,
  getShapeText,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'paragraph controls preserve individual paragraphs and rich text across bilingual save, undo and reload',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-paragraph-browser-'));
    let preview, browser, page;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={7} height={4} paragraphs={[{runs:[{text:'English',format:{bold:true}}]},{runs:[{text:'日本語',format:{italic:true}}]},{runs:[{text:'Third paragraph'}]}]} /></Slide><Slide><Text x={1} y={1} width={7} height={3} /></Slide></Presentation>`,
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
      const shape = async () =>
        getSlideShapes(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        )[0];
      await saved();
      await editor.locator('.hit').first().click();
      let panel = editor.getByRole('region', { name: 'Paragraph formatting', exact: true });
      await panel.getByLabel('List style', { exact: true }).selectOption('bullet');
      await saved();
      for (let i = 0; i < 3; i++) assert.equal(getParagraphBullet(await shape(), i), 'bullet');
      await panel.getByLabel('Apply to paragraphs', { exact: true }).selectOption('1');
      await panel.getByLabel('List level', { exact: true }).selectOption({ value: '2' });
      await saved();
      assert.equal(getParagraphLevel(await shape(), 1), 2);
      assert.equal(getParagraphLevel(await shape(), 0), 0);
      await panel.getByLabel('List style', { exact: true }).selectOption('number');
      await saved();
      await panel.getByLabel('Paragraph alignment', { exact: true }).selectOption('center');
      await saved();
      await panel.getByLabel('Line spacing mode', { exact: true }).selectOption('pct');
      await saved();
      await panel.getByLabel('Line spacing value', { exact: true }).fill('1.5');
      await panel.getByLabel('Line spacing value', { exact: true }).press('Tab');
      await saved();
      assert.deepEqual(getParagraphLineSpacing(await shape(), 1), { kind: 'pct', value: 1.5 });
      assert.equal(getParagraphLineSpacing(await shape(), 0), null);
      await panel.getByLabel('Before paragraph (pt)', { exact: true }).fill('6');
      await panel.getByLabel('Before paragraph (pt)', { exact: true }).press('Tab');
      await saved();
      await panel.getByLabel('After paragraph (pt)', { exact: true }).fill('3');
      await panel.getByLabel('After paragraph (pt)', { exact: true }).press('Tab');
      await saved();
      assert.deepEqual(getParagraphSpacing(await shape(), 1), { beforePts: 6, afterPts: 3 });
      await panel.getByLabel('Apply to paragraphs', { exact: true }).selectOption('-1');
      assert.equal(await panel.getByLabel('List style', { exact: true }).inputValue(), '');
      assert.equal(await panel.getByLabel('List level', { exact: true }).inputValue(), '');
      await editor.locator('.lang select').selectOption('ja');
      locale = 'ja';
      panel = editor.getByRole('region', { name: '段落の書式', exact: true });
      await panel.getByLabel('リストの種類', { exact: true }).selectOption('none');
      await saved();
      for (let i = 0; i < 3; i++) assert.equal(getParagraphBullet(await shape(), i), 'none');
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(getParagraphBullet(await shape(), 1), 'number');
      assert.equal(getParagraphBullet(await shape(), 0), 'bullet');
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.equal(getParagraphBullet(await shape(), 1), 'none');
      await panel.getByLabel('対象の段落', { exact: true }).selectOption('1');
      await panel.getByLabel('行間の指定方法', { exact: true }).selectOption('pts');
      await saved();
      await panel.getByLabel('行間の値', { exact: true }).fill('24');
      await panel.getByLabel('行間の値', { exact: true }).press('Tab');
      await saved();
      await page.screenshot({ path: '/tmp/pptx-pr287-paragraph-ja.png', fullPage: true });
      await page.reload();
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      await saved();
      await editor.locator('.hit').first().click();
      await panel.getByLabel('対象の段落', { exact: true }).selectOption('1');
      assert.equal(await panel.getByLabel('行間の値', { exact: true }).inputValue(), '24');
      const result = await shape();
      assert.deepEqual(getParagraphLineSpacing(result, 1), { kind: 'pts', value: 24 });
      assert.equal(getParagraphAlignment(result, 1), 'ctr');
      assert.equal(getShapeParagraphElements(result, 0)[0].format.bold, true);
      assert.equal(getShapeParagraphElements(result, 1)[0].format.italic, true);
      await panel.getByLabel('行間の指定方法', { exact: true }).selectOption('inherit');
      await saved();
      assert.equal(getParagraphLineSpacing(await shape(), 1), null);
      await editor.locator('.thumb-row').nth(1).click();
      await editor.locator('.hit').first().click();
      await panel.getByLabel('リストの種類', { exact: true }).selectOption('number');
      await saved();
      await panel.getByLabel('段落の配置', { exact: true }).selectOption('right');
      await saved();
      await editor.locator('.hit').first().dblclick();
      await editor.locator('.inline-edit').fill('First line\n最初の入力');
      await editor.locator('.inline-edit').press('Control+Enter');
      await saved();
      const emptyShape = async () =>
        getSlideShapes(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[1],
        )[0];
      assert.equal(getShapeText(await emptyShape()), 'First line\n最初の入力');
      assert.equal(getParagraphBullet(await emptyShape(), 0), 'number');
      assert.equal(getParagraphAlignment(await emptyShape(), 0), 'r');
      assert.equal(getParagraphBullet(await emptyShape(), 1), 'number');
      assert.equal(getParagraphAlignment(await emptyShape(), 1), 'r');
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(getShapeText(await emptyShape()), '');
      await panel.getByLabel('リストの種類', { exact: true }).waitFor();
      assert.equal(await panel.getByLabel('段落の配置', { exact: true }).inputValue(), 'right');
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.equal(getShapeText(await emptyShape()), 'First line\n最初の入力');
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-paragraph-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

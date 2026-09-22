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
  getSlideLayout,
  getSlideLayoutName,
  isShapePlaceholder,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'new slides use the chosen layout with editable placeholders, history and bilingual persistence',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-new-layout-'));
    let preview, browser, page;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation>{['Before','After'].map(text => <Slide><Text x={1} y={1} width={7} height={1}>{text}</Text></Slide>)}</Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const thumbs = editor.locator('.thumb-row');
      let ja = false;
      const saved = () =>
        editor
          .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
          .waitFor();
      const slides = async () =>
        getSlides(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
        );
      await saved();
      await thumbs.nth(0).click();
      await editor
        .locator('.nav-actions')
        .getByRole('button', { name: 'New slide from layout', exact: true })
        .click();
      const dialog = editor.getByRole('dialog', { name: 'New slide from layout', exact: true });
      await dialog
        .getByLabel('Slide layout', { exact: true })
        .selectOption({ label: 'Title and Content' });
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.equal(await thumbs.count(), 2);
      // The ribbon reaches the same picker, independently of the slide navigator.
      await editor
        .locator('.ribbon')
        .getByRole('button', { name: 'New slide from layout', exact: true })
        .click();
      await dialog
        .getByLabel('Slide layout', { exact: true })
        .selectOption({ label: 'Title and Content' });
      await dialog.getByRole('button', { name: 'Insert slide', exact: true }).click();
      await saved();
      assert.equal(await thumbs.count(), 3);
      assert.match(await thumbs.nth(1).getAttribute('class'), /active/);
      let deck = await slides();
      assert.equal(getSlideLayoutName(getSlideLayout(deck[1])), 'Title and Content');
      assert.ok(getSlideShapes(deck[1]).filter(isShapePlaceholder).length >= 2);
      assert.equal(getShapeText(getSlideShapes(deck[0])[0]), 'Before');
      assert.equal(getShapeText(getSlideShapes(deck[2])[0]), 'After');
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await slides()).length, 2);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.equal((await slides()).length, 3);
      await editor
        .locator('.hit')
        .first()
        .dblclick({ position: { x: 30, y: 20 } });
      const input = editor.locator('.inline-edit');
      await input.fill('新しいスライド / New slide');
      await input.press('Control+Enter');
      await saved();
      assert.ok(
        getSlideShapes((await slides())[1]).some(
          (shape) => getShapeText(shape) === '新しいスライド / New slide',
        ),
      );
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await editor
        .locator('.nav-actions')
        .getByRole('button', { name: 'レイアウトからスライドを追加', exact: true })
        .click();
      const jp = editor.getByRole('dialog', { name: 'レイアウトからスライドを追加', exact: true });
      await jp
        .getByLabel('スライドのレイアウト', { exact: true })
        .selectOption({ label: 'タイトルスライド' });
      await page.screenshot({ path: '/tmp/pptx-new-slide-layout-ja.png', fullPage: true });
      await jp.getByRole('button', { name: 'スライドを挿入', exact: true }).click();
      await saved();
      assert.equal(getSlideLayoutName(getSlideLayout((await slides())[2])), 'Title Slide');
      await page.reload();
      await saved();
      deck = await slides();
      assert.equal(deck.length, 4);
      assert.ok(
        getSlideShapes(deck[1]).some(
          (shape) => getShapeText(shape) === '新しいスライド / New slide',
        ),
      );
      assert.equal(getSlideLayoutName(getSlideLayout(deck[2])), 'Title Slide');
      assert.equal(getShapeText(getSlideShapes(deck[3])[0]), 'After');
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-new-slide-layout-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

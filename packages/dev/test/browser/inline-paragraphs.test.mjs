import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getTableCells,
  getSlides,
  getSlideShapes,
  getParagraphBullet,
  getParagraphAlignment,
  getShapeParagraphElements,
  getShapeText,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'inline paragraph formatting follows caret and selection with pending text edits',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-inline-paragraph-'));
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
      await editor.locator('.hit').first().dblclick();
      const input = editor.locator('.inline-edit');
      const select = async (start, end = start) => {
        await input.focus();
        await input.evaluate(
          (node, range) => {
            node.setSelectionRange(range.start, range.end);
            node.dispatchEvent(new Event('select', { bubbles: true }));
          },
          { start, end },
        );
      };
      const bar = editor.locator('.text-format-bar');
      await select(9);
      await bar.getByLabel('Paragraph alignment', { exact: true }).selectOption('center');
      await saved();
      assert.equal(getParagraphAlignment(await shape(), 1), 'ctr');
      assert.notEqual(getParagraphAlignment(await shape(), 0), 'ctr');
      assert.notEqual(getParagraphAlignment(await shape(), 2), 'ctr');
      await select(0, 8);
      await bar.getByLabel('List style', { exact: true }).selectOption('bullet');
      await saved();
      assert.equal(getParagraphBullet(await shape(), 0), 'bullet');
      assert.notEqual(getParagraphBullet(await shape(), 1), 'bullet');
      await input.fill('English\n日本語\nThird paragraph\nNew paragraph');
      await select(30);
      await bar.getByLabel('Paragraph alignment', { exact: true }).selectOption('right');
      await saved();
      assert.equal(getParagraphAlignment(await shape(), 3), 'r');
      assert.notEqual(getParagraphAlignment(await shape(), 2), 'r');
      assert.equal(getShapeParagraphElements(await shape(), 0)[0].format.bold, true);
      assert.equal(getShapeParagraphElements(await shape(), 1)[0].format.italic, true);
      await bar.getByRole('button', { name: 'Done', exact: true }).click();
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(getShapeText(await shape()), 'English\n日本語\nThird paragraph');
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      await editor.locator('.lang select').selectOption('ja');
      locale = 'ja';
      await editor.locator('.hit').first().dblclick();
      await select(8, 28);
      await bar.getByLabel('段落の配置', { exact: true }).selectOption('justify');
      await saved();
      assert.equal(getParagraphAlignment(await shape(), 1), 'just');
      assert.equal(getParagraphAlignment(await shape(), 2), 'just');
      assert.equal(getParagraphAlignment(await shape(), 3), 'r');
      await page.reload();
      await saved();
      assert.equal(getParagraphAlignment(await shape(), 1), 'just');
      assert.equal(getParagraphBullet(await shape(), 0), 'bullet');
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'inline table paragraph formatting applies only within the edited cell',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-inline-paragraph-'));
    let preview, browser, page;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Table} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Table x={1} y={1} width={6} height={3} rows={[["First","Other"]]} /></Slide></Presentation>`,
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
      const hit = editor.locator('.hit').first();
      const bounds = await hit.boundingBox();
      await hit.dblclick({ position: { x: bounds.width / 4, y: bounds.height / 2 } });
      const input = editor.locator('.inline-edit');
      await input.fill('First\nSecond');
      await input.evaluate((node) => {
        node.setSelectionRange(8, 8);
        node.dispatchEvent(new Event('select', { bubbles: true }));
      });
      const bar = editor.locator('.text-format-bar');
      await bar.getByLabel('Paragraph alignment', { exact: true }).selectOption('right');
      await bar.getByLabel('List style', { exact: true }).selectOption('number');
      await saved();
      let cells = getTableCells(await shape());
      assert.equal(getParagraphAlignment(cells[0][0], 1), 'r');
      assert.equal(getParagraphBullet(cells[0][0], 1), 'number');
      assert.notEqual(getParagraphAlignment(cells[0][0], 0), 'r');
      assert.notEqual(getParagraphAlignment(cells[0][1], 0), 'r');
      await bar.getByRole('button', { name: 'Done', exact: true }).click();
      await page.reload();
      await saved();
      cells = getTableCells(await shape());
      assert.equal(getParagraphAlignment(cells[0][0], 1), 'r');
      assert.equal(getParagraphBullet(cells[0][0], 1), 'number');
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

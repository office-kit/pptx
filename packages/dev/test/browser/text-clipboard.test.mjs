import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlides,
  getSlideShapes,
  getShapeParagraphElements,
  getTableCells,
  getTableCellParagraphs,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'formatted clipboard copies pending text, cuts, pastes into cells and supports undo',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-text-clipboard-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text,Table} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={6} height={1}>English日本語</Text><Table x={1} y={3} width={6} height={2} rows={[["old","Other"]]} /></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (e) => errors.push(e.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const input = editor.locator('.inline-edit');
      const select = async (start, end = start) => {
        await input.focus();
        await input.evaluate(
          (node, range) => {
            node.setSelectionRange(...range);
            node.dispatchEvent(new Event('select', { bubbles: true }));
          },
          [start, end],
        );
      };
      const snapshot = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        const shapes = getSlideShapes(getSlides(pres)[0]);
        return {
          shape: getShapeParagraphElements(shapes[0], 0),
          cell: getTableCellParagraphs(getTableCells(shapes[1])[0][0])[0].elements,
        };
      };
      await saved();
      await editor
        .locator('.hit')
        .first()
        .dblclick({ position: { x: 30, y: 20 } });
      await select(0, 7);
      await input.press('Control+b');
      await select(10);
      await page.keyboard.insertText('です');
      await select(0, 12);
      await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
      await input.press('ControlOrMeta+c');
      const copied = await input.evaluate((node) => {
        const data = new DataTransfer();
        node.dispatchEvent(
          new ClipboardEvent('copy', { clipboardData: data, bubbles: true, cancelable: true }),
        );
        return Object.fromEntries([...data.types].map((type) => [type, data.getData(type)]));
      });
      assert.equal(copied['text/plain'], 'English日本語です');
      await select(10, 12);
      await input.evaluate((node) =>
        node.dispatchEvent(
          new ClipboardEvent('cut', {
            clipboardData: new DataTransfer(),
            bubbles: true,
            cancelable: true,
          }),
        ),
      );
      assert.equal(await input.inputValue(), 'English日本語');
      await input.press('Control+Enter');
      await saved();
      await editor
        .locator('.hit')
        .nth(1)
        .dblclick({ position: { x: 30, y: 20 } });
      await select(0, 3);
      const paste = () =>
        input.evaluate((node, contents) => {
          const data = new DataTransfer();
          for (const [type, value] of Object.entries(contents)) data.setData(type, value);
          node.dispatchEvent(
            new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }),
          );
        }, copied);
      await input.press('ControlOrMeta+v');
      assert.equal(await input.inputValue(), 'English日本語です');
      await input.press('Control+Enter');
      await saved();
      let result = await snapshot();
      assert.equal(result.cell.map((r) => r.text).join(''), 'English日本語です');
      assert.equal(result.cell[0].format.bold, true);
      assert.notEqual(result.cell.at(-1).format?.bold, true);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await snapshot()).cell.map((r) => r.text).join(''), 'old');
      await editor
        .locator('.hit')
        .nth(1)
        .dblclick({ position: { x: 30, y: 20 } });
      await select(0, 3);
      await input.press('ControlOrMeta+Shift+v');
      await input.evaluate(async (node) => {
        for (let i = 0; i < 50 && node.value === 'old'; i++)
          await new Promise((resolve) => setTimeout(resolve, 20));
      });
      assert.equal(await input.inputValue(), 'English日本語です');
      await input.press('Control+Enter');
      await saved();
      assert.ok((await snapshot()).cell.every((r) => r.format?.bold !== true));
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      await editor.locator('.lang select').selectOption('ja');
      await editor
        .locator('.hit')
        .nth(1)
        .dblclick({ position: { x: 30, y: 20 } });
      await select(0, 3);
      await paste();
      await input.press('Escape');
      assert.equal((await snapshot()).cell.map((r) => r.text).join(''), 'old');
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlides,
  getSlideShapes,
  getTableCells,
  getTableCellText,
  getShapeBounds,
  isTableShape,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'cell clipboard preserves quoted text, supports overlapping ranges and undoable cut',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-table-clipboard-'));
    let preview, browser, page;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Table} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Table x={1} y={1} width={6} height={3} rows={[["A","B","C"],["D","E","F"],["G","H","I"]]} /></Slide></Presentation>`,
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
      const table = async () =>
        getSlideShapes(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        ).find(isTableShape);
      const values = async () =>
        getTableCells(await table()).map((row) => row.map(getTableCellText));
      const cell = (r, c) =>
        editor.getByRole('button', { name: `${ja ? 'セル' : 'Cell'} ${r}, ${c}`, exact: true });
      await saved();
      const bounds = getShapeBounds(await table());
      await editor.locator('.hit').first().click();
      const clipboard = async (type, text) =>
        cell(1, 1).evaluate(
          (node, { type, text }) => {
            const data = new DataTransfer();
            if (text !== undefined) data.setData('text/plain', text);
            const event = new ClipboardEvent(type, {
              clipboardData: data,
              bubbles: true,
              cancelable: true,
            });
            node.dispatchEvent(event);
            return { prevented: event.defaultPrevented, text: data.getData('text/plain') };
          },
          { type, text },
        );
      await cell(1, 1).click();
      const source = '"日本語\nEnglish"\t"a\tb"\n"say ""Hi"""\t';
      assert.equal((await clipboard('paste', source)).prevented, true);
      await saved();
      const expected = [
        ['日本語\nEnglish', 'a\tb', 'C'],
        ['say "Hi"', '', 'F'],
        ['G', 'H', 'I'],
      ];
      assert.deepEqual(await values(), expected);
      const copied = await clipboard('copy');
      assert.equal(copied.prevented, true);
      assert.equal(copied.text, source);
      await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
      await cell(1, 1).press('ControlOrMeta+c');
      assert.equal(await page.evaluate(() => navigator.clipboard.readText()), source);
      await cell(2, 2).click();
      await cell(2, 2).press('ControlOrMeta+v');
      await saved();
      assert.deepEqual(await values(), [
        ['日本語\nEnglish', 'a\tb', 'C'],
        ['say "Hi"', '日本語\nEnglish', 'a\tb'],
        ['G', 'say "Hi"', ''],
      ]);
      assert.deepEqual(getShapeBounds(await table()), bounds);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await values(), expected);
      await cell(1, 1).click();
      await cell(2, 2).click({ modifiers: ['Shift'] });
      assert.equal((await clipboard('cut')).text, copied.text);
      await saved();
      assert.deepEqual(await values(), [
        ['', '', 'C'],
        ['', '', 'F'],
        ['G', 'H', 'I'],
      ]);
      await cell(2, 2).click();
      await clipboard('paste', copied.text);
      await saved();
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await values(), expected);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await cell(3, 1).click();
      await clipboard('paste', '末尾\tEnglish');
      await saved();
      assert.deepEqual((await values())[2], ['末尾', 'English', 'I']);
      await clipboard('paste', '"broken');
      await editor
        .getByText('クリップボードの表データを読み取れませんでした', { exact: true })
        .waitFor();
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await values(), expected);
      await page.reload();
      await saved();
      assert.deepEqual(await values(), expected);
      assert.deepEqual(getShapeBounds(await table()), bounds);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

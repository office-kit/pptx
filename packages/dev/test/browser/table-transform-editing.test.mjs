import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  createPresentation,
  addBlankSlide,
  addSlideTable,
  addSlideTextBox,
  groupShapes,
  setShapeBounds,
  setShapeRotation,
  setShapeFlip,
  savePresentation,
  loadPresentation,
  getSlides,
  getSlideShapes,
  isTableShape,
  getTableCells,
  getTableCellText,
  inches,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

for (const grouped of [false, true]) {
  for (const flip of [
    { horizontal: false, vertical: false },
    { horizontal: true, vertical: false },
    { horizontal: false, vertical: true },
    { horizontal: true, vertical: true },
  ]) {
    test(
      `table cell selection and input match rendered transforms (grouped=${grouped}, H=${flip.horizontal}, V=${flip.vertical})`,
      { timeout: 60000 },
      async () => {
        const dir = await mkdtemp(join(tmpdir(), 'office-table-transform-'));
        let preview, browser, page;
        try {
          const pres = createPresentation(),
            slide = addBlankSlide(pres);
          const table = addSlideTable(slide, {
            x: inches(2),
            y: inches(2),
            w: inches(4),
            h: inches(2),
            rows: [
              ['日本語 English', 'B'],
              ['C', 'D'],
            ],
          });
          setShapeRotation(table, 20);
          setShapeFlip(table, flip);
          if (grouped) {
            const sibling = addSlideTextBox(slide, {
              x: inches(7),
              y: inches(2),
              w: inches(1),
              h: inches(1),
              text: 'Sibling',
            });
            const group = groupShapes([table, sibling]);
            setShapeRotation(group, -25);
            setShapeFlip(group, { horizontal: true, vertical: false });
            setShapeBounds(group, { x: inches(1), y: inches(1), w: inches(8), h: inches(3) });
          }
          const source = join(dir, 'source.pptx'),
            file = join(dir, 'deck.tsx');
          await writeFile(source, await savePresentation(pres));
          await writeFile(
            file,
            `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
          );
          preview = await startPreview(file);
          browser = await chromium.launch({ headless: true });
          page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
          page.setDefaultTimeout(10000);
          const errors = [];
          page.on('pageerror', (e) => errors.push(e.message));
          await page.goto(preview.url);
          await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
          const editor = page.frameLocator('#editor-frame');
          await editor.getByText('Saved to this project', { exact: true }).waitFor();
          if (grouped) await editor.locator('.hit').dblclick();
          // Click the center of the rendered cell's symmetric text area, independent
          // of the editor's own cell coordinate calculations.
          const foreign = editor
            .locator('.paint foreignObject')
            .filter({ hasText: '日本語 English' });
          const rect = await foreign.boundingBox();
          const point = { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
          await page.mouse.dblclick(point.x, point.y);
          const input = editor.locator('.inline-edit');
          await input.waitFor();
          assert.equal(await input.innerText(), '日本語 English');
          // Starting a cell edit adds the text format bar, which shrinks the
          // canvas and re-fits the slide a frame later. Read the cell and both
          // overlays in one layout pass so a comparison can never straddle it.
          const centers = await input.evaluate((node) => {
            const center = (element) => {
              const rect = element.getBoundingClientRect();
              return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
            };
            return {
              cell: center(
                [...document.querySelectorAll('.paint foreignObject')].find((n) =>
                  n.textContent.includes('日本語 English'),
                ),
              ),
              input: center(node),
              selection: center(document.querySelector('.cell-selection')),
            };
          });
          assert.ok(Math.abs(centers.input.x - centers.cell.x) < 1, JSON.stringify(centers));
          assert.ok(Math.abs(centers.input.y - centers.cell.y) < 1, JSON.stringify(centers));
          assert.ok(Math.abs(centers.selection.x - centers.cell.x) < 1, JSON.stringify(centers));
          assert.ok(Math.abs(centers.selection.y - centers.cell.y) < 1, JSON.stringify(centers));
          const matrices = await input.evaluate((node) => {
            const text = [...document.querySelectorAll('.paint foreignObject')].find((n) =>
              n.textContent.includes('日本語 English'),
            );
            const expected = text.getScreenCTM();
            const actual = new DOMMatrix(getComputedStyle(node.parentElement).transform)
              .multiply(new DOMMatrix(getComputedStyle(node).transform))
              .scale(Number(getComputedStyle(node).getPropertyValue('--text-zoom')));
            return {
              actual: [actual.a, actual.b, actual.c, actual.d],
              expected: [expected.a, expected.b, expected.c, expected.d],
            };
          });
          matrices.actual.forEach((value, i) =>
            assert.ok(Math.abs(value - matrices.expected[i]) < 0.01, JSON.stringify(matrices)),
          );
          await input.fill('編集済み Edited');
          await page.keyboard.press('ControlOrMeta+Enter');
          await editor.getByText('Saved to this project', { exact: true }).waitFor();
          const loaded = await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          );
          const savedTable = getSlideShapes(getSlides(loaded)[0]).find(isTableShape);
          assert.deepEqual(
            getTableCells(savedTable).map((row) => row.map(getTableCellText)),
            [
              ['編集済み Edited', 'B'],
              ['C', 'D'],
            ],
          );
          await editor.locator('.lang select').selectOption('ja');
          await page.reload();
          await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
          if (grouped) await editor.locator('.hit').dblclick();
          const savedRect = await editor
            .locator('.paint foreignObject')
            .filter({ hasText: '編集済み Edited' })
            .boundingBox();
          await page.mouse.dblclick(
            savedRect.x + savedRect.width / 2,
            savedRect.y + savedRect.height / 2,
          );
          assert.equal(await editor.locator('.inline-edit').innerText(), '編集済み Edited');
          assert.equal(
            await editor.locator('.inline-edit').getAttribute('aria-label'),
            'セルのテキスト',
          );
          await page.screenshot({ path: '/tmp/pptx-table-transform-ja.png', fullPage: true });
          assert.deepEqual(errors, []);
        } catch (error) {
          await page?.screenshot({ path: '/tmp/pptx-table-transform-failure.png', fullPage: true });
          throw error;
        } finally {
          await browser?.close();
          await preview?.close();
          await rm(dir, { recursive: true, force: true });
        }
      },
    );
  }
}

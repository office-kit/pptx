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
  setShapeTextDirection,
  setShapeTextMargins,
  setShapeTextAnchor,
  setTableCellAnchor,
  getShapeTextDirection,
  getShapeText,
  getTableCells,
  setTableCellTextDirection,
  setTableCellMargins,
  savePresentation,
  loadPresentation,
  getSlides,
  getSlideShapes,
  getTableCellTextDirection,
  getTableCellText,
  inches,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

for (const anchor of ['top', 'center', 'bottom'])
  for (const tableCell of [true, false])
    for (const direction of [
      'vert',
      'vert270',
      'eaVert',
      'mongolianVert',
      'wordArtVert',
      'wordArtVertRtl',
    ]) {
      test(
        `${tableCell ? 'table' : 'shape'} ${direction} ${anchor} text retains its direction and glyph positions when edited`,
        { timeout: 60000 },
        async () => {
          const dir = await mkdtemp(join(tmpdir(), 'office-table-direction-'));
          let browser, preview;
          try {
            const pres = createPresentation();
            const bounds = {
              x: inches(2),
              y: inches(1),
              w: inches(3),
              h: inches(3),
            };
            const slide = addBlankSlide(pres);
            const shape = tableCell
              ? addSlideTable(slide, { ...bounds, rows: [['日本語 Text']] })
              : addSlideTextBox(slide, { ...bounds, text: '日本語 Text' });
            const margins = {
              left: inches(0.3),
              right: inches(0.1),
              top: inches(0.2),
              bottom: inches(0.4),
            };
            if (tableCell) {
              const cell = getTableCells(shape)[0][0];
              setTableCellTextDirection(cell, direction);
              setTableCellMargins(cell, margins);
              setTableCellAnchor(cell, anchor);
            } else {
              setShapeTextDirection(shape, direction);
              setShapeTextMargins(shape, margins);
              setShapeTextAnchor(shape, anchor);
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
            const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
            await page.goto(preview.url);
            await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
            const editor = page.frameLocator('#editor-frame');
            await editor.getByText('Saved to this project', { exact: true }).waitFor();
            const painted = editor
              .locator('.paint foreignObject > div')
              .filter({ hasText: '日本語 Text' });
            const writingMode =
              direction === 'vert270' || direction === 'mongolianVert'
                ? 'vertical-lr'
                : 'vertical-rl';
            assert.equal(
              await painted.evaluate((node) => getComputedStyle(node).writingMode),
              writingMode,
            );
            const glyphs = (node) => {
              const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT);
              let text;
              while ((text = walker.nextNode())) {
                if (!text.textContent.includes('日本語 Text')) continue;
                return Array.from({ length: text.textContent.length }, (_, index) => {
                  const range = document.createRange();
                  range.setStart(text, index);
                  range.setEnd(text, index + 1);
                  const rect = range.getBoundingClientRect();
                  return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
                });
              }
              throw new Error('Missing text');
            };
            const before = await painted.evaluate(glyphs);
            await editor.locator('.hit').dblclick();
            const input = editor.locator('.inline-edit');
            await input.waitFor();
            assert.equal(
              await input.evaluate((node) => getComputedStyle(node).writingMode),
              writingMode,
            );
            const after = await input.evaluate(glyphs);
            for (let i = 0; i < before.length; i++)
              for (const key of ['x', 'y', 'width', 'height']) {
                assert.ok(
                  Math.abs(after[i][key] - before[i][key]) < 2,
                  `${direction} glyph ${i} ${key}: ${before[i][key]} -> ${after[i][key]}`,
                );
              }
            await input.fill('編集済み');
            await input.press('ControlOrMeta+Enter');
            await editor.getByText('Saved to this project', { exact: true }).waitFor();
            const saved = await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            );
            const savedShape = getSlideShapes(getSlides(saved)[0])[0];
            if (tableCell) {
              const savedCell = getTableCells(savedShape)[0][0];
              assert.equal(getTableCellText(savedCell), '編集済み');
              assert.equal(getTableCellTextDirection(savedCell), direction);
            } else {
              assert.equal(getShapeText(savedShape), '編集済み');
              assert.equal(getShapeTextDirection(savedShape), direction);
            }
          } finally {
            await browser?.close();
            await preview?.close();
            await rm(dir, { recursive: true, force: true });
          }
        },
      );
    }

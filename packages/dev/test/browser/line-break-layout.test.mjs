import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  createPresentation,
  addBlankSlide,
  addSlideTable,
  addSlideTextBox,
  setShapeParagraphs,
  getTableCells,
  setTableCellParagraphs,
  inches,
  savePresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';
import { installRichTextSelection } from '../helpers/rich-text.mjs';

for (const leading of [false, true])
  for (const kind of ['table', 'shape'])
    test(
      `oversized ${leading ? 'leading ' : ''}${kind} line break does not shift surrounding text when editing`,
      { timeout: 60000 },
      async () => {
        const dir = await mkdtemp(join(tmpdir(), 'office-line-break-layout-'));
        let preview, browser;
        try {
          const pres = createPresentation();
          const slide = addBlankSlide(pres);
          const bounds = { x: inches(1), y: inches(1), w: inches(6), h: inches(3) };
          const shape =
            kind === 'table'
              ? addSlideTable(slide, { ...bounds, rows: [['']] })
              : addSlideTextBox(slide, { ...bounds, text: '' });
          const paragraphs = [
            {
              runs: [
                { text: 'Before', format: { size: 28 } },
                { text: 'After', format: { size: 28 } },
              ],
            },
          ];
          if (kind === 'table') setTableCellParagraphs(getTableCells(shape)[0][0], paragraphs);
          else setShapeParagraphs(shape, paragraphs);
          const parts = unzipSync(await savePresentation(pres));
          const name = 'ppt/slides/slide1.xml';
          parts[name] = strToU8(
            strFromU8(parts[name]).replace(
              leading ? '<a:r>' : '</a:r>',
              leading
                ? '<a:br><a:rPr sz="8000"/></a:br><a:r>'
                : '</a:r><a:br><a:rPr sz="8000"/></a:br>',
            ),
          );
          const source = join(dir, 'source.pptx');
          await writeFile(source, zipSync(parts));
          const file = join(dir, 'deck.tsx');
          await writeFile(
            file,
            `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
          );
          preview = await startPreview(file);
          browser = await chromium.launch({ headless: true });
          const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
          await installRichTextSelection(page);
          await page.goto(preview.url);
          await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
          const editor = page.frameLocator('#editor-frame');
          await editor.getByText('Saved to this project', { exact: true }).waitFor();
          const glyphBounds = async (locator) =>
            locator.evaluate((node) => {
              const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT);
              let text;
              while ((text = walker.nextNode())) {
                if (text.textContent === 'After') {
                  const range = document.createRange();
                  range.selectNodeContents(text);
                  const rect = range.getBoundingClientRect();
                  return { x: rect.x, y: rect.y, height: rect.height };
                }
              }
              throw new Error('After text not found');
            });
          const before = await glyphBounds(
            editor.locator('.paint foreignObject').filter({ hasText: 'After' }).first(),
          );
          await editor.locator('.hit').first().dblclick();
          const input = editor.locator('.inline-edit');
          await input.waitFor();
          const after = await glyphBounds(input);
          assert.ok(
            Math.abs(after.y - before.y) < 2,
            `editing shifted text: ${JSON.stringify({ before, after })}`,
          );
          if (leading) return;
          const copied = await input.evaluate((node) => {
            node.focus();
            window.selectEditorText(node, 6, 7);
            node.dispatchEvent(new Event('select', { bubbles: true }));
            const data = new DataTransfer();
            node.dispatchEvent(
              new ClipboardEvent('copy', { clipboardData: data, bubbles: true, cancelable: true }),
            );
            return JSON.parse(data.getData('application/x-office-kit-text+json'));
          });
          assert.equal(copied.text, '\n');
          assert.equal(copied.formats[0].format.size, 80);
          await page.keyboard.insertText('X');
          await input.getByText('X', { exact: true }).waitFor();
          const inserted = await input
            .locator('span')
            .filter({ hasText: /^X$/ })
            .last()
            .evaluate((node) => ({
              size: node.style.fontSize,
              lineHeight: getComputedStyle(node).lineHeight,
              height: node.getBoundingClientRect().height,
            }));
          assert.equal(inserted.size, 'calc(80pt * var(--text-zoom))');
          assert.ok(
            parseFloat(inserted.lineHeight) > 0,
            'inserted text must regain its line height',
          );
          assert.ok(inserted.height > 50, 'inserted large text must be visible');
          assert.equal(await input.textContent(), 'BeforeXAfter');
        } finally {
          await browser?.close();
          await preview?.close();
          await rm(dir, { recursive: true, force: true });
        }
      },
    );

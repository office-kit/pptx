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
  getTableCells,
  setTableCellParagraphs,
  inches,
  savePresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'empty table paragraph keeps its font size when editing and inspecting the caret',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-table-empty-paragraph-'));
    let preview, browser;
    try {
      const pres = createPresentation();
      const table = addSlideTable(addBlankSlide(pres), {
        x: inches(1),
        y: inches(1),
        w: inches(6),
        h: inches(3),
        rows: [['']],
      });
      setTableCellParagraphs(getTableCells(table)[0][0], [
        { runs: [], endFormat: { size: 36, font: 'Courier New', bold: true, color: '#AA2244' } },
        { runs: [{ text: 'After', format: { size: 18 } }] },
      ]);
      const source = join(dir, 'source.pptx');
      await writeFile(source, await savePresentation(pres));
      const file = join(dir, 'deck.tsx');
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
      const first = input.locator('[data-text-paragraph]').first();
      const style = await first.evaluate((node) => ({
        size: node.style.fontSize,
        family: node.style.fontFamily,
        weight: getComputedStyle(node).fontWeight,
        color: getComputedStyle(node).color,
      }));
      assert.equal(style.size, 'calc(36pt * var(--text-zoom))');
      assert.match(style.family, /Courier New/);
      assert.equal(style.weight, '700');
      assert.equal(style.color, 'rgb(170, 34, 68)');
      await first.evaluate((node) => {
        node.closest('[contenteditable]').focus();
        const range = document.createRange();
        range.setStart(node, 0);
        range.collapse(true);
        getSelection().removeAllRanges();
        getSelection().addRange(range);
        document.dispatchEvent(new Event('selectionchange'));
      });
      await page.keyboard.press('Control+T');
      const dialog = editor.getByRole('dialog', { name: 'Font', exact: true });
      await dialog.waitFor();
      assert.equal(await dialog.getByLabel('Font size', { exact: true }).inputValue(), '36');
      assert.equal(
        await dialog.getByLabel('Latin text font', { exact: true }).inputValue(),
        'Courier New',
      );
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

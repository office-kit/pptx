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
  getTableCellParagraphs,
  loadPresentation,
  isTableShape,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'font dialog preserves mixed table cell formatting when changing spacing',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-font-dialog-table-'));
    let preview;
    let browser;
    try {
      await writeFile(
        join(dir, 'deck.tsx'),
        `import {Presentation,Slide,Table} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Table x={1} y={1} width={8} height={3} rows={[[{paragraphs:[{runs:[{text:'Bold',format:{bold:true}}]}]},{paragraphs:[{runs:[{text:'Italic',format:{bold:false,italic:true}}]}]}],['Outside','Other']]} styleCell={({row,column}) => row === 1 && column === 0 ? {format:{color:'#FF0000'}} : undefined} /></Slide></Presentation>`,
      );
      preview = await startPreview(join(dir, 'deck.tsx'));
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const cell = (row, column) =>
        editor.getByRole('button', { name: `Cell ${row}, ${column}`, exact: true });
      const readCells = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(`${preview.url}/deck.pptx`)).arrayBuffer()),
        );
        const table = getSlideShapes(getSlides(pres)[0]).find(isTableShape);
        return getTableCells(table).map((row) =>
          row.map((item) =>
            getTableCellParagraphs(item)[0].elements.map((element) => ({
              text: element.text,
              format: element.format,
            })),
          ),
        );
      };
      await saved();
      const before = await readCells();
      assert.equal(before[0][0][0].format.bold, true);
      assert.equal(before[0][1][0].format.italic, true);
      assert.equal(before[1][0][0].format.color, '#FF0000');

      // A right-clicked table targets a cell, whose menu has no Format item, so open
      // the pane on the slide; it then follows the selection to the table.
      await editor.locator('.stage').click({ button: 'right', position: { x: 8, y: 8 } });
      await editor.getByRole('menuitem', { name: 'Format Background...', exact: true }).click();
      const tableHit = await editor.locator('.hit').first().boundingBox();
      assert.ok(tableHit, 'the table hit target is visible');
      // The opened pane narrows the canvas, so let the locator find the moved table.
      await editor
        .locator('.hit')
        .first()
        .click({ position: { x: 2, y: 2 } });
      await cell(1, 1).click();
      await cell(1, 2).click({ modifiers: ['Shift'] });
      assert.equal(await editor.locator('.cell-grid button[aria-pressed="true"]').count(), 2);
      const homeFontDialog = editor.getByRole('button', { name: 'Font dialog', exact: true });
      assert.equal(await homeFontDialog.isEnabled(), true);
      const home = editor.getByRole('tabpanel', { name: 'Home', exact: true });
      const bold = home.getByRole('button', { name: 'Bold', exact: true });
      assert.equal(await bold.isEnabled(), true);
      await bold.click();
      await saved();
      const boldCells = await readCells();
      assert.equal(boldCells[0][0][0].format.bold, true);
      assert.equal(boldCells[0][1][0].format.bold, true);
      assert.deepEqual(boldCells[1], before[1]);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await readCells(), before);
      await homeFontDialog.click();
      const dialog = editor.getByRole('dialog', { name: 'Font', exact: true });
      await dialog.waitFor();
      await dialog.getByRole('tab', { name: 'Character Spacing', exact: true }).click();
      await dialog.getByLabel('Spacing', { exact: true }).selectOption('expanded');
      await dialog.getByLabel('By (pt)', { exact: true }).fill('1.5');
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await saved();

      const after = await readCells();
      assert.equal(after[0][0][0].format.bold, true);
      assert.equal(after[0][0][0].format.spc, 150);
      assert.equal(after[0][1][0].format.italic, true);
      assert.equal(after[0][1][0].format.spc, 150);
      assert.equal(after[1][0][0].format.color, '#FF0000');
      assert.equal(after[1][0][0].format.spc, undefined);
      assert.equal(after[1][1][0].format.spc, undefined);

      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await readCells(), before);
      await page.reload();
      await saved();
      assert.deepEqual(await readCells(), before);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

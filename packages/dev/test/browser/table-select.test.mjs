import * as pptx from '@office-kit/pptx';
import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test(
  'Table context selection commits text and targets row, column and table operations',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-table-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={5} height={1}>Alpha [a] ALPHA</Text></Slide><Slide><Text x={1} y={1} width={5} height={1}>Beta alpha</Text></Slide></Presentation>`,
    );
    const proc = spawn(process.execPath, [
      fileURLToPath(new URL('../../dist/cli.mjs', import.meta.url)),
      'dev',
      file,
      '--port',
      '0',
    ]);
    let browser;
    try {
      const url = await new Promise((resolve, reject) => {
        let output = '';
        proc.stdout.on('data', (data) => {
          output += data;
          const match = output.match(/Preview: (http:\/\/\S+)/);
          if (match) resolve(match[1]);
        });
        proc.stderr.on('data', (data) => process.stderr.write(data));
        proc.on('error', reject);
        proc.on('exit', (code) => reject(new Error('Server exited: ' + code)));
      });
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
      await page.goto(url);
      await page.locator('.shape-hit').first().waitFor();
      await page.locator('.shape-hit').first().dblclick();
      const shapeText = page.getByRole('textbox', { name: 'Edit text', exact: true });
      await shapeText.evaluate((field) => field.setSelectionRange(0, 5));
      for (const [key, level] of [
        [']', 1],
        ['[', 0],
      ]) {
        await shapeText.dispatchEvent('keydown', { key, metaKey: true });
        await page.waitForFunction((level) => {
          const state = window.office.getState();
          return (
            !state.building &&
            state.editor.slides[0].shapes[0].paragraphs[0].properties.level === level
          );
        }, level);
        await shapeText.waitFor();
        assert.deepEqual(
          await shapeText.evaluate((field) => [field.selectionStart, field.selectionEnd]),
          [0, 5],
        );
      }
      for (const init of [
        { key: 'e', metaKey: true, shiftKey: true },
        { key: 'e', metaKey: true, isComposing: true },
        { key: 'e', metaKey: true, ctrlKey: true },
      ]) {
        assert.equal(
          await shapeText.evaluate(
            (field, init) =>
              field.dispatchEvent(
                new KeyboardEvent('keydown', { ...init, bubbles: true, cancelable: true }),
              ),
            init,
          ),
          true,
        );
      }
      await shapeText.press('Escape');
      await page
        .locator('.shape-hit.selected')
        .dispatchEvent('keydown', { key: '!', code: 'Digit1', metaKey: true, shiftKey: true });
      await page.getByRole('complementary', { name: 'Format Shape', exact: true }).waitFor();
      const paintPane = page.getByRole('complementary', { name: 'Format Shape', exact: true });
      await paintPane.getByRole('tab', { name: 'Fill & Line', exact: true }).click();
      await paintPane.getByRole('radio', { name: 'Solid fill', exact: true }).check();
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          window.office.getState().editor.slides[0].shapes[0].paint.fill.kind === 'solid',
      );
      await paintPane.getByRole('radio', { name: 'Solid line', exact: true }).check();
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          window.office.getState().editor.slides[0].shapes[0].paint.stroke.kind === 'solid',
      );
      await paintPane.getByLabel('Line color', { exact: true }).fill('#13579b');
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].shapes[0].paint.stroke.color === '#13579B',
      );
      await paintPane.getByLabel('Fill transparency', { exact: true }).fill('25');
      await paintPane.getByLabel('Fill transparency', { exact: true }).press('Tab');
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].shapes[0].paint.fillOpacity === 0.75,
      );
      await paintPane.getByLabel('Line transparency slider', { exact: true }).fill('65');
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].shapes[0].paint.strokeOpacity === 0.35,
      );
      assert.equal(
        await paintPane.getByLabel('Line transparency', { exact: true }).inputValue(),
        '65',
      );
      await paintPane.getByLabel('Line width', { exact: true }).fill('3');
      await paintPane.getByLabel('Line width', { exact: true }).press('Tab');
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].shapes[0].paint.stroke.widthEmu === 38100,
      );
      await paintPane.getByLabel('Dash type', { exact: true }).selectOption('lgDashDot');
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          window.office.getState().editor.slides[0].shapes[0].paint.dash === 'lgDashDot',
      );
      assert.equal(
        await page.evaluate(
          () => window.office.getState().editor.slides[0].shapes[0].paint.stroke.color,
        ),
        '#13579B',
      );
      for (const [label, key, value] of [
        ['Compound type', 'compound', 'dbl'],
        ['Cap type', 'cap', 'sq'],
        ['Join type', 'join', 'bevel'],
      ]) {
        await paintPane.getByLabel(label, { exact: true }).selectOption(value);
        await page.waitForFunction(
          ({ key, value }) =>
            !window.office.getState().building &&
            window.office.getState().editor.slides[0].shapes[0].paint[key] === value,
          { key, value },
        );
      }
      assert.equal(
        await page.evaluate(
          () => window.office.getState().editor.slides[0].shapes[0].paint.strokeOpacity,
        ),
        0.35,
      );
      for (const [label, key, value, option] of [
        ['Begin Arrow type', 'headArrow', 'triangle', 'Triangle Arrow'],
        ['End Arrow type', 'tailArrow', 'oval', 'Oval Arrow'],
      ]) {
        await paintPane.getByRole('button', { name: label, exact: true }).click();
        await page
          .getByRole('listbox', { name: label, exact: true })
          .getByRole('option', { name: option, exact: true })
          .click();
        await page.waitForFunction(
          ({ key, value }) =>
            !window.office.getState().building &&
            window.office.getState().editor.slides[0].shapes[0].paint[key]?.type === value,
          { key, value },
        );
      }
      const beginSize = paintPane.getByRole('button', { name: 'Begin Arrow size', exact: true });
      await beginSize.and(page.locator(':enabled')).waitFor();
      await beginSize.press('ArrowDown');
      const sizeGallery = page.getByRole('listbox', { name: 'Begin Arrow size', exact: true });
      assert.equal(await sizeGallery.getByRole('option').count(), 9);
      assert.equal(await sizeGallery.locator('svg').count(), 9);
      if (process.env.PPTX_UI_SCREENSHOT)
        await page.screenshot({ path: process.env.PPTX_UI_SCREENSHOT });
      await page.keyboard.press('Home');
      await page.keyboard.press('ArrowDown');
      await page.keyboard.press('ArrowDown');
      assert.equal(
        await sizeGallery
          .getByRole('option', { name: 'Wide, Short', exact: true })
          .evaluate((element) => element === document.activeElement),
        true,
      );
      await page.keyboard.press('Enter');
      await page.waitForFunction(() => {
        const arrow = window.office.getState().editor.slides[0].shapes[0].paint.headArrow;
        return !window.office.getState().building && arrow.width === 'lg' && arrow.length === 'sm';
      });
      await beginSize.and(page.locator(':enabled')).waitFor();
      await beginSize.press('ArrowDown');
      assert.equal(
        await sizeGallery
          .getByRole('option', { name: 'Wide, Short', exact: true })
          .getAttribute('aria-selected'),
        'true',
      );
      await page.keyboard.press('End');
      await page.keyboard.press('Escape');
      assert.equal(await sizeGallery.isVisible(), false);
      assert.equal(await beginSize.evaluate((element) => element === document.activeElement), true);
      assert.equal(await beginSize.getAttribute('title'), 'Wide, Short');
      await paintPane.getByRole('button', { name: 'Begin Arrow type', exact: true }).click();
      await page
        .getByRole('listbox', { name: 'Begin Arrow type', exact: true })
        .getByRole('option', { name: 'No Arrow', exact: true })
        .click();
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          window.office.getState().editor.slides[0].shapes[0].paint.headArrow.type === 'none',
      );
      assert.equal(await beginSize.isDisabled(), true);
      await paintPane.getByRole('radio', { name: 'No line', exact: true }).check();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].shapes[0].paint.stroke.kind === 'none',
      );
      assert.equal(await paintPane.getByLabel('Line width', { exact: true }).isDisabled(), true);
      await paintPane.getByRole('tab', { name: 'Fill & Line', exact: true }).press('ArrowRight');
      assert.equal(
        await paintPane
          .getByRole('tab', { name: 'Size & Properties', exact: true })
          .getAttribute('aria-selected'),
        'true',
      );
      await page.getByRole('button', { name: 'Close Format Shape', exact: true }).click();

      await page.getByRole('tab', { name: 'Insert', exact: true }).click();
      await page.locator('[data-edit=table]').click();
      await page.getByRole('gridcell', { name: '3 columns, 2 rows', exact: true }).click();
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          window.office.getState().editor.slides[0].shapes.some((s) => s.table),
      );
      await page.locator('.shape-hit.selected').waitFor();
      await page.getByRole('tab', { name: 'Table Layout', exact: true }).waitFor();
      const first = await page.locator('#stage [data-pptx-cell="0,0"]').first().boundingBox();
      await page.mouse.click(first.x + first.width / 2, first.y + first.height / 2);
      const field = page.getByRole('textbox', { name: 'Edit cell 1, 1', exact: true });
      await field.waitFor();
      await page.getByRole('tab', { name: 'Table Layout', exact: true }).click();
      await field.fill('Keep typed text');
      const choose = async (label) => {
        await page.getByRole('menuitem', { name: 'Select', exact: true }).hover();
        await page.getByRole('menuitem', { name: label, exact: true }).click();
      };
      await field.click({ button: 'right' });
      await choose('Select Row');
      await page.getByLabel('1 rows, 3 columns selected', { exact: true }).waitFor();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.slides[0].shapes.find((s) => s.table)?.table.cells[0][0]
            .text === 'Keep typed text',
      );
      // Keyboard Delete clears cells without removing the table or its grid.
      await page.keyboard.press('Backspace');
      await page.waitForFunction(() => {
        const state = window.office.getState();
        const table = state.editor.slides[0].shapes.find((s) => s.table)?.table;
        return !state.building && table?.cells[0][0].text === '' && table.rowHeights.length === 2;
      });
      await page.getByRole('tab', { name: 'Table Layout', exact: true }).click();
      await page.locator('[data-edit=table-insert-below]').click({ trial: true });
      await page
        .locator('.shape-hit.selected')
        .click({ button: 'right', position: { x: 20, y: 20 } });
      await choose('Select Column');
      // A whole row was selected, so all its columns participate in this selection.
      await page.getByLabel('2 rows, 3 columns selected', { exact: true }).waitFor();
      await page.keyboard.press('Escape');
      const cell = await page.locator('#stage [data-pptx-cell="0,1"]').first().boundingBox();
      await page.mouse.click(cell.x + cell.width / 2, cell.y + cell.height / 2, {
        button: 'right',
      });
      const middle = page.getByRole('textbox', { name: 'Edit cell 1, 2', exact: true });
      await middle.waitFor();
      assert.equal(
        await page.getByRole('menuitem', { name: 'Cut', exact: true }).isDisabled(),
        true,
      );
      assert.equal(
        await page.getByRole('menuitem', { name: 'Copy', exact: true }).isDisabled(),
        true,
      );
      await page.evaluate(() => navigator.clipboard.writeText('Context paste'));
      await page.getByRole('menuitem', { name: 'Paste', exact: true }).click();
      await page.waitForFunction(
        (field) => field.value === 'Context paste',
        await middle.elementHandle(),
      );
      await middle.evaluate((field) => {
        field.focus();
        field.setSelectionRange(0, 7);
      });
      await middle.dispatchEvent('contextmenu', {
        bubbles: true,
        composed: true,
        clientX: 600,
        clientY: 350,
      });
      await page.getByRole('menuitem', { name: 'Cut', exact: true }).click();
      await page.waitForFunction((field) => field.value === ' paste', await middle.elementHandle());
      assert.equal(await page.evaluate(() => navigator.clipboard.readText()), 'Context');
      await page.getByRole('menubar').getByRole('menuitem', { name: 'Edit', exact: true }).click();
      await page.getByRole('menuitem', { name: 'Undo', exact: true }).click();
      assert.equal(await middle.evaluate((field) => field.value), 'Context paste');
      await middle.evaluate((field) => {
        field.focus();
        field.setSelectionRange(0, 7);
      });
      await page
        .getByRole('menubar')
        .getByRole('menuitem', { name: 'Format', exact: true })
        .click();
      await page.getByRole('menuitem', { name: 'Font...', exact: true }).click();
      await page
        .getByRole('dialog', { name: 'Font', exact: true })
        .getByRole('button', { name: 'Cancel', exact: true })
        .click();
      assert.deepEqual(
        await middle.evaluate((field) => [field.selectionStart, field.selectionEnd]),
        [0, 7],
      );
      await page
        .getByRole('menubar')
        .getByRole('menuitem', { name: 'Format', exact: true })
        .click();
      await page.getByRole('menuitem', { name: 'Paragraph...', exact: true }).click();
      await page.keyboard.press('Escape');
      assert.deepEqual(
        await middle.evaluate((field) => [field.selectionStart, field.selectionEnd]),
        [0, 7],
      );
      for (const shortcut of [
        { key: 't', code: 'KeyT', metaKey: true },
        { key: 'µ', code: 'KeyM', metaKey: true, altKey: true },
      ]) {
        const prevented = await middle.evaluate(
          (field, init) =>
            !field.dispatchEvent(
              new KeyboardEvent('keydown', {
                ...init,
                bubbles: true,
                composed: true,
                cancelable: true,
              }),
            ),
          shortcut,
        );
        assert.equal(prevented, true);
        await page
          .getByRole('dialog', { name: shortcut.altKey ? 'Paragraph' : 'Font', exact: true })
          .waitFor();
        await page.keyboard.press('Escape');
        assert.deepEqual(
          await middle.evaluate((field) => [field.selectionStart, field.selectionEnd]),
          [0, 7],
        );
      }
      for (const [key, level] of [
        [']', 1],
        ['[', 0],
      ]) {
        await middle.dispatchEvent('keydown', { key, metaKey: true });
        await page.waitForFunction((level) => {
          const state = window.office.getState();
          return (
            !state.building &&
            state.editor.slides[0].shapes.find((s) => s.table).table.cells[0][1].paragraphs[0]
              .properties.level === level
          );
        }, level);
        await middle.waitFor();
        assert.deepEqual(
          await middle.evaluate((field) => [field.selectionStart, field.selectionEnd]),
          [0, 7],
        );
      }
      for (const [key, expected] of [
        ['e', 'center'],
        ['j', 'justify'],
        ['r', 'right'],
        ['l', 'left'],
      ]) {
        await middle.dispatchEvent('keydown', { key, metaKey: true });
        await page.waitForFunction((align) => {
          const state = window.office.getState();
          const cell = state.editor.slides[0].shapes.find((s) => s.table).table.cells[0][1];
          return !state.building && cell.paragraphs[0].properties.align === align;
        }, expected);
        await middle.waitFor();
        assert.deepEqual(
          await middle.evaluate((field) => [field.selectionStart, field.selectionEnd]),
          [0, 7],
        );
      }
      for (const [shiftKey, baseline] of [
        [false, -0.25],
        [false, 0],
        [true, 0.3],
        [true, 0],
      ]) {
        await middle.dispatchEvent('keydown', {
          key: shiftKey ? '+' : '=',
          code: 'Equal',
          metaKey: true,
          ctrlKey: true,
          shiftKey,
        });
        await page.waitForFunction((baseline) => {
          const state = window.office.getState();
          const cell = state.editor.slides[0].shapes.find((s) => s.table).table.cells[0][1];
          const runs = cell.paragraphs[0].elements.filter((run) => run.text);
          return !state.building && runs[0].format.baseline === baseline;
        }, baseline);
        await middle.waitFor();
        assert.deepEqual(
          await middle.evaluate((field) => [field.selectionStart, field.selectionEnd]),
          [0, 7],
        );
      }
      const shortcutDeck = await pptx.loadPresentation(
        await (await page.request.get(url + '/deck.pptx')).body(),
      );
      const shortcutRuns = pptx
        .getTableCellParagraphs(
          pptx.getTableCell(pptx.getSlideTables(pptx.getSlides(shortcutDeck)[0])[0], 0, 1),
        )[0]
        .elements.filter((run) => run.text);
      assert.equal(shortcutRuns[0].text, 'Context');
      assert.equal(shortcutRuns[0].format.baseline, 0);
      assert.equal(shortcutRuns.at(-1).text, ' paste');
      assert.equal(shortcutRuns.at(-1).format.baseline ?? 0, 0);
      const openFont = async () => {
        await middle.dispatchEvent('contextmenu', {
          bubbles: true,
          composed: true,
          clientX: 600,
          clientY: 350,
        });
        await page.getByRole('menuitem', { name: 'Font...', exact: true }).click();
        return page.getByRole('dialog', { name: 'Font', exact: true });
      };
      let font = await openFont();
      await page.screenshot({ path: '/tmp/pptx-font-dialog.png' });
      await font.getByRole('spinbutton', { name: 'Size', exact: true }).fill('36');
      await font.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.deepEqual(
        await middle.evaluate((field) => [field.selectionStart, field.selectionEnd]),
        [0, 7],
      );
      font = await openFont();
      await font.getByRole('spinbutton', { name: 'Size', exact: true }).fill('24');
      await font.getByRole('combobox', { name: 'Font style', exact: true }).selectOption('both');
      await font.getByRole('combobox', { name: 'Position', exact: true }).selectOption('0.3');
      await font.getByRole('spinbutton', { name: 'Offset', exact: true }).fill('40');
      await font.getByRole('tab', { name: 'Character Spacing', exact: true }).click();
      await font.getByRole('combobox', { name: 'Spacing', exact: true }).selectOption('condensed');
      await font.getByRole('spinbutton', { name: 'By', exact: true }).fill('1.5');
      await font.getByRole('checkbox', { name: 'Kerning for fonts', exact: true }).check();
      await font.getByRole('spinbutton', { name: 'Points and above', exact: true }).fill('14');
      await font.getByRole('button', { name: 'OK', exact: true }).click();
      await page.locator('.font-dialog').waitFor({ state: 'detached' });
      await middle.waitFor();
      assert.deepEqual(
        await middle.evaluate((field) => [field.selectionStart, field.selectionEnd]),
        [0, 7],
      );
      assert.equal(
        await middle
          .locator('span')
          .filter({ hasText: 'Context' })
          .evaluate((span) => getComputedStyle(span).fontKerning),
        'normal',
      );
      const renderedFormat = await middle
        .locator('span')
        .filter({ hasText: 'Context' })
        .evaluate((span) => {
          const style = getComputedStyle(span);
          return { size: parseFloat(style.fontSize), shift: parseFloat(style.verticalAlign) };
        });
      assert.ok(Math.abs(renderedFormat.shift / renderedFormat.size - 0.4 / 0.65) < 0.001);
      const formattedDeck = await pptx.loadPresentation(
        await (await page.request.get(url + '/deck.pptx')).body(),
      );
      const formattedCell = pptx.getTableCell(
        pptx.getSlideTables(pptx.getSlides(formattedDeck)[0])[0],
        0,
        1,
      );
      const paragraphs = pptx.getTableCellParagraphs(formattedCell);
      const runs = paragraphs.flatMap((paragraph) => paragraph.elements).filter((run) => run.text);
      assert.equal(runs[0].text, 'Context');
      assert.equal(runs[0].format.size, 24);
      assert.equal(runs[0].format.bold, true);
      assert.equal(runs[0].format.italic, true);
      assert.equal(runs[0].format.baseline, 0.4);
      assert.equal(runs[0].format.spc, -150);
      assert.equal(runs[0].format.kern, 1400);
      assert.equal(runs.at(-1).text, ' paste');
      assert.notEqual(runs.at(-1).format.size, 24);
      await page.getByRole('tab', { name: 'Home', exact: true }).click();
      await page.locator('[data-edit=change-case]').click();
      await page.getByRole('menuitem', { name: 'UPPERCASE', exact: true }).click();
      await page.waitForFunction(() => {
        const state = window.office.getState();
        return (
          !state.building &&
          state.editor.slides[0].shapes.find((s) => s.table).table.cells[0][1].text ===
            'CONTEXT paste'
        );
      });
      await middle.waitFor();
      assert.deepEqual(
        await middle.evaluate((field) => [field.selectionStart, field.selectionEnd]),
        [0, 7],
      );
      const caseDeck = await pptx.loadPresentation(
        await (await page.request.get(url + '/deck.pptx')).body(),
      );
      const caseRuns = pptx
        .getTableCellParagraphs(
          pptx.getTableCell(pptx.getSlideTables(pptx.getSlides(caseDeck)[0])[0], 0, 1),
        )[0]
        .elements.filter((run) => run.text);
      assert.equal(caseRuns.map((run) => run.text).join(''), 'CONTEXT paste');
      for (const run of caseRuns.filter((run) => !run.text.includes('paste'))) {
        assert.equal(run.format.size, 24);
        assert.equal(run.format.bold, true);
        assert.equal(run.format.baseline, 0.4);
      }
      await page.getByRole('menubar').getByRole('menuitem', { name: 'Edit', exact: true }).click();
      await page.getByRole('menuitem', { name: 'Undo', exact: true }).click();
      await middle.waitFor();
      assert.equal(await middle.evaluate((field) => field.value), 'Context paste');
      for (const [label, expected] of [
        ['Capitalize Each Word', 'Context Paste'],
        ['tOGGLE cASE', 'cONTEXT pASTE'],
        ['Sentence case.', 'Context paste'],
      ]) {
        await middle.evaluate((field) => {
          field.focus();
          field.setSelectionRange(0, field.value.length);
        });
        await page.locator('[data-edit=change-case]').click();
        await page.getByRole('menuitem', { name: label, exact: true }).click();
        await page.waitForFunction((expected) => {
          const state = window.office.getState();
          return (
            !state.building &&
            state.editor.slides[0].shapes.find((s) => s.table).table.cells[0][1].text === expected
          );
        }, expected);
        await middle.waitFor();
        assert.deepEqual(
          await middle.evaluate((field) => [field.selectionStart, field.selectionEnd]),
          [0, expected.length],
        );
      }

      await middle.evaluate((field) => {
        field.focus();
        field.setSelectionRange(0, field.value.length);
      });
      font = await openFont();
      assert.equal(await font.getByLabel('Size', { exact: true }).inputValue(), '');
      assert.equal(await font.getByLabel('Font style', { exact: true }).inputValue(), '');
      await font.getByRole('button', { name: 'Cancel', exact: true }).click();
      await middle.evaluate((field) => {
        field.focus();
        field.setSelectionRange(0, 7);
      });
      const openParagraph = async () => {
        await middle.dispatchEvent('contextmenu', {
          bubbles: true,
          composed: true,
          clientX: 600,
          clientY: 350,
        });
        await page.getByRole('menuitem', { name: 'Paragraph...', exact: true }).click();
        return page.getByRole('dialog', { name: 'Paragraph', exact: true });
      };
      let paragraphDialog = await openParagraph();
      await paragraphDialog.getByRole('spinbutton', { name: 'Before', exact: true }).fill('12');
      await paragraphDialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.deepEqual(
        await middle.evaluate((field) => [field.selectionStart, field.selectionEnd]),
        [0, 7],
      );
      paragraphDialog = await openParagraph();
      await paragraphDialog.getByRole('spinbutton', { name: 'Before', exact: true }).fill('6');
      await paragraphDialog.getByRole('spinbutton', { name: 'After', exact: true }).fill('8');
      await paragraphDialog
        .getByRole('spinbutton', { name: 'Before text', exact: true })
        .fill('0.2');
      await paragraphDialog
        .getByRole('combobox', { name: 'Special', exact: true })
        .selectOption('hanging');
      await paragraphDialog.getByRole('spinbutton', { name: 'By', exact: true }).fill('0.1');
      await paragraphDialog
        .getByRole('combobox', { name: 'Line spacing', exact: true })
        .selectOption('pts');
      await paragraphDialog.getByRole('spinbutton', { name: 'At', exact: true }).fill('30');
      await paragraphDialog.getByRole('button', { name: 'OK', exact: true }).click();
      await page.locator('.paragraph-dialog').waitFor({ state: 'detached' });
      assert.deepEqual(
        await middle.evaluate((field) => [field.selectionStart, field.selectionEnd]),
        [0, 7],
      );
      const paragraphDeck = await pptx.loadPresentation(
        await (await page.request.get(url + '/deck.pptx')).body(),
      );
      const paragraphCell = pptx.getTableCell(
        pptx.getSlideTables(pptx.getSlides(paragraphDeck)[0])[0],
        0,
        1,
      );
      const editedParagraph = pptx.getTableCellParagraphs(paragraphCell)[0];
      assert.equal(editedParagraph.properties.marL, 182880);
      assert.equal(editedParagraph.properties.indent, -91440);
      assert.equal(editedParagraph.properties.spcBefPts, 6);
      assert.equal(editedParagraph.properties.spcAftPts, 8);
      assert.deepEqual(editedParagraph.properties.lineSpacing, { kind: 'pts', value: 30 });
      assert.equal(editedParagraph.elements[0].format.size, 24);
      await middle.click({ button: 'right' });
      await choose('Select Column');
      await page.getByRole('tab', { name: 'Home', exact: true }).click();
      await page.locator('[data-edit=demote]').click();
      await page.waitForFunction(() => {
        const state = window.office.getState();
        const cells = state.editor.slides[0].shapes.find((s) => s.table).table.cells;
        return !state.building && cells.every((row) => row[1].paragraphs[0].properties.level === 1);
      });
      await page.getByLabel('2 rows, 1 columns selected', { exact: true }).waitFor();
      await page
        .locator('.shape-hit.selected')
        .click({ button: 'right', position: { x: 20, y: 20 } });
      await page.getByRole('menuitem', { name: 'Paragraph...', exact: true }).click();
      const mixedDialog = page.getByRole('dialog', { name: 'Paragraph', exact: true });
      assert.equal(await mixedDialog.getByLabel('Before', { exact: true }).inputValue(), '');
      assert.equal(await mixedDialog.getByLabel('After', { exact: true }).inputValue(), '');
      assert.equal(await mixedDialog.getByLabel('Line spacing', { exact: true }).inputValue(), '');
      await mixedDialog.getByLabel('After', { exact: true }).fill('10');
      await mixedDialog.getByRole('button', { name: 'OK', exact: true }).click();
      await page.locator('.paragraph-dialog').waitFor({ state: 'detached' });
      await page.getByLabel('2 rows, 1 columns selected', { exact: true }).waitFor();
      const mixedDeck = await pptx.loadPresentation(
        await (await page.request.get(url + '/deck.pptx')).body(),
      );
      const mixedTable = pptx.getSlideTables(pptx.getSlides(mixedDeck)[0])[0];
      for (const row of [0, 1]) {
        const properties = pptx.getTableCellParagraphs(pptx.getTableCell(mixedTable, row, 1))[0]
          .properties;
        assert.equal(properties.spcAftPts, 10);
        assert.equal(properties.level, 1);
        assert.equal(properties.spcBefPts ?? 0, row === 0 ? 6 : 0);
      }

      await page
        .locator('.shape-hit.selected')
        .click({ button: 'right', position: { x: 20, y: 20 } });
      await page.getByRole('menuitem', { name: 'Delete', exact: true }).hover();
      await page.getByRole('menuitem', { name: 'Delete Columns', exact: true }).click();
      await page.waitForFunction(() => {
        const state = window.office.getState();
        return (
          !state.building &&
          state.editor.slides[0].shapes.find((s) => s.table)?.table.columnWidths.length === 2
        );
      });
      await page.getByRole('tab', { name: 'Table Layout', exact: true }).click();
      await page.locator('[data-edit=table-insert-below]').click({ trial: true });
      await page.locator('[data-edit=table-select]').click();
      await page.getByRole('menuitem', { name: 'Select Table', exact: true }).click();
      await page.getByLabel('2 rows, 2 columns selected', { exact: true }).waitFor();
      await page.locator('[data-edit=table-insert-below]').click();
      await page.waitForFunction(() => {
        const state = window.office.getState();
        return (
          !state.building &&
          state.editor.slides[0].shapes.find((s) => s.table)?.table.rowHeights.length === 4
        );
      });
      await page.getByRole('textbox', { name: 'Edit cell 3, 1', exact: true }).waitFor();
      const downloaded = await pptx.loadPresentation(
        await (await page.request.get(url + '/deck.pptx')).body(),
      );
      const table = pptx.getSlideTables(pptx.getSlides(downloaded)[0])[0];
      assert.equal(pptx.getTableColumnWidths(table).length, 2);
      assert.equal(pptx.getTableRowHeights(table).length, 4);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      proc.kill('SIGTERM');
      await new Promise((resolve) => {
        if (proc.exitCode !== null) resolve();
        else proc.once('exit', resolve);
      });
      await rm(dir, { recursive: true, force: true });
    }
  },
);

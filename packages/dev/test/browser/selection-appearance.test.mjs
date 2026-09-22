import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeFillColor,
  getShapeFill,
  getShapeStroke,
  getShapeStrokeColor,
  getSlides,
  getSlideShapes,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'fill and outline apply to all selected shapes in English and Japanese',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-selection-appearance-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={2} height={1}>日本語</Text><Text x={4} y={1} width={2} height={1}>English</Text><Text x={1} y={3} width={2} height={1}>Untouched</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      let ja = false;
      const saved = () =>
        editor
          .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
          .waitFor();
      const colors = async (reader) => {
        const deck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getSlideShapes(getSlides(deck)[0]).map(reader);
      };
      await saved();
      const initialFill = await colors(getShapeFillColor);
      const initialStroke = await colors(getShapeStrokeColor);
      await editor.locator('.hit').nth(0).click();
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      assert.equal(await editor.locator('.hit.selected').count(), 2);
      const changeColor = async (index, color) => {
        await editor
          .locator('.bespoke input[type=color]')
          .nth(index)
          .evaluate((node, value) => {
            node.value = value;
            node.dispatchEvent(new Event('input', { bubbles: true }));
            node.dispatchEvent(new Event('change', { bubbles: true }));
          }, color);
        await saved();
      };
      await changeColor(0, '#123456');
      assert.deepEqual(await colors(getShapeFillColor), ['#123456', '#123456', initialFill[2]]);
      assert.equal(
        await editor.locator('.bespoke input[type=color]').nth(0).inputValue(),
        '#123456',
      );
      await editor.locator('.hit').nth(2).click();
      assert.notEqual(
        await editor.locator('.bespoke input[type=color]').nth(0).inputValue(),
        '#123456',
      );
      await editor
        .locator('.hit')
        .nth(0)
        .click({ modifiers: ['Shift'] });
      assert.equal(await editor.locator('[data-paint-state=fill]').textContent(), 'Mixed');
      await editor
        .locator('.hit')
        .nth(2)
        .click({ modifiers: ['Shift'] });
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await colors(getShapeFillColor), initialFill);
      assert.notEqual(
        await editor.locator('.bespoke input[type=color]').nth(0).inputValue(),
        '#123456',
      );
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual(await colors(getShapeFillColor), ['#123456', '#123456', initialFill[2]]);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await editor.locator('.hit').nth(2).click();
      await editor
        .locator('.hit')
        .nth(0)
        .click({ modifiers: ['Shift'] });
      assert.equal(await editor.locator('[data-paint-state=fill]').textContent(), '混在');
      await editor
        .locator('.hit')
        .nth(2)
        .click({ modifiers: ['Shift'] });
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      await changeColor(1, '#abcdef');
      assert.deepEqual(await colors(getShapeStrokeColor), ['#ABCDEF', '#ABCDEF', initialStroke[2]]);
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await colors(getShapeStrokeColor), initialStroke);
      await changeColor(1, '#abcdef');
      await page.reload();
      await saved();
      await editor.locator('.hit').nth(0).click();
      assert.equal(
        await editor.locator('.bespoke input[type=color]').nth(0).inputValue(),
        '#123456',
      );
      assert.equal(
        await editor.locator('.bespoke input[type=color]').nth(1).inputValue(),
        '#abcdef',
      );
      assert.deepEqual(await colors(getShapeFillColor), ['#123456', '#123456', initialFill[2]]);
      assert.deepEqual(await colors(getShapeStrokeColor), ['#ABCDEF', '#ABCDEF', initialStroke[2]]);
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      const originalFillKinds = await colors((shape) => getShapeFill(shape).kind);
      const originalStrokeKinds = await colors((shape) => getShapeStroke(shape).kind);
      await editor
        .locator('.bespoke')
        .getByRole('button', { name: '塗りつぶしなし', exact: true })
        .click();
      await saved();
      assert.deepEqual(await colors((shape) => getShapeFill(shape).kind), [
        'none',
        'none',
        originalFillKinds[2],
      ]);
      assert.equal(await editor.locator('[data-paint-state=fill]').textContent(), 'なし');
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await colors(getShapeFillColor), ['#123456', '#123456', initialFill[2]]);
      await editor.locator('.lang select').selectOption('en');
      ja = false;
      await editor
        .locator('.bespoke')
        .getByRole('button', { name: 'No fill', exact: true })
        .click();
      await saved();
      await editor
        .locator('.bespoke')
        .getByRole('button', { name: 'No outline', exact: true })
        .click();
      await saved();
      assert.deepEqual(await colors((shape) => getShapeStroke(shape).kind), [
        'none',
        'none',
        originalStrokeKinds[2],
      ]);
      assert.equal(await editor.locator('[data-paint-state=stroke]').textContent(), 'None');
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await colors(getShapeStrokeColor), ['#ABCDEF', '#ABCDEF', initialStroke[2]]);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      assert.deepEqual(await colors((shape) => getShapeFill(shape).kind), [
        'none',
        'none',
        originalFillKinds[2],
      ]);
      assert.deepEqual(await colors((shape) => getShapeStroke(shape).kind), [
        'none',
        'none',
        originalStrokeKinds[2],
      ]);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

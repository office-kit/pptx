import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeBounds,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'single click enters text editing without shifting the stage or nearby shapes',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-text-edit-entry-'));
    let browser;
    let preview;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={4} height={1}>Editable text</Text><Text x={6} y={1} width={2} height={1}>Neighbor</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const stage = editor.locator('.stage');
      const hits = editor.locator('.hit');
      const glyph = editor
        .locator('.paint foreignObject > div')
        .filter({ hasText: 'Editable text' })
        .first();
      const hitRect = async (index) => hits.nth(index).boundingBox();
      const layout = async () => ({
        stage: await stage.boundingBox(),
        first: await hitRect(0),
        second: await hitRect(1),
      });
      const before = await layout();
      assert.ok(before.stage && before.first && before.second);
      await page.screenshot({ path: '/tmp/text-edit-entry-before.png', fullPage: true });

      // Click actual glyphs rather than the center of the full-width text container.
      const glyphBox = await glyph.boundingBox();
      assert.ok(glyphBox);
      const glyphPoint = await glyph.evaluate((node) => {
        const walker = node.ownerDocument.createTreeWalker(node, NodeFilter.SHOW_TEXT);
        const text = walker.nextNode();
        if (!text) throw new Error('editable glyph has no text node');
        const range = node.ownerDocument.createRange();
        range.setStart(text, 1);
        range.setEnd(text, Math.min(5, text.textContent.length));
        const rect = range.getBoundingClientRect();
        const box = node.getBoundingClientRect();
        return {
          x: box.left + (rect.left - box.left) + rect.width / 2,
          y: box.top + (rect.top - box.top) + rect.height / 2,
          boxX: box.left,
          boxY: box.top,
          rawX: rect.left,
          rawY: rect.top,
          width: rect.width,
          height: rect.height,
        };
      });
      const glyphTextBox = {
        x: glyphPoint.x - glyphPoint.boxX,
        y: glyphPoint.y - glyphPoint.boxY,
        width: glyphPoint.width,
        height: glyphPoint.height,
      };
      await hits.nth(0).click({
        position: {
          x: glyphBox.x - before.first.x + glyphPoint.x - glyphPoint.boxX,
          y: glyphBox.y - before.first.y + glyphPoint.y - glyphPoint.boxY,
        },
      });
      await page.screenshot({ path: '/tmp/text-edit-entry-after-click.png', fullPage: true });
      const textEditor = editor.locator('.canvas-shell .inline-edit').first();
      await textEditor.waitFor({ timeout: 3000 });
      assert.equal(
        await editor.getByRole('spinbutton', { name: 'Font size', exact: true }).inputValue(),
        '18',
      );
      await page.screenshot({ path: '/tmp/text-edit-entry.png', fullPage: true });
      const caret = await textEditor.evaluate((node) => {
        const selection = node.ownerDocument.getSelection();
        const range = selection?.rangeCount ? selection.getRangeAt(0) : null;
        const rect = range?.getBoundingClientRect();
        return {
          collapsed: selection?.isCollapsed ?? false,
          offset: selection?.anchorOffset ?? 0,
          width: rect?.width ?? Number.NaN,
        };
      });
      assert.equal(caret.collapsed, true);
      assert.ok(caret.offset > 0 && caret.offset < 'Editable text'.length);
      assert.ok(caret.width < 2, `caret width was ${caret.width}`);
      const textInputBox = await textEditor.boundingBox();
      assert.ok(textInputBox);
      const editTextSize = await textEditor.evaluate((node) => {
        const walker = node.ownerDocument.createTreeWalker(node, NodeFilter.SHOW_TEXT);
        const text = walker.nextNode();
        if (!text) throw new Error('inline editor has no text node');
        const range = node.ownerDocument.createRange();
        range.setStart(text, 1);
        range.setEnd(text, Math.min(5, text.textContent.length));
        const rect = range.getBoundingClientRect();
        const box = node.getBoundingClientRect();
        return {
          rawX: rect.left,
          rawY: rect.top,
          x: rect.left - box.left,
          y: rect.top - box.top,
          width: rect.width,
          height: rect.height,
        };
      });
      assert.ok(Math.abs(editTextSize.rawX - glyphPoint.rawX) < 2, 'glyph x shifted on edit');
      assert.ok(Math.abs(editTextSize.rawY - glyphPoint.rawY) < 2, 'glyph y shifted on edit');
      assert.ok(
        Math.abs(editTextSize.width - glyphTextBox.width) < 2,
        'glyph width shifted on edit',
      );
      assert.ok(
        Math.abs(editTextSize.height - glyphTextBox.height) < 2,
        'glyph height shifted on edit',
      );
      const afterEntry = await layout();
      for (const key of ['x', 'y', 'width', 'height']) {
        assert.ok(
          Math.abs(afterEntry.stage[key] - before.stage[key]) < 0.5,
          `${key} stage shifted`,
        );
        assert.ok(
          Math.abs(afterEntry.second[key] - before.second[key]) < 0.5,
          `${key} neighbor shifted`,
        );
      }
      assert.ok(Math.abs(afterEntry.first.x - before.first.x) < 0.5);
      assert.ok(Math.abs(afterEntry.first.y - before.first.y) < 0.5);

      const editorRect = textInputBox;
      const wordPoint = await textEditor.evaluate((node) => {
        const walker = node.ownerDocument.createTreeWalker(node, NodeFilter.SHOW_TEXT);
        const text = walker.nextNode();
        if (!text) throw new Error('inline editor has no text node');
        const range = node.ownerDocument.createRange();
        range.setStart(text, 1);
        range.setEnd(text, Math.min(5, text.textContent.length));
        const rect = range.getBoundingClientRect();
        const box = node.getBoundingClientRect();
        return {
          x: rect.left - box.left + rect.width / 2,
          y: rect.top - box.top + rect.height / 2,
        };
      });
      await page.mouse.dblclick(editorRect.x + wordPoint.x, editorRect.y + wordPoint.y);
      const word = await textEditor.evaluate((node) => {
        const selection = node.ownerDocument.getSelection();
        return { collapsed: selection?.isCollapsed ?? true, text: selection?.toString() ?? '' };
      });
      assert.equal(word.collapsed, false);
      assert.ok(word.text.length > 0 && word.text.length < 'Editable text'.length);

      await textEditor.press('Escape');
      await textEditor.waitFor({ state: 'hidden' });
      const beforeMove = await layout();
      const point = {
        x: beforeMove.first.x + beforeMove.first.width / 4,
        y: beforeMove.first.y + 2,
      };
      await page.mouse.move(point.x, point.y);
      await page.mouse.down();
      await page.mouse.move(point.x + 30, point.y + 18, { steps: 5 });
      await page.mouse.up();
      const afterMove = await layout();
      assert.ok(Math.abs(afterMove.stage.x - beforeMove.stage.x) < 0.5);
      assert.ok(Math.abs(afterMove.stage.y - beforeMove.stage.y) < 0.5);
      assert.ok(Math.abs(afterMove.second.x - beforeMove.second.x) < 0.5);
      assert.ok(Math.abs(afterMove.second.y - beforeMove.second.y) < 0.5);
      assert.ok(afterMove.first.x > beforeMove.first.x + 10);
      assert.ok(afterMove.first.y > beforeMove.first.y + 5);

      let moved;
      for (let attempt = 0; attempt < 20; attempt += 1) {
        const saved = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        moved = getShapeBounds(getSlideShapes(getSlides(saved)[0])[0]);
        if (moved.x > inches(1)) break;
        await new Promise((resolve) => setTimeout(resolve, 100));
      }
      assert.ok(moved);
      assert.ok(moved.x > inches(1));
      assert.ok(moved.y > inches(1));
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

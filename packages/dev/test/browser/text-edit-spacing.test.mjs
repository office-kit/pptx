import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { startPreview } from '../helpers/server.mjs';

function textRange(locator) {
  return locator.evaluate((node) => {
    const text = node.ownerDocument.createTreeWalker(node, NodeFilter.SHOW_TEXT).nextNode();
    if (!text) throw new Error('text shape has no text node');
    const range = node.ownerDocument.createRange();
    range.setStart(text, 0);
    range.setEnd(text, Math.min(5, text.textContent.length));
    const rect = range.getBoundingClientRect();
    return {
      x: rect.left,
      y: rect.top,
      width: rect.width,
      height: rect.height,
    };
  });
}

function glyphPoint(locator) {
  return locator.evaluate((node) => {
    const text = node.ownerDocument.createTreeWalker(node, NodeFilter.SHOW_TEXT).nextNode();
    if (!text) throw new Error('text shape has no text node');
    const range = node.ownerDocument.createRange();
    range.setStart(text, 1);
    range.setEnd(text, Math.min(4, text.textContent.length));
    const rect = range.getBoundingClientRect();
    const box = node.getBoundingClientRect();
    return { x: rect.left - box.left + rect.width / 2, y: rect.top - box.top + rect.height / 2 };
  });
}

test(
  'text editing preserves positive and negative character spacing',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-text-edit-spacing-'));
    let browser;
    let preview;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={4} height={1} paragraphs={[{runs:[{text:'Wide',format:{size:24,spc:300}}]}]}/><Text x={1} y={2.5} width={4} height={1} paragraphs={[{runs:[{text:'Tight',format:{size:24,spc:-150}}]}]}/></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const hits = editor.locator('.hit');
      for (const zoom of ['100%', '200%']) {
        await editor.getByTitle('Zoom...', { exact: true }).click();
        const dialog = editor.getByRole('dialog', { name: 'Zoom', exact: true });
        await dialog.getByRole('radio', { name: zoom, exact: true }).check();
        await dialog.getByRole('button', { name: 'OK', exact: true }).click();
        for (const [label, index] of [
          ['Wide', 0],
          ['Tight', 1],
        ]) {
          const glyph = editor
            .locator('.paint foreignObject > div')
            .filter({ hasText: label })
            .first();
          const hit = hits.nth(index);
          const hitBox = await hit.boundingBox();
          const glyphBox = await glyph.boundingBox();
          const stageBefore = await editor.locator('.stage').boundingBox();
          assert.ok(hitBox && glyphBox && stageBefore, `missing geometry for ${label}`);
          const before = await textRange(glyph);
          const point = await glyphPoint(glyph);
          await hit.click({
            position: { x: glyphBox.x - hitBox.x + point.x, y: glyphBox.y - hitBox.y + point.y },
          });
          const input = editor.locator('.canvas-shell .inline-edit').first();
          await input.waitFor({ timeout: 3000 });
          const stageAfter = await editor.locator('.stage').boundingBox();
          assert.ok(stageAfter, `missing stage after entering ${label}`);
          assert.ok(Math.abs(stageAfter.x - stageBefore.x) < 2, `${label}: stage x shifted`);
          assert.ok(Math.abs(stageAfter.y - stageBefore.y) < 2, `${label}: stage y shifted`);
          assert.ok(
            Math.abs(stageAfter.width - stageBefore.width) < 2,
            `${label}: stage width shifted`,
          );
          assert.ok(
            Math.abs(stageAfter.height - stageBefore.height) < 2,
            `${label}: stage height shifted`,
          );
          const after = await textRange(input);
          assert.ok(
            Math.abs(after.x - before.x) < 2,
            `${label}: x shifted (before=${before.x}, after=${after.x}, zoom=${zoom})`,
          );
          assert.ok(
            Math.abs(after.y - before.y) < 2,
            `${label}: y shifted (before=${before.y}, after=${after.y}, zoom=${zoom})`,
          );
          assert.ok(
            Math.abs(after.width - before.width) < 2,
            `${label}: character spacing lost (before=${before.width}, after=${after.width})`,
          );
          assert.ok(Math.abs(after.height - before.height) < 2, `${label}: height shifted`);
          await input.press('Escape');
          await input.waitFor({ state: 'hidden' });
        }
      }
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

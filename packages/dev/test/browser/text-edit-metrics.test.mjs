import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { startPreview } from '../helpers/server.mjs';

function rangeMetrics(locator, textIndex = 0) {
  return locator.evaluate((node, textIndex) => {
    const walker = node.ownerDocument.createTreeWalker(node, NodeFilter.SHOW_TEXT);
    let text;
    for (let i = 0; i <= textIndex; i += 1) {
      do {
        text = walker.nextNode();
      } while (text && !text.textContent.trim());
    }
    if (!text) throw new Error('text shape has no text node');
    const end = Math.min(4, text.textContent.length);
    const range = node.ownerDocument.createRange();
    range.setStart(text, 0);
    range.setEnd(text, end);
    const rect = range.getBoundingClientRect();
    return { x: rect.left, y: rect.top, width: rect.width, height: rect.height };
  }, textIndex);
}

function glyphClickPoint(locator, textIndex = 0) {
  return locator.evaluate((node, textIndex) => {
    const walker = node.ownerDocument.createTreeWalker(node, NodeFilter.SHOW_TEXT);
    let text;
    for (let i = 0; i <= textIndex; i += 1) {
      do {
        text = walker.nextNode();
      } while (text && !text.textContent.trim());
    }
    if (!text) throw new Error('text shape has no text node');
    const range = node.ownerDocument.createRange();
    range.setStart(text, 1);
    range.setEnd(text, Math.min(4, text.textContent.length));
    const rect = range.getBoundingClientRect();
    const box = node.getBoundingClientRect();
    return { x: rect.left - box.left + rect.width / 2, y: rect.top - box.top + rect.height / 2 };
  }, textIndex);
}

test(
  'text editing preserves glyph geometry across font sizes, runs, and lines',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-text-edit-metrics-'));
    let browser;
    let preview;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={0.5} width={2.7} height={1} size={10}>Ten px</Text><Text x={4.2} y={0.5} width={2.7} height={1} size={12}>Twelve px</Text><Text x={7.4} y={0.5} width={3.2} height={1} size={24}>Twenty four</Text><Text x={1} y={2.1} width={4.5} height={1.1} paragraphs={[{runs:[{text:'Bold ',format:{size:24,bold:true}},{text:'mixed',format:{size:12}}]}]}/><Text x={6.2} y={2.1} width={4.5} height={1.5} paragraphs={[{runs:[{text:'Line one',format:{size:44}}]},{runs:[{text:'Line two',format:{size:44}}]}]}/><Text x={1} y={4.1} width={4.5} height={1} paragraphs={[{runs:[{text:'Double underline',format:{size:28,underline:'dbl'}}]}]}/><Text x={6.2} y={4.1} width={4.5} height={1} paragraphs={[{runs:[{text:'Wavy underline',format:{size:28,underline:'wavy',strike:true}}]}]}/></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const hits = editor.locator('.hit');
      const cases = [
        ['Ten px', 0, 0],
        ['Twelve px', 1, 0],
        ['Twenty four', 2, 0],
        ['Bold mixed', 3, 0],
        ['Bold mixed', 3, 1],
        ['Line one', 4, 0],
        ['Line one', 4, 1],
        ['Double underline', 5, 0, 'double'],
        ['Wavy underline', 6, 0, 'wavy'],
      ];
      for (const [label, index, textIndex, decorationStyle] of cases) {
        const glyph = editor
          .locator('.paint foreignObject > div')
          .filter({ hasText: label })
          .first();
        const hit = hits.nth(index);
        const hitBox = await hit.boundingBox();
        const glyphBox = await glyph.boundingBox();
        assert.ok(hitBox && glyphBox, `missing geometry for ${label}`);
        const before = await rangeMetrics(glyph, textIndex);
        const point = await glyphClickPoint(glyph, textIndex);
        await hit.click({
          position: { x: glyphBox.x - hitBox.x + point.x, y: glyphBox.y - hitBox.y + point.y },
        });
        const input = editor.locator('.inline-edit').first();
        await input.waitFor({ timeout: 3000 });
        const after = await rangeMetrics(input, textIndex);
        if (decorationStyle) {
          const decorated = await input
            .locator('span, u')
            .evaluateAll((spans) =>
              spans
                .filter((span) => getComputedStyle(span).textDecorationLine.includes('underline'))
                .map((span) => getComputedStyle(span).textDecorationStyle),
            );
          assert.ok(decorated.includes(decorationStyle), `${label}: style changed on edit entry`);
          if (label === 'Wavy underline') {
            const strikes = await input
              .locator('span, u')
              .evaluateAll((spans) =>
                spans
                  .filter((span) =>
                    getComputedStyle(span).textDecorationLine.includes('line-through'),
                  )
                  .map((span) => getComputedStyle(span).textDecorationStyle),
              );
            assert.deepEqual(
              strikes,
              ['solid'],
              'strike must remain solid alongside wavy underline',
            );
          }
        }
        assert.ok(Math.abs(after.x - before.x) < 2, `${label}: x shifted`);
        assert.ok(
          Math.abs(after.y - before.y) < 2,
          `${label} run ${textIndex}: y shifted ${JSON.stringify({ before, after })}`,
        );
        assert.ok(Math.abs(after.width - before.width) < 2, `${label}: width shifted`);
        assert.ok(Math.abs(after.height - before.height) < 2, `${label}: height shifted`);
        await input.press('Escape');
        await input.waitFor({ state: 'hidden' });
      }
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

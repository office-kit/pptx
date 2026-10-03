import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  addBlankSlide,
  addSlideTable,
  addSlideTextBox,
  createPresentation,
  getTableCells,
  groupShapes,
  inches,
  savePresentation,
  setShapeTextAnchor,
  setShapeTextBodyRotationDeg,
  setShapeTextFormat,
  setShapeTextDirection,
  setShapeTextMargins,
  setShapeBounds,
  setTableCellTextFormat,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

for (const direction of ['horz', 'vert270'])
  test(
    `inline character effects track live ${direction} text at canvas zoom and survive save/reopen`,
    { timeout: 60000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-inline-effects-'));
      let preview, browser;
      try {
        const pres = createPresentation();
        const slide = addBlankSlide(pres);
        const shape = addSlideTextBox(slide, {
          x: inches(1),
          y: inches(1),
          w: inches(6),
          h: inches(2.5),
          text: 'Shadow\treflection at 150%',
        });
        setShapeTextAnchor(shape, 'center');
        setShapeTextMargins(shape, {
          left: inches(0.35),
          right: inches(0.1),
          top: inches(0.2),
          bottom: inches(0.45),
        });
        setShapeTextBodyRotationDeg(shape, direction === 'horz' ? 15 : 0);
        setShapeTextDirection(shape, direction);
        setShapeTextFormat(shape, {
          reflection: { scaleY: -0.75, startOpacity: 0.55, endPosition: 0.4 },
          innerShadow: { color: '#000000', blurEmu: 38100, offsetEmu: 19050, angleDeg: 225 },
        });
        const source = join(dir, 'source.pptx');
        const file = join(dir, 'deck.tsx');
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
        const zoomIn = editor.getByTitle('Zoom in (Ctrl+=)', { exact: true });
        for (let i = 0; i < 5; i++) await zoomIn.click();
        const paintedText = editor.locator('.paint foreignObject').first();
        const firstWordRect = async (locator) =>
          locator.evaluate((el) => {
            const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
            let node;
            while ((node = walker.nextNode())) {
              const start = node.textContent.indexOf('Shadow');
              if (start < 0) continue;
              const range = document.createRange();
              range.setStart(node, start);
              range.setEnd(node, start + 6);
              const rect = range.getBoundingClientRect();
              return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
            }
            throw new Error('Missing Shadow text');
          });
        const beforeText = await firstWordRect(paintedText);
        const beforeEffect = await editor
          .locator('.paint text[filter*="text-inner-shadow"]')
          .first()
          .boundingBox();
        assert.ok(beforeEffect);
        await page.mouse.click(
          beforeText.x + beforeText.width / 2,
          beforeText.y + beforeText.height / 2,
        );
        const input = editor.locator('.inline-edit');
        await input.waitFor();
        const effects = editor.locator('.inline-effects');
        await effects.waitFor();
        assert.ok((await effects.locator('text[filter*="text-inner-shadow"]').count()) > 0);
        assert.ok((await effects.locator('text[aria-hidden="true"]').count()) > 0);
        assert.equal(
          await editor
            .locator('.paint [data-pptx-shape-id] foreignObject')
            .first()
            .evaluate((el) => getComputedStyle(el).visibility),
          'hidden',
        );

        const afterText = await firstWordRect(input);
        const afterEffect = await effects
          .locator('text[filter*="text-inner-shadow"]')
          .first()
          .boundingBox();
        assert.ok(afterEffect);
        const liveEffectGlyph = effects
          .locator('text[filter*="text-inner-shadow"]')
          .filter({ hasText: 'Shadow' })
          .first();
        const liveEffectGlyphBox = await liveEffectGlyph.boundingBox();
        assert.ok(liveEffectGlyphBox, `${direction} live effect glyph missing`);
        const liveEffectGlyphGeometry = await liveEffectGlyph.evaluate((el) => {
          const box = el.getBBox();
          const matrix = el.getScreenCTM();
          if (!matrix) throw new Error('Missing SVG screen transform');
          const points = [
            new DOMPoint(box.x, box.y),
            new DOMPoint(box.x + box.width, box.y),
            new DOMPoint(box.x, box.y + box.height),
            new DOMPoint(box.x + box.width, box.y + box.height),
          ].map((point) => point.matrixTransform(matrix));
          return {
            x: Math.min(...points.map((point) => point.x)),
            y: Math.min(...points.map((point) => point.y)),
            width:
              Math.max(...points.map((point) => point.x)) -
              Math.min(...points.map((point) => point.x)),
            height:
              Math.max(...points.map((point) => point.y)) -
              Math.min(...points.map((point) => point.y)),
          };
        });
        for (const key of ['x', 'y']) {
          assert.ok(
            Math.abs(afterText[key] - liveEffectGlyphGeometry[key]) < 4,
            `${direction} live effect glyph ${key}: ${afterText[key]} vs ${liveEffectGlyphGeometry[key]}`,
          );
        }
        for (const key of ['x', 'y', 'width', 'height']) {
          assert.ok(
            Math.abs(beforeText[key] - afterText[key]) < 2,
            `${direction} glyph ${key}: ${beforeText[key]} -> ${afterText[key]}`,
          );
          assert.ok(
            Math.abs(beforeEffect[key] - afterEffect[key]) < 2,
            `${direction} effect ${key}: ${beforeEffect[key]} -> ${afterEffect[key]}`,
          );
        }
        await page.screenshot({ path: join(tmpdir(), `inline-live-effects-${direction}.png`) });
        await input.fill('Updated shadow text');
        // SVG line layout can omit the space consumed by a wrap boundary.
        assert.match((await effects.textContent()).replace(/\s/g, ''), /Updatedshadowtext/);
        await input.press('Control+z');
        assert.notEqual((await input.innerText()).trim(), 'Updated shadow text');
        await input.fill('Updated shadow text');
        await input.press('Control+Enter');
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        await page.reload();
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        assert.match(
          (await editor.locator('.paint [data-pptx-paragraph]').first().textContent()).replace(
            /\s/g,
            '',
          ),
          /Updatedshadowtext/,
        );
        const reopenedHit = editor.locator('.hit').first();
        const reopenedBox = await reopenedHit.boundingBox();
        assert.ok(reopenedBox);
        await reopenedHit.dblclick({
          force: true,
          position: { x: reopenedBox.width / 2, y: reopenedBox.height / 2 },
        });
        await editor.locator('.inline-edit').waitFor();
        assert.equal(
          (await editor.locator('.inline-edit').innerText()).trim(),
          'Updated shadow text',
        );
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );

for (const grouped of [false, true])
  test(
    `inline character effects stay visible while editing a table cell (grouped=${grouped})`,
    { timeout: 60000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-inline-table-effects-'));
      let preview, browser;
      try {
        const pres = createPresentation();
        const slide = addBlankSlide(pres);
        const table = addSlideTable(slide, {
          x: inches(1),
          y: inches(1),
          w: inches(6),
          h: inches(2.5),
          rows: [['Cell shadow']],
        });
        if (grouped) {
          const sibling = addSlideTextBox(slide, {
            x: inches(8),
            y: inches(1),
            w: inches(1),
            h: inches(1),
            text: 'Sibling',
          });
          const group = groupShapes([table, sibling]);
          // Deliberately apply nonuniform group scaling so the editor's table
          // shell and the effects SVG must agree through the ancestor matrix.
          setShapeBounds(group, { x: inches(1), y: inches(1), w: inches(7), h: inches(4) });
        }
        const cell = getTableCells(table)[0][0];
        setTableCellTextFormat(cell, {
          reflection: { scaleY: -0.75, startOpacity: 0.55, endPosition: 0.4 },
          innerShadow: { color: '#000000', blurEmu: 38100, offsetEmu: 19050, angleDeg: 225 },
        });
        const source = join(dir, 'source.pptx');
        const file = join(dir, 'deck.tsx');
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
        await editor.locator('.hit').first().dblclick();
        if (grouped) await editor.locator('.hit.selected').dblclick();
        const input = editor.locator('.inline-edit');
        await input.waitFor();
        const effects = editor.locator('.inline-effects');
        await effects.waitFor();
        const liveEffectGlyph = effects
          .locator('text[filter*="text-inner-shadow"]')
          .filter({ hasText: 'Cell shadow' })
          .first();
        assert.ok((await liveEffectGlyph.count()) > 0);
        assert.ok((await effects.locator('text[aria-hidden="true"]').count()) > 0);
        const liveTextRect = await input.evaluate((el) => {
          const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
          let node;
          while ((node = walker.nextNode())) {
            const start = node.textContent.indexOf('Cell shadow');
            if (start < 0) continue;
            const range = document.createRange();
            range.setStart(node, start);
            range.setEnd(node, start + 'Cell shadow'.length);
            const rect = range.getBoundingClientRect();
            return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
          }
          throw new Error('Missing table cell text');
        });
        const liveEffectGlyphGeometry = await liveEffectGlyph.evaluate((el) => {
          const box = el.getBBox();
          const matrix = el.getScreenCTM();
          if (!matrix) throw new Error('Missing SVG screen transform');
          const points = [
            new DOMPoint(box.x, box.y),
            new DOMPoint(box.x + box.width, box.y),
            new DOMPoint(box.x, box.y + box.height),
            new DOMPoint(box.x + box.width, box.y + box.height),
          ].map((point) => point.matrixTransform(matrix));
          return {
            x: Math.min(...points.map((point) => point.x)),
            y: Math.min(...points.map((point) => point.y)),
            width:
              Math.max(...points.map((point) => point.x)) -
              Math.min(...points.map((point) => point.x)),
            height:
              Math.max(...points.map((point) => point.y)) -
              Math.min(...points.map((point) => point.y)),
          };
        });
        for (const key of ['x', 'y']) {
          assert.ok(
            Math.abs(liveTextRect[key] - liveEffectGlyphGeometry[key]) < 4,
            `table live effect glyph ${key}: ${liveTextRect[key]} vs ${liveEffectGlyphGeometry[key]}`,
          );
        }
        await input.fill('Updated cell shadow');
        assert.match((await effects.textContent()).replace(/\s/g, ''), /Updatedcellshadow/);
        const updatedEffectGlyph = effects
          .locator('text[filter*="text-inner-shadow"]')
          .filter({ hasText: 'Updated cell shadow' })
          .first();
        assert.ok((await updatedEffectGlyph.count()) > 0);
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );

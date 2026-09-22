import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  createPresentation,
  addBlankSlide,
  addSlideTextBox,
  groupShapes,
  setShapeBounds,
  setShapeRotation,
  setShapeFlip,
  savePresentation,
  loadPresentation,
  getSlides,
  getSlideShapes,
  getShapeText,
  inches,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

for (const nested of [false, true]) {
  for (const flip of [
    { horizontal: false, vertical: false },
    { horizontal: true, vertical: false },
    { horizontal: false, vertical: true },
    { horizontal: true, vertical: true },
  ]) {
    test(
      `grouped text preserves preview orientation and scale in both languages (nested=${nested}, H=${flip.horizontal}, V=${flip.vertical})`,
      { timeout: 60000 },
      async () => {
        const dir = await mkdtemp(join(tmpdir(), 'office-group-text-'));
        let preview, browser, page;
        try {
          const pres = createPresentation(),
            slide = addBlankSlide(pres);
          const first = addSlideTextBox(slide, {
            x: inches(2),
            y: inches(2),
            w: inches(2),
            h: inches(1),
            text: '日本語 English',
          });
          const second = addSlideTextBox(slide, {
            x: inches(5),
            y: inches(2),
            w: inches(1),
            h: inches(1),
            text: 'Sibling',
          });
          setShapeRotation(first, 20);
          setShapeFlip(first, { vertical: true, horizontal: false });
          const group = groupShapes([first, second]);
          setShapeRotation(group, 30);
          setShapeFlip(group, flip);
          setShapeBounds(group, { x: inches(1), y: inches(2), w: inches(6), h: inches(0.5) });
          if (nested) {
            const outside = addSlideTextBox(slide, {
              x: inches(8),
              y: inches(2),
              w: inches(1),
              h: inches(1),
              text: 'Outer sibling',
            });
            const outer = groupShapes([group, outside]);
            setShapeRotation(outer, -25);
            setShapeFlip(outer, { horizontal: true, vertical: false });
            setShapeBounds(outer, { x: inches(1), y: inches(1), w: inches(8), h: inches(2) });
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
          page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
          const errors = [];
          page.on('pageerror', (e) => errors.push(e.message));
          await page.goto(preview.url);
          await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
          const editor = page.frameLocator('#editor-frame');
          await editor.getByText('Saved to this project', { exact: true }).waitFor();
          await editor.locator('.hit').dblclick();
          if (nested) await editor.locator('.hit.selected').dblclick();
          await editor.locator('.hit.selected').dblclick();
          const input = editor.locator('.inline-edit');
          await input.waitFor();
          const hitBounds = await editor.locator('.hit.selected').boundingBox();
          const inputBounds = await input.boundingBox();
          assert.ok(
            Math.abs(hitBounds.x + hitBounds.width / 2 - inputBounds.x - inputBounds.width / 2) < 1,
          );
          assert.ok(
            Math.abs(hitBounds.y + hitBounds.height / 2 - inputBounds.y - inputBounds.height / 2) <
              1,
          );
          const matrices = await input.evaluate((node) => {
            const foreign = [...document.querySelectorAll('.paint foreignObject')].find((n) =>
              n.textContent.includes('日本語'),
            );
            if (!foreign) throw new Error('Missing rendered text');
            const preview = foreign.getScreenCTM();
            const own = new DOMMatrix(getComputedStyle(node).transform);
            const parent = new DOMMatrix(getComputedStyle(node.parentElement).transform);
            const zoom = Number(getComputedStyle(node).getPropertyValue('--text-zoom'));
            const actual = parent.multiply(own).scale(zoom);
            return {
              actual: [actual.a, actual.b, actual.c, actual.d],
              expected: [preview.a, preview.b, preview.c, preview.d],
            };
          });
          matrices.actual.forEach((value, index) =>
            assert.ok(Math.abs(value - matrices.expected[index]) < 0.01, JSON.stringify(matrices)),
          );
          await input.fill('編集済み Edited');
          await page.keyboard.press('ControlOrMeta+Enter');
          await editor.getByText('Saved to this project', { exact: true }).waitFor();
          const loaded = await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          );
          assert.deepEqual(getSlideShapes(getSlides(loaded)[0]).map(getShapeText).filter(Boolean), [
            '編集済み Edited',
            'Sibling',
            ...(nested ? ['Outer sibling'] : []),
          ]);
          await editor.locator('.lang select').selectOption('ja');
          await page.reload();
          await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
          await editor.locator('.hit').dblclick();
          if (nested) await editor.locator('.hit.selected').dblclick();
          await editor.locator('.hit.selected').dblclick();
          await editor.getByRole('textbox', { name: 'テキストを編集', exact: true }).waitFor();
          await page.screenshot({ path: '/tmp/pptx-group-text-ja.png', fullPage: true });
          assert.deepEqual(errors, []);
        } catch (error) {
          await page?.screenshot({ path: '/tmp/pptx-group-text-failure.png', fullPage: true });
          throw error;
        } finally {
          await browser?.close();
          await preview?.close();
          await rm(dir, { recursive: true, force: true });
        }
      },
    );
  }
}

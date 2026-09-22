import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlides,
  getSlideShapes,
  getShapeBounds,
  getShapeRotation,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'resizing keeps aspect ratio with Shift and anchors rotated shapes in both languages',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-resize-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={3} height={1}>English</Text><Text x={7} y={3} width={3} height={1} rotation={45}>日本語</Text></Slide></Presentation>`,
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
      const shapes = async () =>
        getSlideShapes(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        );
      const center = async (locator) => {
        const b = await locator.boundingBox();
        return { x: b.x + b.width / 2, y: b.y + b.height / 2 };
      };
      const canvasCenter = async (locator) => {
        const position = await center(locator);
        const stage = await editor.locator('.stage').boundingBox();
        return { x: position.x - stage.x, y: position.y - stage.y };
      };
      const handle = (direction) =>
        editor.getByRole('button', {
          name: ja
            ? `サイズ変更（${{ nw: '左上', se: '右下' }[direction]}）`
            : `Resize ${direction}`,
          exact: true,
        });
      const drag = async (direction, dx, dy, shift = false) => {
        const p = await center(handle(direction));
        await page.mouse.move(p.x, p.y);
        if (shift) await page.keyboard.down('Shift');
        await page.mouse.down();
        await page.mouse.move(p.x + dx, p.y + dy, { steps: 8 });
        await page.mouse.up();
        if (shift) await page.keyboard.up('Shift');
        await saved();
      };
      await saved();
      const originals = (await shapes()).map(getShapeBounds);
      await editor.locator('.hit').nth(0).click();
      const nw = await canvasCenter(handle('nw'));
      await drag('se', 60, 45, true);
      const resized = getShapeBounds((await shapes())[0]);
      assert.ok(
        Math.abs(resized.w / resized.h - 3) < 0.00001,
        'Shift must preserve the original aspect ratio',
      );
      assert.ok(resized.w > originals[0].w);
      const anchored = await canvasCenter(handle('nw'));
      assert.ok(
        Math.hypot(nw.x - anchored.x, nw.y - anchored.y) < 1,
        'opposite corner stays fixed relative to the slide',
      );
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(getShapeBounds((await shapes())[0]), originals[0]);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual(getShapeBounds((await shapes())[0]), resized);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await editor.locator('.hit').nth(1).click();
      const fixed = await canvasCenter(handle('nw'));
      const start = await canvasCenter(handle('se'));
      await drag('se', 45, 15);
      const end = await canvasCenter(handle('se'));
      const stillFixed = await canvasCenter(handle('nw'));
      assert.ok(
        Math.hypot(fixed.x - stillFixed.x, fixed.y - stillFixed.y) < 1,
        'opposite rotated corner stays fixed',
      );
      assert.ok(
        Math.hypot(end.x - start.x - 45, end.y - start.y - 15) < 1,
        'rotated handle follows the pointer',
      );
      const result = (await shapes())[1];
      assert.equal(getShapeRotation(result), 45);
      const finalBounds = getShapeBounds(result);
      await page.reload();
      await saved();
      assert.deepEqual(getShapeBounds((await shapes())[1]), finalBounds);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

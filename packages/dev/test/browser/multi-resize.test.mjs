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
  'selection corner handles proportionally resize rotated objects and preserve history',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-multi-resize-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={2} y={2} width={2} height={1}>English</Text><Text x={6} y={2} width={2} height={1} rotation={45}>日本語</Text><Text x={10} y={5} width={1} height={1}>Other</Text></Slide></Presentation>`,
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
      const geometry = async () =>
        getSlideShapes(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        ).map((shape) => ({ ...getShapeBounds(shape), rotation: getShapeRotation(shape) }));
      const select = async () => {
        await editor.locator('.stage').click({ position: { x: 5, y: 5 } });
        await editor.locator('.hit').nth(0).click();
        await editor
          .locator('.hit')
          .nth(1)
          .click({ modifiers: ['Shift'] });
      };
      const handle = (direction) =>
        editor.getByRole('button', {
          name: ja
            ? `選択範囲を拡大・縮小（${{ nw: '左上', se: '右下' }[direction]}）`
            : `Scale selection ${direction}`,
          exact: true,
        });
      const centre = async (locator) => {
        const b = await locator.boundingBox();
        return { x: b.x + b.width / 2, y: b.y + b.height / 2 };
      };
      const relative = async (direction) => {
        const p = await centre(handle(direction));
        const stage = await editor.locator('.stage').boundingBox();
        return { x: p.x - stage.x, y: p.y - stage.y };
      };
      const beginResize = async (direction, dx, dy) => {
        const p = await centre(handle(direction));
        await page.mouse.move(p.x, p.y);
        await page.mouse.down();
        await page.mouse.move(p.x + dx, p.y + dy, { steps: 8 });
      };
      await saved();
      const original = await geometry();
      await select();
      assert.equal(
        await handle('se').count(),
        1,
        'multiple selection exposes corner resize handles',
      );
      assert.equal(await editor.locator('.multi-selection .handle').count(), 4);
      const anchor = await relative('nw');
      await beginResize('se', 80, 30);
      await page.mouse.up();
      await saved();
      const enlarged = await geometry();
      const scale = enlarged[0].w / original[0].w;
      assert.ok(scale > 1.1);
      const fixed = await relative('nw');
      assert.ok(Math.hypot(fixed.x - anchor.x, fixed.y - anchor.y) < 1);
      const pivot = { x: 2 * 914400, y: (2.5 - 1.5 / Math.SQRT2) * 914400 };
      for (let i = 0; i < 2; i++) {
        assert.ok(Math.abs(enlarged[i].h - original[i].h * scale) <= 2);
        assert.ok(Math.abs(enlarged[i].w - original[i].w * scale) <= 2);
        assert.ok(Math.abs(enlarged[i].x - (pivot.x + (original[i].x - pivot.x) * scale)) <= 2);
        assert.ok(Math.abs(enlarged[i].y - (pivot.y + (original[i].y - pivot.y) * scale)) <= 2);
        assert.equal(enlarged[i].rotation, original[i].rotation);
      }
      assert.deepEqual(enlarged[2], original[2]);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await geometry(), original);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await select();
      await beginResize('nw', 50, 25);
      await page.keyboard.press('Escape');
      await page.mouse.up();
      await saved();
      assert.deepEqual(await geometry(), original);
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual(await geometry(), enlarged);
      await select();
      const se = await relative('se');
      await beginResize('nw', 40, 15);
      await page.mouse.up();
      await saved();
      const shrunk = await geometry();
      assert.ok(shrunk[0].w < enlarged[0].w);
      const stillFixed = await relative('se');
      assert.ok(Math.hypot(se.x - stillFixed.x, se.y - stillFixed.y) < 1);
      await page.reload();
      await saved();
      assert.deepEqual(await geometry(), shrunk);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  getSlides,
  getSlideShapes,
  getShapeBounds,
  getShapeRotation,
  inches,
  isShapeAspectRatioLocked,
  loadPresentation,
  savePresentation,
  setShapeAspectRatioLocked,
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
      const firstBox = await editor.locator('.hit').nth(0).boundingBox();
      assert.ok(firstBox);
      await page.mouse.click(firstBox.x + 2, firstBox.y + 2);
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
      await editor.locator('.canvas-shell .inline-edit').press('Escape');
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

test(
  'canvas resize follows the saved aspect-ratio lock for corners while side handles stay free',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-resize-aspect-lock-'));
    let preview, browser;
    try {
      const deck = createPresentation();
      const slide = addBlankSlide(deck);
      const shape = addSlideTextBox(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(2),
        text: 'Aspect lock',
      });
      setShapeAspectRatioLocked(shape, true);
      const source = join(dir, 'source.pptx');
      await writeFile(source, await savePresentation(deck));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const read = async () => {
        const savedDeck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getShapeBounds(getSlideShapes(getSlides(savedDeck)[0])[0]);
      };
      const readLock = async () => {
        const savedDeck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return isShapeAspectRatioLocked(getSlideShapes(getSlides(savedDeck)[0])[0]);
      };
      const select = async () => {
        await editor
          .locator('.hit')
          .first()
          .click({ button: 'right', position: { x: 2, y: 2 } });
        await editor.getByRole('menuitem', { name: 'Size and Position...', exact: true }).click();
        await editor.getByRole('tab', { name: 'Size & Properties', exact: true }).click();
      };
      const center = async (locator) => {
        const box = await locator.boundingBox();
        return { x: box.x + box.width / 2, y: box.y + box.height / 2 };
      };
      const resize = async (direction, dx, dy) => {
        const point = await center(
          editor.getByRole('button', { name: `Resize ${direction}`, exact: true }),
        );
        await page.mouse.move(point.x, point.y);
        await page.mouse.down();
        await page.mouse.move(point.x + dx, point.y + dy, { steps: 8 });
        await page.mouse.up();
        await saved();
      };

      await saved();
      await select();
      const lock = editor.getByRole('checkbox', { name: 'Lock aspect ratio', exact: true });
      assert.equal(await lock.isChecked(), true);
      assert.equal(await readLock(), true);
      const original = await read();

      await resize('se', 60, 10);
      const cornerLocked = await read();
      assert.ok(
        Math.abs(cornerLocked.w / cornerLocked.h - original.w / original.h) < 0.00001,
        'saved aspect lock must preserve the original ratio during corner resize',
      );
      assert.ok(cornerLocked.w > original.w && cornerLocked.h > original.h);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), original);
      assert.equal(await readLock(), true);

      await select();
      await resize('e', 60, 25);
      const sideLocked = await read();
      assert.ok(sideLocked.w > original.w);
      assert.equal(sideLocked.x, original.x);
      assert.equal(sideLocked.y, original.y);
      assert.equal(sideLocked.h, original.h);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), original);
      assert.equal(await readLock(), true);

      await select();
      await lock.uncheck();
      await saved();
      assert.equal(await readLock(), false);
      await select();
      await resize('se', 60, 10);
      const cornerFree = await read();
      assert.ok(cornerFree.w > original.w && cornerFree.h > original.h);
      assert.notEqual(cornerFree.w / cornerFree.h, original.w / original.h);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), original);
      assert.equal(await readLock(), false);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

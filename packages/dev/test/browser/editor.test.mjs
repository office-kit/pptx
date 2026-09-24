import { chromium } from 'playwright';
import {
  loadPresentation,
  getSlides,
  getSlideText,
  getSlideShapes,
  getShapeParagraphElements,
} from '@office-kit/pptx';
import { buildDeck } from '../../dist/index.mjs';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test(
  'canvas editing, history, notes and reload persist without changing TSX',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-editor-'));
    const file = join(dir, 'deck.tsx');
    const source = `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={8} height={1}>Original title</Text></Slide><Slide><Text x={1} y={1} width={8} height={1}>Second slide</Text></Slide></Presentation>`;
    await writeFile(file, source);
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
      page.on('pageerror', (e) => errors.push(e.message));
      await page.goto(url);
      const state = () => page.evaluate(() => window.office.getState());
      const settled = async (predicate) => {
        try {
          await page.waitForFunction(predicate);
        } catch (e) {
          console.error(await state(), await page.locator('#edit-message').textContent());
          throw e;
        }
        await page.waitForFunction(
          () =>
            !window.office.getState().building &&
            !document.querySelector('[data-edit="undo"]').disabled,
        );
      };
      await page.locator('.shape-hit').first().waitFor();
      const guideBefore = await state();
      const guideShape = guideBefore.editor.slides[0].shapes[0];
      const guideCanvas = await page.locator('#slide').boundingBox();
      const guideHit = await page.locator('.shape-hit').first().boundingBox();
      const guideScale = guideCanvas.width / guideBefore.editor.width;
      const guideX = guideCanvas.x + guideCanvas.width / 2 + 2;
      const guideY = guideHit.y + guideHit.height / 2;
      await page.mouse.move(guideHit.x + guideHit.width / 2, guideY);
      await page.mouse.down();
      await page.mouse.move(guideX, guideY, { steps: 5 });
      assert.equal(await page.locator('.smart-guide[data-axis="x"]').count(), 1);
      const snappedLeft = await page
        .locator('.shape-hit.selected')
        .evaluate((node) => parseFloat(node.style.left));
      assert.ok(
        Math.abs(
          snappedLeft - ((guideBefore.editor.width - guideShape.bounds.w) / 2) * guideScale,
        ) < 0.1,
      );
      await page.keyboard.down('Meta');
      await page.mouse.move(guideX + 1, guideY);
      assert.equal(await page.locator('.smart-guide').count(), 0);
      const freeLeft = await page
        .locator('.shape-hit.selected')
        .evaluate((node) => parseFloat(node.style.left));
      assert.ok(Math.abs(freeLeft - snappedLeft - 3) < 0.1);
      await page.keyboard.up('Meta');
      await page.keyboard.press('Escape');
      await page.mouse.up();
      assert.equal(await page.locator('.smart-guide').count(), 0);
      assert.equal((await state()).revision, guideBefore.revision);
      assert.deepEqual((await state()).editor, guideBefore.editor);
      await page.locator('#tab-view').click();
      await page.locator('#view-grid-guides').click();
      const smartOption = page.getByRole('menuitemcheckbox', { name: 'Smart Guides', exact: true });
      assert.equal(await smartOption.getAttribute('aria-checked'), 'true');
      await smartOption.click();
      await page.mouse.move(guideHit.x + guideHit.width / 2, guideY);
      await page.mouse.down();
      await page.mouse.move(guideX, guideY, { steps: 5 });
      assert.equal(await page.locator('.smart-guide').count(), 0);
      const unsnappedLeft = await page
        .locator('.shape-hit.selected')
        .evaluate((node) => parseFloat(node.style.left));
      assert.ok(Math.abs(unsnappedLeft - snappedLeft - 2) < 0.1);
      await page.keyboard.press('Escape');
      await page.mouse.up();
      await page.locator('#show-guides').check();
      await page.locator('.drawing-guide').nth(1).waitFor();
      assert.equal(await page.locator('.drawing-guide').count(), 2);
      const verticalGuide = page.locator('.drawing-guide[data-axis="x"]');
      const guideStart = await verticalGuide.boundingBox();
      const guideDragY = guideStart.y + guideStart.height * 0.7;
      await page.mouse.move(guideStart.x + 3, guideDragY);
      await page.mouse.down();
      await page.mouse.move(guideStart.x + 43, guideDragY);
      assert.match(await page.locator('.guide-distance').textContent(), /^\d+\.\d{2} cm$/);
      assert.ok(Math.abs((await verticalGuide.boundingBox()).x - guideStart.x - 40) < 0.1);
      await page.mouse.up();
      await settled(() => window.office.getState().editor.guides[0].offset !== 0);
      assert.equal(await page.locator('.guide-distance').count(), 0);
      await page.mouse.down();
      await page.mouse.move(guideStart.x + 63, guideDragY);
      await page.keyboard.press('Escape');
      await page.mouse.up();
      assert.ok(Math.abs((await verticalGuide.boundingBox()).x - guideStart.x - 40) < 0.1);
      await verticalGuide.click({
        button: 'right',
        position: { x: 3, y: guideStart.height * 0.7 },
      });
      await page.getByRole('menuitem', { name: 'Color', exact: true }).click();
      await page.getByRole('menuitemcheckbox', { name: 'Blue', exact: true }).click();
      await settled(() => window.office.getState().editor.guides[0].color === '#2873c4');
      assert.equal(
        await verticalGuide.locator('div').evaluate((node) => node.style.borderLeftColor),
        'rgb(40, 115, 196)',
      );
      await page.locator('#view-grid-guides').click();
      await page.getByRole('menuitem', { name: 'Add Vertical Guide', exact: true }).click();
      await settled(() => window.office.getState().editor.guides.length === 3);
      assert.equal(await page.locator('.drawing-guide').count(), 3);
      const addedGuide = page.locator('.drawing-guide[data-guide-id="3"]');
      await addedGuide.click({ button: 'right', position: { x: 3, y: guideStart.height * 0.7 } });
      await page.getByRole('menuitem', { name: 'Delete', exact: true }).click();
      await settled(() => window.office.getState().editor.guides.length === 2);
      assert.equal(await page.locator('.drawing-guide').count(), 2);
      await page.reload();
      await page.locator('.shape-hit').first().waitFor();
      assert.equal(await page.locator('.drawing-guide').count(), 2);
      assert.ok(Math.abs((await verticalGuide.boundingBox()).x - guideStart.x - 40) < 0.1);
      assert.equal(
        await verticalGuide.locator('div').evaluate((node) => node.style.borderLeftColor),
        'rgb(40, 115, 196)',
      );
      assert.ok((await state()).revision > guideBefore.revision);
      assert.deepEqual((await state()).editor.slides, guideBefore.editor.slides);
      await page.locator('[data-edit="undo"]').click();
      await settled(() => window.office.getState().editor.guides.length === 3);
      assert.equal(await page.locator('.drawing-guide').count(), 3);
      await page.locator('[data-edit="redo"]').click();
      await settled(() => window.office.getState().editor.guides.length === 2);
      await page.locator('#tab-view').click();
      await page.locator('#view-grid-guides').click();
      assert.equal(await smartOption.getAttribute('aria-checked'), 'false');
      assert.equal(
        await page
          .getByRole('menuitemcheckbox', { name: 'Guides', exact: true })
          .getAttribute('aria-checked'),
        'true',
      );
      await smartOption.click();
      await page.locator('#show-guides').uncheck();
      await page.locator('.drawing-guide').waitFor({ state: 'detached' });
      assert.equal(await page.locator('.drawing-guide').count(), 0);
      await page.locator('#tab-home').click();
      await page.locator('.shape-hit').first().click();
      await page.keyboard.press('F2');
      const frameLayout = await page
        .getByRole('textbox', { name: 'Edit text', exact: true })
        .evaluate((field) => {
          const model = window.office.getState().editor;
          const shape = model.slides[0].shapes[0];
          const style = getComputedStyle(field);
          const scale = field.getBoundingClientRect().width / shape.bounds.w;
          return {
            left: parseFloat(style.paddingLeft),
            expectedLeft: shape.textFrame.margins.left * scale,
            top: parseFloat(style.paddingTop),
            expectedTop: shape.textFrame.margins.top * scale,
            anchor: style.alignContent,
            wrap: style.whiteSpace,
          };
        });
      assert.ok(Math.abs(frameLayout.left - frameLayout.expectedLeft) < 0.05);
      assert.ok(Math.abs(frameLayout.top - frameLayout.expectedTop) < 0.05);
      assert.equal(frameLayout.anchor, 'start');
      assert.equal(frameLayout.wrap, 'pre-wrap');
      await page.getByRole('textbox', { name: 'Edit text', exact: true }).fill('編集済み title');
      await page.keyboard.press('Escape');
      await settled(
        () => window.office.getState().editor.slides[0].shapes[0].text === '編集済み title',
      );
      await page.locator('[data-edit="bold"]').click();
      await settled(() => window.office.getState().editor.slides[0].shapes[0].format.bold === true);
      await page.locator('.shape-hit').first().click();
      const before = (await state()).editor.slides[0].shapes[0].bounds.x;
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(
        (x) => window.office.getState().editor.slides[0].shapes[0].bounds.x === x + 12700,
        before,
      );
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          !document.querySelector('[data-edit="undo"]').disabled,
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        (x) => window.office.getState().editor.slides[0].shapes[0].bounds.x === x,
        before,
      );
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          !document.querySelector('[data-edit="redo"]').disabled,
      );
      await page.locator('[data-edit="redo"]').click();
      await page.waitForFunction(
        (x) => window.office.getState().editor.slides[0].shapes[0].bounds.x === x + 12700,
        before,
      );
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          !document.querySelector('[data-edit="undo"]').disabled,
      );
      await page.locator('[data-edit="notes"]').click();
      await page.locator('#speaker-notes').fill('Speaker note');
      await page.locator('#speaker-notes').press('Tab');
      await settled(() => window.office.getState().editor.slides[0].notes === 'Speaker note');
      await page.getByRole('button', { name: 'Slide 2', exact: true }).click();
      assert.equal(await page.locator('#speaker-notes').inputValue(), '');
      await page.getByRole('button', { name: 'Slide 1', exact: true }).click();
      assert.equal(await page.locator('#speaker-notes').inputValue(), 'Speaker note');
      await page.reload();
      await page.waitForFunction(
        () => window.office?.getState().editor?.slides[0].shapes[0].text === '編集済み title',
      );
      assert.equal((await state()).editor.slides[0].shapes[0].format.bold, true);
      assert.equal(await readFile(file, 'utf8'), source);
      // Display-only guide toggles do not create document edits.
      assert.equal(JSON.parse(await readFile(file + '.edits.json', 'utf8')).cursor, 8);
      await page.locator('[data-edit="shapes"]').first().click();
      const gallery = page.getByRole('menu', { name: 'Shapes', exact: true });
      assert.equal(await gallery.locator('.gallery-heading').count(), 9);
      assert.ok((await gallery.getByRole('menuitem').count()) > 130);
      await page.keyboard.press('ArrowRight');
      assert.equal(
        await page.evaluate(() => document.activeElement.getAttribute('aria-label')),
        'Arrow',
      );
      await page.keyboard.press('ArrowDown');
      assert.equal(
        await page.evaluate(() => document.activeElement.getAttribute('aria-label')),
        'Rounded Rectangle',
      );
      await page.keyboard.press('End');
      assert.equal(
        await page.evaluate(() => document.activeElement.getAttribute('aria-label')),
        'Action Button: Custom',
      );
      await page.keyboard.press('Escape');
      assert.equal(await gallery.count(), 0);
      await page.locator('[data-edit="shapes"]').first().click();
      await page.getByRole('menuitem', { name: 'Arrow', exact: true }).click();
      const canvas = await page.locator('#slide').boundingBox();
      await page.mouse.move(canvas.x + 280, canvas.y + 270);
      await page.mouse.down();
      await page.keyboard.down('Shift');
      await page.mouse.move(canvas.x + 90, canvas.y + 276, { steps: 5 });
      const linePreview = await page.locator('.draw-box svg').boundingBox();
      assert.ok(linePreview.width > 0 && linePreview.height > 0);
      await page.mouse.up();
      await page.keyboard.up('Shift');
      await settled(() => window.office.getState().editor.slides[0].shapes.length === 2);
      const insertedLine = (await state()).editor.slides[0].shapes[1];
      assert.equal(insertedLine.kind, 'connector');
      assert.equal(insertedLine.bounds.h, 0);
      assert.equal(await page.locator('.shape-hit.selected .shape-handle').count(), 2);
      const endpoint = await page.getByLabel('End point', { exact: true }).boundingBox();
      await page.mouse.move(endpoint.x + endpoint.width / 2, endpoint.y + endpoint.height / 2);
      await page.mouse.down();
      await page.mouse.move(canvas.x + 350, canvas.y + 320, { steps: 5 });
      assert.equal(await page.locator('.endpoint-preview').count(), 1);
      await page.mouse.up();
      await settled(() => window.office.getState().editor.slides[0].shapes[1].bounds.h > 0);
      const editedLine = (await state()).editor.slides[0].shapes[1];
      assert.equal(editedLine.flip.horizontal, false);
      assert.ok(
        Math.abs(editedLine.bounds.x - (insertedLine.bounds.x + insertedLine.bounds.w)) < 2,
      );

      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(
        (x) => window.office.getState().editor.slides[0].shapes[1].bounds.x > x,
        editedLine.bounds.x,
      );
      await settled(() => !window.office.getState().building);
      await page.evaluate(async () => {
        const state = window.office.getState();
        const response = await fetch('/edit', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            revision: state.revision,
            command: {
              type: 'update',
              slide: 0,
              ids: [state.editor.slides[0].shapes[1].id],
              changes: { rotation: 37 },
            },
          }),
        });
        if (!response.ok) throw new Error(await response.text());
      });
      await settled(() => window.office.getState().editor.slides[0].shapes[1].rotation === 37);
      const endpointRevision = (await state()).revision;
      await page.getByLabel('Start point', { exact: true }).waitFor({ state: 'visible' });
      await page.getByLabel('End point', { exact: true }).waitFor({ state: 'visible' });
      const fixedBefore = await page.getByLabel('Start point', { exact: true }).boundingBox();
      const rotatedEnd = await page.getByLabel('End point', { exact: true }).boundingBox();
      await page.mouse.move(
        rotatedEnd.x + rotatedEnd.width / 2,
        rotatedEnd.y + rotatedEnd.height / 2,
      );
      await page.mouse.down();
      await page.mouse.move(
        rotatedEnd.x + rotatedEnd.width / 2 + 60,
        rotatedEnd.y + rotatedEnd.height / 2 - 40,
        { steps: 5 },
      );
      await page.mouse.up();
      await settled(() => !window.office.getState().building);
      await page.waitForFunction(
        (revision) => window.office.getState().revision > revision,
        endpointRevision,
      );
      const fixedAfter = await page.getByLabel('Start point', { exact: true }).boundingBox();
      assert.ok(Math.abs(fixedBefore.x - fixedAfter.x) < 1);
      assert.ok(Math.abs(fixedBefore.y - fixedAfter.y) < 1);
      await page.keyboard.press('Delete');
      await settled(() => window.office.getState().editor.slides[0].shapes.length === 1);
      const connectionTarget = (await state()).editor.slides[0].shapes[0];
      const site = connectionTarget.connectionSites[3];
      const scale = canvas.width / (await state()).editor.width;
      const siteX = canvas.x + site.x * scale,
        siteY = canvas.y + site.y * scale;
      await page.locator('[data-edit="shapes"]').first().click();
      await page.getByRole('menuitem', { name: 'Arrow', exact: true }).click();
      await page.mouse.move(siteX, siteY);
      assert.ok((await page.locator('.connection-site').count()) >= 4);
      await page.mouse.down();
      await page.mouse.move(siteX + 100, siteY + 100, { steps: 5 });
      await page.mouse.up();
      await settled(() => window.office.getState().editor.slides[0].shapes.length === 2);
      const attachedLine = (await state()).editor.slides[0].shapes[1];
      assert.deepEqual(attachedLine.connections.start, {
        shapeId: connectionTarget.id,
        siteIndex: 3,
      });
      await page.locator('.shape-hit').first().click();
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(
        (x) => window.office.getState().editor.slides[0].shapes[0].bounds.x > x,
        connectionTarget.bounds.x,
      );
      await settled(() => !window.office.getState().building);
      const followedLine = (await state()).editor.slides[0].shapes[1];
      assert.ok(Math.abs(followedLine.bounds.x - attachedLine.bounds.x - 12700) < 2);
      // This drag is a separate gesture, outside the text double-click interval.
      await page.waitForTimeout(450);
      const targetHit = await page
        .locator(`[data-shape-id="${connectionTarget.id}"]`)
        .boundingBox();
      await page.mouse.move(targetHit.x + targetHit.width / 2, targetHit.y + targetHit.height / 2);
      await page.mouse.down();
      await page.mouse.move(
        targetHit.x + targetHit.width / 2 + 30,
        targetHit.y + targetHit.height / 2 + 20,
        { steps: 5 },
      );
      const followingPreview = page.locator('.attached-preview');
      assert.equal(await followingPreview.count(), 1);
      const previewLeft = await followingPreview.evaluate((node) => parseFloat(node.style.left));
      assert.ok(Math.abs(previewLeft - followedLine.bounds.x * scale - 30) < 1);
      const dragRevision = (await state()).revision;
      await page.mouse.up();
      await page.waitForFunction(
        (revision) => window.office.getState().revision > revision,
        dragRevision,
      );
      await settled(() => !window.office.getState().building);
      await page.reload();
      await page.waitForFunction(
        () => window.office?.getState().editor?.slides[0].shapes.length === 2,
      );
      assert.deepEqual(
        (await state()).editor.slides[0].shapes[1].connections.start,
        attachedLine.connections.start,
      );
      await page.locator(`[data-shape-id="${attachedLine.id}"]`).click();
      const attachedHandle = await page.getByLabel('Start point', { exact: true }).boundingBox();
      await page.mouse.move(attachedHandle.x + 4, attachedHandle.y + 4);
      await page.mouse.down();
      await page.mouse.move(siteX + 30, siteY + 50, { steps: 5 });
      await page.mouse.up();
      await settled(
        () => window.office.getState().editor.slides[0].shapes[1].connections.start === null,
      );
      await page.keyboard.press('Delete');
      await settled(() => window.office.getState().editor.slides[0].shapes.length === 1);
      await page.locator('[data-edit="shapes"]').first().click();
      await page.getByRole('menuitem', { name: 'Rectangle', exact: true }).click();
      await page.mouse.move(canvas.x + 50, canvas.y + 150);
      await page.mouse.down();
      await page.mouse.move(canvas.x + 180, canvas.y + 220, { steps: 5 });
      await page.mouse.up();
      await settled(() => window.office.getState().editor.slides[0].shapes.length === 2);
      const stackingOrder = (await state()).editor.slides[0].shapes.map((shape) => shape.id);
      await page.keyboard.press('Meta+Shift+B');
      await page.waitForFunction(
        (id) => window.office.getState().editor.slides[0].shapes[0].id === id,
        stackingOrder[1],
      );
      await settled(() => !window.office.getState().building);
      await page.keyboard.press('Meta+Alt+Shift+F');
      await page.waitForFunction(
        (id) => window.office.getState().editor.slides[0].shapes[1].id === id,
        stackingOrder[1],
      );
      await settled(() => !window.office.getState().building);
      await page.keyboard.press('Meta+Alt+Shift+B');
      await page.waitForFunction(
        (id) => window.office.getState().editor.slides[0].shapes[0].id === id,
        stackingOrder[1],
      );
      await settled(() => !window.office.getState().building);
      await page.keyboard.press('Meta+Shift+F');
      await page.waitForFunction(
        (id) => window.office.getState().editor.slides[0].shapes[1].id === id,
        stackingOrder[1],
      );
      await settled(() => !window.office.getState().building);
      assert.deepEqual(
        (await state()).editor.slides[0].shapes.map((shape) => shape.id),
        stackingOrder,
      );
      const boundsBeforeRotation = (await state()).editor.slides[0].shapes[1].bounds;
      await page.keyboard.press('Alt+ArrowRight');
      await settled(() => window.office.getState().editor.slides[0].shapes[1].rotation === 15);
      assert.deepEqual((await state()).editor.slides[0].shapes[1].bounds, boundsBeforeRotation);
      await page.keyboard.press('Alt+ArrowLeft');
      await settled(() => window.office.getState().editor.slides[0].shapes[1].rotation === 0);
      const originalSize = (await state()).editor.slides[0].shapes[1].bounds;
      const historyBeforeModifiers = (await state()).history.undo;
      for (const keys of [
        'Meta+Shift+ArrowRight',
        'Control+Shift+ArrowRight',
        'Alt+Shift+ArrowRight',
      ]) {
        await page.keyboard.press(keys);
      }
      await page.waitForTimeout(150);
      assert.deepEqual((await state()).editor.slides[0].shapes[1].bounds, originalSize);
      assert.equal((await state()).history.undo, historyBeforeModifiers);
      for (const [key, axis, factor] of [
        ['ArrowRight', 'w', 1.1],
        ['ArrowLeft', 'w', 1 / 1.1],
        ['ArrowUp', 'h', 1.1],
        ['ArrowDown', 'h', 1 / 1.1],
      ]) {
        await page.keyboard.press('Shift+' + key);
        await page.waitForFunction(
          ({ axis, expected }) =>
            window.office.getState().editor.slides[0].shapes[1].bounds[axis] === expected,
          { axis, expected: Math.round(originalSize[axis] * factor) },
        );
        await settled(() => !window.office.getState().building);
        const resized = (await state()).editor.slides[0].shapes[1].bounds;
        assert.ok(Math.abs(resized.x + resized.w / 2 - originalSize.x - originalSize.w / 2) <= 1);
        assert.ok(Math.abs(resized.y + resized.h / 2 - originalSize.y - originalSize.h / 2) <= 1);
        assert.equal(resized[axis === 'w' ? 'h' : 'w'], originalSize[axis === 'w' ? 'h' : 'w']);
        await page.keyboard.press('Meta+z');
        await page.waitForFunction(
          (box) =>
            JSON.stringify(window.office.getState().editor.slides[0].shapes[1].bounds) ===
            JSON.stringify(box),
          originalSize,
        );
        await settled(() => !window.office.getState().building);
      }
      const shapeBefore = (await state()).editor.slides[0].shapes[1];
      const widthBeforeProperties = await page
        .locator('#slide')
        .evaluate((element) => element.clientWidth);
      await page.locator('.shape-hit.selected').click({ button: 'right' });
      await page.getByRole('menuitem', { name: 'Size and Position...', exact: true }).click();
      const properties = page.getByRole('complementary', { name: 'Format Shape', exact: true });
      await properties.getByRole('checkbox', { name: 'Lock aspect ratio', exact: true }).check();
      await settled(() => window.office.getState().editor.slides[0].shapes[1].aspectRatioLocked);
      // A mostly vertical corner drag must resize both dimensions when aspect is locked.
      const lockedBeforeDrag = (await state()).editor.slides[0].shapes[1];
      const beforeDragRevision = (await state()).revision;
      const lockedCorner = await page
        .locator('.shape-hit.selected [data-handle="se"]')
        .boundingBox();
      assert.ok(lockedCorner);
      await page.mouse.move(
        lockedCorner.x + lockedCorner.width / 2,
        lockedCorner.y + lockedCorner.height / 2,
      );
      await page.mouse.down();
      await page.mouse.move(
        lockedCorner.x + lockedCorner.width / 2,
        lockedCorner.y + lockedCorner.height / 2 + 24,
        { steps: 5 },
      );
      await page.waitForFunction(() =>
        document.querySelector('#slide').shadowRoot.querySelector('.drag-preview svg'),
      );
      assert.deepEqual((await state()).editor.slides[0].shapes[1], lockedBeforeDrag);
      assert.equal((await state()).revision, beforeDragRevision);
      const transientHeight = await page
        .locator('.drag-preview [data-pptx-shape-id="' + lockedBeforeDrag.id + '"]')
        .evaluate((node) => node.getBBox().height);
      assert.ok(transientHeight > 0);
      await page.mouse.up();
      await page.waitForFunction(
        (oldHeight) =>
          window.office.getState().editor.slides[0].shapes[1].bounds.h !== oldHeight &&
          !document.querySelector('[data-edit="undo"]').disabled,
        lockedBeforeDrag.bounds.h,
      );
      assert.equal(await page.locator('.drag-preview').count(), 0);
      const lockedAfterDrag = (await state()).editor.slides[0].shapes[1];
      const savedPreviewHeight = await page
        .locator('[data-pptx-shape-id="' + lockedBeforeDrag.id + '"]')
        .evaluate((node) => node.getBBox().height);
      assert.ok(Math.abs(savedPreviewHeight - transientHeight) < 0.1);
      assert.ok(lockedAfterDrag.bounds.w > lockedBeforeDrag.bounds.w);
      assert.ok(
        Math.abs(
          lockedAfterDrag.bounds.w / lockedAfterDrag.bounds.h -
            lockedBeforeDrag.bounds.w / lockedBeforeDrag.bounds.h,
        ) < 0.00001,
      );
      const oppositeCorner = (shape) => {
        const angle = (shape.rotation * Math.PI) / 180;
        const { x, y, w, h } = shape.bounds;
        return {
          x: x + w / 2 - (w / 2) * Math.cos(angle) + (h / 2) * Math.sin(angle),
          y: y + h / 2 - (w / 2) * Math.sin(angle) - (h / 2) * Math.cos(angle),
        };
      };
      const lockedFixedBefore = oppositeCorner(lockedBeforeDrag);
      const lockedFixedAfter = oppositeCorner(lockedAfterDrag);
      assert.ok(Math.abs(lockedFixedBefore.x - lockedFixedAfter.x) < 2);
      assert.ok(Math.abs(lockedFixedBefore.y - lockedFixedAfter.y) < 2);
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        (oldWidth) =>
          window.office.getState().editor.slides[0].shapes[1].bounds.w === oldWidth &&
          !document.querySelector('[data-edit="undo"]').disabled,
        lockedBeforeDrag.bounds.w,
      );
      assert.deepEqual((await state()).editor.slides[0].shapes[1], lockedBeforeDrag);
      const cancelRevision = (await state()).revision;
      const cancelCorner = await page
        .locator('.shape-hit.selected [data-handle="se"]')
        .boundingBox();
      await page.mouse.move(
        cancelCorner.x + cancelCorner.width / 2,
        cancelCorner.y + cancelCorner.height / 2,
      );
      await page.mouse.down();
      await page.mouse.move(
        cancelCorner.x + cancelCorner.width / 2 + 20,
        cancelCorner.y + cancelCorner.height / 2 + 20,
      );
      await page.waitForFunction(() =>
        document.querySelector('#slide').shadowRoot.querySelector('.drag-preview svg'),
      );
      await page.keyboard.press('Escape');
      await page.mouse.up();
      assert.equal(await page.locator('.drag-preview').count(), 0);
      assert.equal((await state()).revision, cancelRevision);
      assert.deepEqual((await state()).editor.slides[0].shapes[1], lockedBeforeDrag);
      assert.equal(
        await page
          .locator('[data-pptx-shape-id="' + lockedBeforeDrag.id + '"]')
          .evaluate((node) => getComputedStyle(node).visibility),
        'visible',
      );
      await properties.getByRole('spinbutton', { name: 'Width', exact: true }).fill('12');
      await properties.getByRole('spinbutton', { name: 'Width', exact: true }).press('Enter');
      await settled(() => window.office.getState().editor.slides[0].shapes[1].bounds.w === 4320000);
      const numericBounds = (await state()).editor.slides[0].shapes[1].bounds;
      if (process.env.PPTX_EDITOR_SCREENSHOT)
        await page.screenshot({ path: process.env.PPTX_EDITOR_SCREENSHOT });
      assert.equal(numericBounds.x, shapeBefore.bounds.x);
      assert.equal(numericBounds.y, shapeBefore.bounds.y);
      assert.equal(
        numericBounds.h,
        Math.round((shapeBefore.bounds.h * 4320000) / shapeBefore.bounds.w),
      );
      await page.locator('[data-edit="undo"]').click();
      await settled(() => window.office.getState().editor.slides[0].shapes[1].bounds.w !== 4320000);
      assert.deepEqual((await state()).editor.slides[0].shapes[1].bounds, shapeBefore.bounds);
      await page.locator('[data-edit="undo"]').click();
      await settled(() => !window.office.getState().editor.slides[0].shapes[1].aspectRatioLocked);
      assert.equal(
        await properties
          .getByRole('checkbox', { name: 'Lock aspect ratio', exact: true })
          .isChecked(),
        false,
      );
      await properties
        .locator('summary')
        .filter({ hasText: /^Text Box$/ })
        .click();
      const beforeCenteredAnchor = (await state()).editor.slides[0].shapes[1];
      await properties
        .getByRole('combobox', { name: 'Vertical alignment', exact: true })
        .selectOption('bottom-centered');
      await settled(() => window.office.getState().editor.slides[0].shapes[1].anchorCenter);
      assert.equal((await state()).editor.slides[0].shapes[1].textFrame.anchor, 'bottom');
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction((before) => {
        const shape = window.office.getState().editor.slides[0].shapes[1];
        return (
          shape.anchorCenter === before.anchorCenter &&
          shape.textFrame.anchor === before.textFrame.anchor &&
          !document.querySelector('[data-edit="undo"]').disabled
        );
      }, beforeCenteredAnchor);
      await properties.getByRole('spinbutton', { name: 'Left margin', exact: true }).fill('0.5');
      await properties.getByRole('spinbutton', { name: 'Left margin', exact: true }).press('Enter');
      await settled(
        () => window.office.getState().editor.slides[0].shapes[1].textFrame.margins.left === 180000,
      );
      await properties.getByRole('checkbox', { name: 'Wrap text in shape', exact: true }).uncheck();
      await settled(() => !window.office.getState().editor.slides[0].shapes[1].textFrame.wrap);
      await page.locator('[data-edit="undo"]').click();
      await settled(() => window.office.getState().editor.slides[0].shapes[1].textFrame.wrap);
      await page.locator('[data-edit="undo"]').click();
      await settled(
        () => window.office.getState().editor.slides[0].shapes[1].textFrame.margins.left !== 180000,
      );
      assert.deepEqual((await state()).editor.slides[0].shapes[1].textFrame, shapeBefore.textFrame);
      await properties.getByRole('button', { name: 'Columns...', exact: true }).click();
      const columnsDialog = page.getByRole('dialog', { name: 'Columns', exact: true });
      await columnsDialog.getByRole('spinbutton', { name: 'Number of columns:' }).fill('2');
      await columnsDialog.getByRole('spinbutton', { name: 'Spacing between columns:' }).fill('0.5');
      if (process.env.PPTX_COLUMNS_SCREENSHOT)
        await page.screenshot({ path: process.env.PPTX_COLUMNS_SCREENSHOT });
      await columnsDialog.getByRole('button', { name: 'OK', exact: true }).click();
      await settled(
        () =>
          window.office.getState().editor.slides[0].shapes[1].textFrame.columns?.gapEmu === 180000,
      );
      assert.deepEqual((await state()).editor.slides[0].shapes[1].textFrame.columns, {
        count: 2,
        gapEmu: 180000,
      });
      await properties.getByRole('button', { name: 'Columns...', exact: true }).click();
      await columnsDialog.getByRole('spinbutton', { name: 'Number of columns:' }).fill('3');
      await columnsDialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.equal((await state()).editor.slides[0].shapes[1].textFrame.columns.count, 2);
      await page.locator('[data-edit="undo"]').click();
      await settled(
        () => window.office.getState().editor.slides[0].shapes[1].textFrame.columns === null,
      );
      if (process.env.PPTX_AUTOFIT_SCREENSHOT)
        await page.screenshot({ path: process.env.PPTX_AUTOFIT_SCREENSHOT });
      const previousAutofit = (await state()).editor.slides[0].shapes[1].autoFit;
      await properties.getByRole('radio', { name: 'Shrink text on overflow', exact: true }).check();
      await settled(() => window.office.getState().editor.slides[0].shapes[1].autoFit === 'normal');
      assert.equal(
        await properties
          .getByRole('radio', { name: 'Shrink text on overflow', exact: true })
          .isChecked(),
        true,
      );
      const beforeAutofitText = (await state()).editor.slides[0].shapes[1].text;
      await page.locator('.shape-hit').nth(1).dblclick();
      const autofitField = page.getByRole('textbox', { name: 'Edit text', exact: true });
      const longText = 'Shrink this text to fit. '.repeat(60);
      await autofitField.fill(longText);
      const shrunkenSize = await autofitField
        .locator('span')
        .first()
        .evaluate((s) => parseFloat(getComputedStyle(s).fontSize));
      await page.keyboard.press('Meta+z');
      assert.equal(await autofitField.evaluate((f) => f.value), beforeAutofitText);
      await page.keyboard.press('Meta+Shift+z');
      assert.equal(await autofitField.evaluate((f) => f.value), longText);
      assert.equal(
        await autofitField
          .locator('span')
          .first()
          .evaluate((s) => parseFloat(getComputedStyle(s).fontSize)),
        shrunkenSize,
      );
      await page.keyboard.press('Escape');
      await settled(() =>
        window.office.getState().editor.slides[0].shapes[1].text.startsWith('Shrink this text'),
      );
      assert.ok((await state()).editor.slides[0].shapes[1].autoFitParams.fontScale < 1);
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        (expected) =>
          window.office.getState().editor.slides[0].shapes[1].text === expected &&
          !document.querySelector('[data-edit="undo"]').disabled,
        beforeAutofitText,
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        (expected) =>
          window.office.getState().editor.slides[0].shapes[1].autoFit === expected &&
          !document.querySelector('[data-edit="undo"]').disabled,
        previousAutofit,
      );
      const beforeGrowthMode = (await state()).editor.slides[0].shapes[1];
      await properties
        .getByRole('radio', { name: 'Resize shape to fit text', exact: true })
        .check();
      await settled(() => window.office.getState().editor.slides[0].shapes[1].autoFit === 'shape');
      const beforeGrowth = (await state()).editor.slides[0].shapes[1];
      const beforeMeasure = await state();
      const measured = await page.evaluate(async () => {
        const current = window.office.getState();
        const response = await fetch('/edit', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            revision: current.revision,
            preview: true,
            command: {
              type: 'update',
              slide: 0,
              ids: [current.editor.slides[0].shapes[1].id],
              changes: { format: { size: 72 } },
            },
          }),
        });
        if (!response.ok) throw new Error(await response.text());
        return response.json();
      });
      assert.equal(measured.editor.slides[0].shapes[1].format.size, 72);
      assert.equal(measured.revision, beforeMeasure.revision);
      assert.deepEqual((await state()).editor, beforeMeasure.editor);
      await page.locator('#font-size').fill('72');
      await page.locator('#font-size').press('Enter');
      await settled(() => window.office.getState().editor.slides[0].shapes[1].format.size === 72);
      assert.ok((await state()).editor.slides[0].shapes[1].bounds.h > beforeGrowth.bounds.h);
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        (expected) =>
          window.office.getState().editor.slides[0].shapes[1].format.size === expected &&
          !document.querySelector('[data-edit="undo"]').disabled,
        beforeGrowth.format.size,
      );
      assert.deepEqual((await state()).editor.slides[0].shapes[1], beforeGrowth);

      assert.notEqual(beforeGrowth.bounds.h, beforeGrowthMode.bounds.h);
      assert.equal(beforeGrowth.bounds.w, beforeGrowthMode.bounds.w);
      assert.equal(beforeGrowth.text, beforeGrowthMode.text);
      await page.locator('.shape-hit.selected').dblclick();
      const growthField = page.getByRole('textbox', { name: 'Edit text', exact: true });
      const originalHeight = await growthField.evaluate((f) => parseFloat(f.style.height));
      const growthText = Array.from({ length: 12 }, (_, i) => 'Line ' + i).join('\n');
      await growthField.fill(growthText);
      const grownHeight = await growthField.evaluate((f) => parseFloat(f.style.height));
      assert.ok(grownHeight > originalHeight);
      await page.keyboard.press('Meta+z');
      assert.equal(await growthField.evaluate((f) => parseFloat(f.style.height)), originalHeight);
      await page.keyboard.press('Meta+Shift+z');
      assert.equal(await growthField.evaluate((f) => parseFloat(f.style.height)), grownHeight);
      await page.keyboard.press('Escape');
      await settled(() =>
        window.office.getState().editor.slides[0].shapes[1].text.startsWith('Line 0'),
      );
      const grownShape = (await state()).editor.slides[0].shapes[1];
      assert.ok(grownShape.bounds.h > beforeGrowth.bounds.h);
      assert.equal(grownShape.bounds.w, beforeGrowth.bounds.w);
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        (expected) =>
          window.office.getState().editor.slides[0].shapes[1].text === expected &&
          !document.querySelector('[data-edit="undo"]').disabled,
        beforeGrowth.text,
      );
      assert.deepEqual((await state()).editor.slides[0].shapes[1].bounds, beforeGrowth.bounds);
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        (expected) =>
          window.office.getState().editor.slides[0].shapes[1].autoFit === expected &&
          !document.querySelector('[data-edit="undo"]').disabled,
        previousAutofit,
      );
      assert.deepEqual((await state()).editor.slides[0].shapes[1].bounds, beforeGrowthMode.bounds);
      for (const [direction, writingMode, textOrientation] of [
        ['wordArtVert', 'vertical-rl', 'upright'],
        ['vert', 'vertical-rl', 'sideways'],
        ['vert270', 'vertical-lr', 'sideways'],
      ]) {
        await properties
          .getByRole('combobox', { name: 'Text direction', exact: true })
          .selectOption(direction);
        await page.waitForFunction(
          (expected) =>
            window.office.getState().editor.slides[0].shapes[1].textDirection === expected &&
            !window.office.getState().building &&
            !document.querySelector('[data-edit="undo"]').disabled,
          direction,
        );
        await page.locator('.shape-hit.selected').dblclick();
        const editingField = page.getByRole('textbox', { name: 'Edit text', exact: true });
        await editingField.waitFor({ state: 'visible' });
        const verticalStyle = await editingField.evaluate((element) => {
          const style = getComputedStyle(element);
          const shape = window.office.getState().editor.slides[0].shapes[1];
          const scale = parseFloat(style.width) / shape.bounds.w;
          const margins = shape.textFrame.margins;
          const reversed = shape.textDirection === 'vert270';
          return {
            writingMode: style.writingMode,
            textOrientation: style.textOrientation,
            actualPadding: [
              style.paddingTop,
              style.paddingRight,
              style.paddingBottom,
              style.paddingLeft,
            ].map((value) => parseFloat(value) / scale),
            expectedPadding: reversed
              ? [margins.bottom, margins.left, margins.top, margins.right]
              : [margins.top, margins.right, margins.bottom, margins.left],
          };
        });
        assert.equal(verticalStyle.writingMode, writingMode);
        assert.equal(verticalStyle.textOrientation, textOrientation);
        verticalStyle.actualPadding.forEach((value, index) =>
          assert.ok(Math.abs(value - verticalStyle.expectedPadding[index]) < 200),
        );
        await editingField.press('Escape');
        await page.locator('[data-edit="undo"]').click();
        await settled(
          () => window.office.getState().editor.slides[0].shapes[1].textDirection === 'horz',
        );
      }
      await properties.getByRole('button', { name: 'Close Format Shape', exact: true }).click();
      await page.waitForFunction((width) => {
        const host = document.querySelector('#slide');
        const selected = host.shadowRoot.querySelector('.shape-hit.selected');
        const model = window.office.getState().editor;
        return (
          host.clientWidth === width &&
          Math.abs(
            selected.getBoundingClientRect().width -
              (model.slides[0].shapes[1].bounds.w * width) / model.width,
          ) < 1
        );
      }, widthBeforeProperties);
      const selectedBox = await page.locator('.shape-hit.selected').boundingBox();
      await page.mouse.move(selectedBox.x + 20, selectedBox.y + 20);
      await page.mouse.down();
      await page.mouse.move(selectedBox.x + 60, selectedBox.y + 40, { steps: 5 });
      await page.mouse.up();
      await page.waitForFunction(
        (x) => window.office.getState().editor.slides[0].shapes[1].bounds.x > x,
        shapeBefore.bounds.x,
      );
      await settled(() => !window.office.getState().building);
      const widthBefore = (await state()).editor.slides[0].shapes[1].bounds.w;
      const handle = await page.locator('.shape-hit.selected [data-handle="se"]').boundingBox();
      await page.mouse.move(handle.x + handle.width / 2, handle.y + handle.height / 2);
      await page.mouse.down();
      await page.mouse.move(handle.x + 45, handle.y + 35, { steps: 5 });
      await page.mouse.up();
      await page.waitForFunction(
        (w) => window.office.getState().editor.slides[0].shapes[1].bounds.w > w,
        widthBefore,
      );
      await settled(() => !window.office.getState().building);
      await page.locator('[data-edit="copy"]').click();
      await page.locator('[data-edit="paste"]').click();
      await settled(() => window.office.getState().editor.slides[0].shapes.length === 3);
      await page.locator('[data-edit="paste"]').click();
      await settled(() => window.office.getState().editor.slides[0].shapes.length === 4);
      // Enclosing marquee selects every object, without recording a document edit.
      const historyBeforeSelection = (await state()).history.undo;
      await page.mouse.move(canvas.x + 2, canvas.y + 2);
      await page.mouse.down();
      await page.mouse.move(canvas.x + canvas.width - 2, canvas.y + canvas.height - 2, {
        steps: 8,
      });
      await page.mouse.up();
      await page.waitForFunction(
        () =>
          document.querySelector('#slide').shadowRoot.querySelectorAll('.shape-hit.selected')
            .length === 4,
      );
      assert.equal(await page.locator('.shape-hit.selected').count(), 4);
      assert.equal((await state()).history.undo, historyBeforeSelection);
      const beforeGrouping = (await state()).editor.slides[0].shapes.map((shape) => ({
        id: shape.id,
        bounds: shape.bounds,
      }));
      await page.keyboard.press('Meta+Alt+g');
      await settled(() => window.office.getState().editor.slides[0].shapes.length === 1);
      assert.equal((await state()).editor.slides[0].shapes[0].kind, 'group');
      await page.reload();
      await page.waitForFunction(
        () => window.office?.getState().editor?.slides[0].shapes[0].kind === 'group',
      );
      await page.locator('.shape-hit').first().click();
      await page.keyboard.press('Alt+ArrowRight');
      await settled(() => window.office.getState().editor.slides[0].shapes[0].rotation === 15);
      const childBefore = (await state()).editor.slides[0].shapes[0].children[0];
      await page.locator('[data-edit="selection"]:visible').first().click();
      await page.locator('#selection-pane [aria-expanded="false"]').click();
      await page.locator(`#selection-pane [data-selection-id="${childBefore.id}"]`).click();
      assert.equal(await page.locator('.shape-hit.selected').count(), 1);
      assert.equal(
        await page.locator('.shape-hit.selected').getAttribute('data-shape-id'),
        String(childBefore.id),
      );
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(
        (old) => window.office.getState().editor.slides[0].shapes[0].children[0].bounds.x !== old,
        childBefore.bounds.x,
      );
      const childMoved = (await state()).editor.slides[0].shapes[0].children[0];
      const [a, b, c, d] = childBefore.parentTransform;
      const dx = childMoved.bounds.x - childBefore.bounds.x;
      const dy = childMoved.bounds.y - childBefore.bounds.y;
      assert.ok(Math.abs(a * dx + c * dy - 12700) < 2);
      assert.ok(Math.abs(b * dx + d * dy) < 2);
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        (old) => window.office.getState().editor.slides[0].shapes[0].children[0].bounds.x === old,
        childBefore.bounds.x,
      );
      await page.waitForFunction(() => !document.querySelector('[data-edit="undo"]').disabled);
      await page.locator('.shape-hit.selected').dblclick();
      const childTextField = page.getByRole('textbox', { name: 'Edit text', exact: true });
      await childTextField.fill('Edited inside rotated group');
      await page.keyboard.press('Escape');
      await settled(
        () =>
          window.office.getState().editor.slides[0].shapes[0].children[0].text ===
          'Edited inside rotated group',
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        (text) => window.office.getState().editor.slides[0].shapes[0].children[0].text === text,
        childBefore.text,
      );
      await page.waitForFunction(() => !document.querySelector('[data-edit="undo"]').disabled);
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].shapes[0].rotation === 0,
      );
      await page.waitForFunction(() => !document.querySelector('[data-edit="undo"]').disabled);
      const groupId = (await state()).editor.slides[0].shapes[0].id;
      await page.locator(`#selection-pane [data-selection-id="${groupId}"]`).click();
      await page
        .locator('#selection-pane')
        .getByRole('button', { name: 'Close Selection Pane', exact: true })
        .click();
      await page.locator('#slide').focus();
      assert.equal(
        await page.locator('.shape-hit.selected').getAttribute('data-shape-id'),
        String(groupId),
      );
      await page.keyboard.press('Meta+Alt+Shift+g');
      await settled(() => window.office.getState().editor.slides[0].shapes.length === 4);
      assert.deepEqual(
        (await state()).editor.slides[0].shapes.map((shape) => ({
          id: shape.id,
          bounds: shape.bounds,
        })),
        beforeGrouping,
      );
      await page.waitForFunction(
        () =>
          document.querySelector('#slide').shadowRoot.querySelectorAll('.shape-hit.selected')
            .length === 4,
      );
      assert.equal(await page.locator('.shape-hit.selected').count(), 4);
      const historyBeforeDistribution = (await state()).history.undo;
      const positionsBeforeDistribution = (await state()).editor.slides[0].shapes.map(
        (s) => s.bounds,
      );
      const distributionOrder = (await state()).editor.slides[0].shapes
        .sort((a, b) => a.bounds.x - b.bounds.x)
        .map((s) => s.id);
      await page.locator('[data-edit="arrange"]:visible').click();
      await page.getByRole('menuitem', { name: 'Align', exact: true }).click();
      await page.getByRole('menuitem', { name: 'Distribute Horizontally', exact: true }).click();
      await settled(() => !window.office.getState().building);
      await page.waitForFunction(
        (n) => window.office.getState().history.undo === n + 1,
        historyBeforeDistribution,
      );
      const distributedShapes = (await state()).editor.slides[0].shapes;
      const boxes = distributionOrder.map(
        (id) => distributedShapes.find((s) => s.id === id).bounds,
      );
      const gaps = boxes.slice(1).map((b, i) => b.x - boxes[i].x - boxes[i].w);
      assert.ok(Math.max(...gaps) - Math.min(...gaps) <= 2);
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        (n) => window.office.getState().history.undo === n,
        historyBeforeDistribution,
      );
      await settled(() => !window.office.getState().building);
      assert.deepEqual(
        (await state()).editor.slides[0].shapes.map((s) => s.bounds),
        positionsBeforeDistribution,
      );
      await page.keyboard.press('Escape');
      await page.locator('.shape-hit').first().click();
      await page.getByRole('button', { name: 'Strikethrough', exact: true }).click();
      await settled(
        () => window.office.getState().editor.slides[0].shapes[0].format.strike === true,
      );
      await page.getByRole('button', { name: 'Bullets', exact: true }).click();
      await settled(() => window.office.getState().editor.slides[0].shapes[0].bullets === 'bullet');
      await page.getByRole('button', { name: 'Text Highlight Color', exact: true }).click();
      await page.getByRole('menuitem', { name: 'Yellow', exact: true }).click();
      await settled(
        () => window.office.getState().editor.slides[0].shapes[0].format.highlight === '#FFFF00',
      );
      assert.ok(
        await page
          .locator('#slide span')
          .evaluateAll((spans) =>
            spans.some((span) => getComputedStyle(span).backgroundColor === 'rgb(255, 255, 0)'),
          ),
      );
      await page.locator('.shape-hit').first().click();
      await page.keyboard.press('F2');
      const textField = page.getByRole('textbox', { name: 'Edit text', exact: true });
      await textField.press('ArrowLeft');
      for (let i = 0; i < 3; i++) await textField.press('Shift+ArrowRight');
      assert.deepEqual(
        await textField.evaluate((field) => [field.selectionStart, field.selectionEnd]),
        [0, 3],
      );
      await page.locator('[data-edit="italic"]').click();
      await settled(() => {
        const runs = window.office.getState().editor.slides[0].shapes[0].runs;
        return runs.length >= 2 && runs[0].end === 3 && runs[0].format.italic === true;
      });
      assert.notEqual((await state()).editor.slides[0].shapes[0].runs[1].format.italic, true);
      await page.getByRole('button', { name: 'Text Highlight Color', exact: true }).click();
      await page.getByRole('menuitem', { name: 'Cyan', exact: true }).click();
      await settled(
        () =>
          window.office.getState().editor.slides[0].shapes[0].runs[0].format.highlight ===
          '#00FFFF',
      );
      assert.equal((await state()).editor.slides[0].shapes[0].runs[1].format.highlight, '#FFFF00');
      const visibleRuns = await textField.locator('span').evaluateAll((spans) =>
        spans.map((span) => ({
          text: span.textContent,
          italic: getComputedStyle(span).fontStyle,
          highlight: getComputedStyle(span).backgroundColor,
        })),
      );
      assert.equal(visibleRuns[0].text, '編集済');
      assert.equal(visibleRuns[0].italic, 'italic');
      assert.equal(visibleRuns[0].highlight, 'rgb(0, 255, 255)');
      assert.equal(visibleRuns.at(-1).highlight, 'rgb(255, 255, 0)');
      await textField.press('Control+i');
      await settled(
        () => window.office.getState().editor.slides[0].shapes[0].runs[0].format.italic === false,
      );
      await textField.press('Escape');
      await page.reload();
      await page.waitForFunction(
        () =>
          window.office?.getState().editor?.slides[0].shapes[0].runs[0].format.highlight ===
          '#00FFFF',
      );
      await page.locator('.shape-hit').first().click();
      await page.keyboard.press('F2');
      await textField.press('ArrowLeft');
      await textField.press('Enter');
      assert.equal(await textField.textContent(), '\n編集済み title');
      await textField.press('Control+z');
      assert.equal(await textField.textContent(), '編集済み title');
      await textField.press('Control+Shift+z');
      assert.equal(await textField.textContent(), '\n編集済み title');
      await textField.press('Backspace');
      assert.equal(await textField.textContent(), '編集済み title');
      await textField.evaluate((field) => {
        field.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
        for (const value of ['に', '日本語']) {
          field.dispatchEvent(
            new InputEvent('beforeinput', {
              inputType: 'insertCompositionText',
              isComposing: true,
              bubbles: true,
            }),
          );
          field.value = value + '編集済み title';
          field.setSelectionRange(value.length, value.length);
          field.dispatchEvent(
            new InputEvent('input', {
              inputType: 'insertCompositionText',
              isComposing: true,
              bubbles: true,
            }),
          );
        }
        field.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true }));
      });
      assert.equal(await textField.textContent(), '日本語編集済み title');
      await textField.press('Control+z');
      assert.equal(await textField.textContent(), '編集済み title');
      await textField.pressSequentially('A');
      for (let i = 0; i < 3; i++) await textField.press('ArrowRight');
      await textField.press('Backspace');
      await textField.press('Escape');
      await settled(
        () => window.office.getState().editor.slides[0].shapes[0].text === 'A編集み title',
      );
      const typedRuns = (await state()).editor.slides[0].shapes[0].runs;
      assert.equal(typedRuns[0].format.highlight, '#00FFFF');
      assert.equal(typedRuns.at(-1).format.highlight, '#FFFF00');
      await page.reload();
      await page.waitForFunction(
        () => window.office?.getState().editor?.slides[0].shapes[0].text === 'A編集み title',
      );
      assert.equal(
        (await state()).editor.slides[0].shapes[0].runs.at(-1).format.highlight,
        '#FFFF00',
      );
      await page.locator('[data-edit="undo"]').click();
      await settled(
        () => window.office.getState().editor.slides[0].shapes[0].text === '編集済み title',
      );
      await page.locator('.shape-hit').first().click();
      await page.keyboard.press('F2');
      await textField.press('ArrowLeft');
      const historyBeforeTyping = (await state()).history.undo;
      await page.locator('[data-edit="italic"]').click();
      assert.equal(await page.locator('[data-edit="italic"]').getAttribute('aria-pressed'), 'true');
      assert.equal((await state()).history.undo, historyBeforeTyping);
      await textField.pressSequentially('NEW');
      await page.locator('[data-edit="italic"]').click();
      assert.equal(
        await page.locator('[data-edit="italic"]').getAttribute('aria-pressed'),
        'false',
      );
      await textField.pressSequentially('!');
      await textField.press('Escape');
      await settled(
        () => window.office.getState().editor.slides[0].shapes[0].text === 'NEW!編集済み title',
      );
      const insertedRuns = (await state()).editor.slides[0].shapes[0].runs;
      assert.ok(insertedRuns.filter((r) => r.end <= 3).every((r) => r.format.italic === true));
      assert.equal(insertedRuns.find((r) => r.start === 3).format.italic, false);
      assert.equal(insertedRuns.find((r) => r.start === 4).format.italic, false);
      await page.locator('[data-edit="undo"]').click();
      await settled(
        () => window.office.getState().editor.slides[0].shapes[0].text === '編集済み title',
      );
      await page.locator('.shape-hit').first().click();
      await page.locator('[data-edit="text-anchor"]:visible').click();
      await page.getByRole('menuitem', { name: 'Middle', exact: true }).click();
      await settled(() => window.office.getState().editor.slides[0].shapes[0].anchor === 'center');
      await page.locator('.shape-hit').first().click();
      await page.keyboard.press('F2');
      const centeredText = await textField.evaluate((field) => {
        const box = field.getBoundingClientRect();
        const range = document.createRange();
        range.selectNodeContents(field);
        const text = range.getBoundingClientRect();
        return Math.abs((text.top + text.bottom) / 2 - (box.top + box.bottom) / 2);
      });
      assert.ok(centeredText < 5, `Expected vertically centered text, deviation ${centeredText}px`);
      await textField.press('Escape');
      await page.locator('[data-edit="undo"]').click();
      await settled(() => window.office.getState().editor.slides[0].shapes[0].anchor !== 'center');
      await page.locator('.shape-hit').first().click();
      await page.keyboard.press('F2');
      await textField.fill('one\ntwo\nthree');
      await textField.evaluate((field) => field.setSelectionRange(4, 7));
      await page.locator('[data-edit="align-center"]:visible').click();
      await settled(
        () => window.office.getState().editor.slides[0].shapes[0].paragraphs[1]?.align === 'ctr',
      );
      const paragraphStyles = (await state()).editor.slides[0].shapes[0].paragraphs;
      assert.notEqual(paragraphStyles[0].align, 'ctr');
      assert.notEqual(paragraphStyles[2].align, 'ctr');
      const editingParagraphs = await textField
        .locator('.edit-paragraph')
        .evaluateAll((paragraphs) =>
          paragraphs.map((paragraph) => ({
            text: paragraph.textContent,
            align: getComputedStyle(paragraph).textAlign,
          })),
        );
      assert.deepEqual(
        editingParagraphs.map((p) => p.text),
        ['one\n', 'two\n', 'three'],
      );
      assert.equal(editingParagraphs[1].align, 'center');
      assert.notEqual(editingParagraphs[0].align, 'center');
      assert.notEqual(editingParagraphs[2].align, 'center');
      const lineBoxes = await textField.locator('.edit-paragraph').evaluateAll((paragraphs) =>
        paragraphs.map((paragraph) => ({
          height: paragraph.getBoundingClientRect().height,
          lineHeight: parseFloat(getComputedStyle(paragraph).lineHeight),
          bullet: getComputedStyle(paragraph, '::before').content,
        })),
      );
      assert.ok(lineBoxes.every((line) => line.height < line.lineHeight * 1.2));
      assert.ok(lineBoxes.every((line) => line.bullet.includes('•')));
      await textField.evaluate((field) => field.setSelectionRange(3, 3));
      await textField.press('ArrowRight');
      assert.deepEqual(
        await textField.evaluate((field) => [field.selectionStart, field.selectionEnd]),
        [4, 4],
      );
      await textField.pressSequentially('X');
      assert.equal(await textField.textContent(), 'one\nXtwo\nthree');
      await textField.press('Control+z');
      await textField.press('Backspace');
      assert.equal(await textField.textContent(), 'onetwo\nthree');
      await textField.press('Control+z');
      assert.equal(await textField.textContent(), 'one\ntwo\nthree');
      await textField.evaluate((field) => field.setSelectionRange(4, 7));
      assert.deepEqual(
        await textField.evaluate((field) => [field.selectionStart, field.selectionEnd]),
        [4, 7],
      );
      await textField.press('Escape');
      await page.locator('[data-edit="undo"]').click();
      await settled(
        () => window.office.getState().editor.slides[0].shapes[0].paragraphs[1]?.align !== 'ctr',
      );
      await page.locator('[data-edit="undo"]').click();
      await settled(
        () => window.office.getState().editor.slides[0].shapes[0].text === '編集済み title',
      );
      await page.getByRole('button', { name: 'Picture ▾', exact: true }).first().click();
      const chooserPromise = page.waitForEvent('filechooser');
      await page.getByRole('menuitem', { name: 'Picture from File...', exact: true }).click();
      const chooser = await chooserPromise;
      await chooser.setFiles({
        name: 'sample.png',
        mimeType: 'image/png',
        buffer: Buffer.from(
          'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jp1sAAAAASUVORK5CYII=',
          'base64',
        ),
      });
      await settled(() => window.office.getState().editor.slides[0].shapes.length === 5);
      assert.equal((await state()).editor.slides[0].shapes.at(-1).kind, 'picture');
      const picture = (await state()).editor.slides[0].shapes.at(-1);
      assert.equal(picture.aspectRatioLocked, true);
      await page.locator(`[data-shape-id="${picture.id}"]`).click();
      await page.keyboard.press('Shift+ArrowRight');
      await page.waitForFunction(
        (width) => window.office.getState().editor.slides[0].shapes.at(-1).bounds.w === width,
        Math.round(picture.bounds.w * 1.1),
      );
      await settled(() => !window.office.getState().building);
      const resizedPicture = (await state()).editor.slides[0].shapes.at(-1).bounds;
      assert.equal(resizedPicture.h, Math.round(picture.bounds.h * 1.1));
      assert.ok(
        Math.abs(
          resizedPicture.x + resizedPicture.w / 2 - picture.bounds.x - picture.bounds.w / 2,
        ) <= 1,
      );
      assert.ok(
        Math.abs(
          resizedPicture.y + resizedPicture.h / 2 - picture.bounds.y - picture.bounds.h / 2,
        ) <= 1,
      );
      await page.keyboard.press('Meta+z');
      await page.waitForFunction(
        (bounds) =>
          JSON.stringify(window.office.getState().editor.slides[0].shapes.at(-1).bounds) ===
          JSON.stringify(bounds),
        picture.bounds,
      );
      await settled(() => !window.office.getState().building);
      assert.ok((await page.locator('#slide image').count()) > 0);
      const slideCount = (await state()).editor.slides.length;
      await page.getByRole('button', { name: 'New Slide Layout', exact: true }).click();
      const layoutMenu = page.getByRole('menu', { name: 'New Slide', exact: true });
      assert.ok((await layoutMenu.locator('svg').count()) >= 3);
      await page.getByRole('menuitem', { name: 'Title and Content', exact: true }).click();
      await page.waitForFunction(
        (count) => window.office.getState().editor.slides.length === count + 1,
        slideCount,
      );
      await settled(() => !window.office.getState().building);
      const insertedSlide = (await state()).editor.slides[(await state()).index];
      const contentLayout = (await state()).editor.layouts.find((layout) => layout.type === 'obj');
      assert.equal(insertedSlide.layout, contentLayout.key);
      assert.ok(insertedSlide.shapes.length >= 2);
      const emptyTitle = insertedSlide.shapes.find((shape) => shape.placeholder === 'title');
      assert.ok(emptyTitle);
      await page.locator(`[data-shape-id="${emptyTitle.id}"]`).click();
      await page.getByRole('textbox', { name: 'Edit text', exact: true }).fill('Layout title');
      await page.getByRole('textbox', { name: 'Edit text', exact: true }).press('Escape');
      await settled(() =>
        window.office
          .getState()
          .editor.slides[1].shapes.some((shape) => shape.text === 'Layout title'),
      );
      await page.locator('[data-edit="undo"]').click();
      await settled(() =>
        window.office.getState().editor.slides[1].shapes.every((shape) => shape.text === ''),
      );
      await page.getByRole('button', { name: 'Layout', exact: true }).click();
      await page
        .getByRole('menu', { name: 'Layout', exact: true })
        .getByRole('menuitemcheckbox', { name: 'Title Slide', exact: true })
        .click();
      await settled(() =>
        window.office
          .getState()
          .editor.slides[1].shapes.some((shape) => shape.placeholder === 'ctrTitle'),
      );
      await page.getByRole('button', { name: 'Reset', exact: true }).click();
      await settled(() => !window.office.getState().building);
      await page.locator('[data-edit="undo"]').click();
      await settled(() => !window.office.getState().building);
      await page.locator('[data-edit="undo"]').click();
      await settled(() =>
        window.office
          .getState()
          .editor.slides[1].shapes.some((shape) => shape.placeholder === 'title'),
      );
      await page.reload();
      await page.waitForFunction(
        (count) => window.office?.getState().editor?.slides.length === count + 1,
        slideCount,
      );
      assert.equal((await state()).editor.slides[1].layout, contentLayout.key);
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        (count) => window.office.getState().editor.slides.length === count,
        slideCount,
      );
      await settled(() => !window.office.getState().building);
      await page.getByRole('button', { name: 'Section', exact: true }).click();
      await page.getByRole('menuitem', { name: 'Add Section', exact: true }).click();
      const sectionDialog = page.getByRole('dialog', { name: 'Rename Section' });
      await sectionDialog.getByLabel('Section name:').fill('Overview');
      await sectionDialog.getByRole('button', { name: 'Rename', exact: true }).click();
      await settled(() => window.office.getState().editor.sections[0]?.name === 'Overview');
      await sectionDialog.waitFor({ state: 'hidden' });
      const sectionHeader = page.locator('.section-heading').first();
      await sectionHeader.click();
      assert.equal(await page.locator('.thumbnail:visible').count(), slideCount);
      await sectionHeader.locator('.section-disclosure').click();
      assert.equal(await page.locator('.thumbnail:visible').count(), 0);
      await sectionHeader.locator('.section-disclosure').click();
      assert.equal(await page.locator('.thumbnail:visible').count(), slideCount);
      await page.reload();
      await page.waitForFunction(
        () => window.office?.getState().editor?.sections[0]?.name === 'Overview',
      );
      await page.locator('.section-heading').first().click({ button: 'right' });
      await page.getByRole('menuitem', { name: 'Remove All Sections', exact: true }).click();
      await settled(() => window.office.getState().editor.sections.length === 0);
      assert.equal(await page.locator('.section-heading').count(), 0);
      await page.locator('[data-edit="undo"]').click();
      await settled(() => window.office.getState().editor.sections[0]?.name === 'Overview');
      const sectionOrder = (await state()).editor.slides.map((slide) => slide.key);
      await page.getByRole('button', { name: 'Slide 2', exact: true }).click();
      await page.getByRole('button', { name: 'Section', exact: true }).click();
      await page.getByRole('menuitem', { name: 'Add Section', exact: true }).click();
      await sectionDialog.getByLabel('Section name:').fill('Details');
      await sectionDialog.getByRole('button', { name: 'Rename', exact: true }).click();
      await settled(() => window.office.getState().editor.sections[1]?.name === 'Details');
      await sectionDialog.waitFor({ state: 'hidden' });
      await page.locator('.section-heading').nth(1).click({ button: 'right' });
      await page.getByRole('menuitem', { name: 'Move Section Up', exact: true }).click();
      await settled(() => window.office.getState().editor.sections[0]?.name === 'Details');
      assert.deepEqual(
        (await state()).editor.slides.map((slide) => slide.key),
        [...sectionOrder.slice(1), sectionOrder[0]],
      );
      await page.locator('.section-heading').nth(0).dragTo(page.locator('.section-heading').nth(1));
      await settled(() => window.office.getState().editor.sections[0]?.name === 'Overview');
      assert.deepEqual(
        (await state()).editor.slides.map((slide) => slide.key),
        sectionOrder,
      );
      await page.locator('.section-heading').nth(1).click();
      await page.keyboard.press('Meta+ArrowUp');
      await settled(() => window.office.getState().editor.sections[0]?.name === 'Details');
      assert.equal(await page.locator('.section-heading:focus').getAttribute('data-section'), '0');
      await page.keyboard.press('Meta+Shift+ArrowDown');
      await settled(() => window.office.getState().editor.sections[1]?.name === 'Details');
      await page.keyboard.press('Shift+F10');
      await page.getByRole('menuitem', { name: 'Remove Section & Slides', exact: true }).click();
      await settled(() => window.office.getState().editor.slides.length === 1);
      await page.locator('[data-edit="undo"]').click();
      await settled(() => window.office.getState().editor.sections.length === 2);
      assert.deepEqual(
        (await state()).editor.slides.map((slide) => slide.key),
        sectionOrder,
      );
      const exported = await loadPresentation(
        await (await page.request.get(url + '/deck.pptx')).body(),
      );
      assert.match(getSlideText(getSlides(exported)[0]), /編集済み title/);
      const exportedRuns = getShapeParagraphElements(getSlideShapes(getSlides(exported)[0])[0], 0);
      assert.equal(exportedRuns[0].text, '編集済');
      assert.equal(exportedRuns[0].format.highlight, '#00FFFF');
      assert.equal(exportedRuns[1].format.highlight, '#FFFF00');
      const rebuilt = await buildDeck(file);
      assert.equal(rebuilt.editor.slides[0].shapes.length, 5);
      assert.equal(rebuilt.editor.slides[0].shapes[0].runs[0].format.highlight, '#00FFFF');
      const denied = await page.request.post(url + '/edit', {
        data: { revision: (await state()).revision, action: 'undo' },
      });
      assert.equal(denied.status(), 403);
      const stale = await page.request.post(url + '/edit', {
        headers: { Origin: url },
        data: { revision: -1, action: 'undo' },
      });
      assert.equal(stale.status(), 409);
      await page.getByRole('button', { name: 'Slide 1', exact: true }).click();
      await page.getByRole('button', { name: 'Section', exact: true }).click();
      await page.getByRole('menuitem', { name: 'Add Section', exact: true }).click();
      await sectionDialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      await settled(() => window.office.getState().editor.sections.length === 3);
      assert.equal((await state()).editor.sections[0].slides.length, 0);
      for (let remaining = 2; remaining >= 1; remaining--) {
        await page.locator('.section-heading').last().click({ button: 'right' });
        await page.getByRole('menuitem', { name: 'Remove Section & Slides', exact: true }).click();
        await page.waitForFunction(
          (count) =>
            window.office.getState().editor.sections.length === count &&
            !window.office.getState().building,
          remaining,
        );
        await settled(() => !window.office.getState().building);
      }
      assert.equal((await state()).editor.slides.length, 0);
      assert.equal(await page.locator('.empty-section-list .section-heading').count(), 1);
      await page.locator('.empty-section-list .section-heading').click({ button: 'right' });
      await page.getByRole('menuitem', { name: 'Rename Section', exact: true }).click();
      await sectionDialog.getByLabel('Section name:').fill('Empty deck');
      await sectionDialog.getByRole('button', { name: 'Rename', exact: true }).click();
      await settled(() => window.office.getState().editor.sections[0]?.name === 'Empty deck');
      await sectionDialog.waitFor({ state: 'hidden' });
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      proc.kill('SIGTERM');
      await new Promise((resolve) =>
        proc.exitCode !== null ? resolve() : proc.once('exit', resolve),
      );
      await rm(dir, { recursive: true, force: true });
    }
  },
);

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
  renameShape,
  getShapeName,
  getShapeId,
  getShapeZIndex,
  getShapeBounds,
  getShapeRotation,
  getShapeFlip,
  getSlideShapes,
  getSlides,
  isShapeHidden,
  savePresentation,
  loadPresentation,
  inches,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'selection pane names, visibility, nested selection and undo persist in both languages',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-selection-pane-'));
    let preview, browser;
    try {
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      const first = addSlideTextBox(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(2),
        h: inches(1),
        text: 'First',
      });
      const second = addSlideTextBox(slide, {
        x: inches(4),
        y: inches(1),
        w: inches(2),
        h: inches(1),
        text: 'Second',
      });
      renameShape(first, 'First');
      renameShape(second, 'Second');
      const group = groupShapes([first, second]);
      renameShape(group, 'Group');
      const third = addSlideTextBox(slide, {
        x: inches(1),
        y: inches(4),
        w: inches(2),
        h: inches(1),
        text: 'Third',
      });
      renameShape(third, 'Third');
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
      const state = async () =>
        getSlideShapes(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        );
      await saved();
      await editor.getByRole('button', { name: 'Arrange', exact: true }).click();
      await editor
        .getByRole('menuitemcheckbox', { name: 'Selection Pane...', exact: true })
        .click();
      let pane = editor.getByRole('region', { name: 'Selection Pane', exact: true });
      assert.deepEqual(await pane.locator('.name').allTextContents(), ['Third', 'Group']);
      assert.equal(await pane.getByRole('button', { name: 'Show All', exact: true }).count(), 0);
      assert.equal(
        await pane.getByRole('button', { name: 'Bring Forward', exact: true }).isDisabled(),
        true,
      );
      await pane.getByRole('button', { name: 'Third', exact: true }).click();
      const arrange = editor.getByRole('button', { name: 'Arrange', exact: true });
      await arrange.click();
      await editor.getByRole('menuitem', { name: 'Align', exact: true }).focus();
      await page.keyboard.press('ArrowRight');
      await editor.getByRole('menuitem', { name: 'Align Right', exact: true }).click();
      await saved();
      assert.ok(
        getShapeBounds((await state()).find((shape) => getShapeName(shape) === 'Third')).x >
          inches(5),
      );
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      await arrange.click();
      await editor.getByRole('menuitem', { name: 'Rotate', exact: true }).click();
      await editor.getByRole('menuitem', { name: 'Rotate Right 90°', exact: true }).click();
      await saved();
      assert.equal(
        getShapeRotation((await state()).find((shape) => getShapeName(shape) === 'Third')),
        90,
      );
      await arrange.click();
      await editor.getByRole('menuitem', { name: 'Rotate', exact: true }).click();
      await editor.getByRole('menuitem', { name: 'Flip Horizontal', exact: true }).click();
      await saved();
      assert.equal(
        getShapeFlip((await state()).find((shape) => getShapeName(shape) === 'Third')).horizontal,
        true,
      );
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      await pane.getByRole('button', { name: 'Group', exact: true }).click();
      await pane.getByRole('button', { name: 'Bring Forward', exact: true }).click();
      await saved();
      assert.deepEqual(await pane.locator('.name').allTextContents(), ['Group', 'Third']);
      await pane.getByRole('button', { name: 'Send Backward', exact: true }).click();
      await saved();
      assert.deepEqual(await pane.locator('.name').allTextContents(), ['Third', 'Group']);
      await pane
        .getByRole('button', { name: 'Group', exact: true })
        .dragTo(pane.getByRole('button', { name: 'Third', exact: true }), {
          targetPosition: { x: 20, y: 2 },
        });
      await saved();
      assert.deepEqual(await pane.locator('.name').allTextContents(), ['Group', 'Third']);
      assert.equal(
        getShapeZIndex((await state()).find((shape) => getShapeName(shape) === 'Group')),
        1,
      );
      await pane.getByRole('button', { name: 'Group', exact: true }).press('Meta+z');
      await saved();
      assert.deepEqual(await pane.locator('.name').allTextContents(), ['Third', 'Group']);
      await pane.getByRole('button', { name: 'Expand Group', exact: true }).click();
      assert.deepEqual(await pane.locator('.name').allTextContents(), [
        'Third',
        'Group',
        'Second',
        'First',
      ]);
      await pane
        .getByRole('button', { name: 'First', exact: true })
        .dragTo(pane.getByRole('button', { name: 'Second', exact: true }), {
          targetPosition: { x: 20, y: 2 },
        });
      await saved();
      assert.deepEqual(await pane.locator('.name').allTextContents(), [
        'Third',
        'Group',
        'First',
        'Second',
      ]);
      assert.equal(
        getShapeZIndex((await state()).find((shape) => getShapeName(shape) === 'First')),
        1,
      );
      await pane
        .getByRole('button', { name: 'First', exact: true })
        .dragTo(pane.getByRole('button', { name: 'Third', exact: true }), {
          targetPosition: { x: 20, y: 2 },
        });
      assert.deepEqual(await pane.locator('.name').allTextContents(), [
        'Third',
        'Group',
        'First',
        'Second',
      ]);
      await pane.getByRole('button', { name: 'First', exact: true }).press('Meta+z');
      await saved();
      await pane.getByRole('button', { name: 'Expand Group', exact: true }).click();
      await pane.getByRole('button', { name: 'Second', exact: true }).click();
      await pane
        .getByRole('button', { name: 'First', exact: true })
        .click({ modifiers: ['Shift'] });
      assert.equal(await pane.locator('.name[aria-pressed="true"]').count(), 2);
      await pane.getByRole('button', { name: 'Third', exact: true }).click({ modifiers: ['Meta'] });
      assert.equal(await pane.locator('.name[aria-pressed="true"]').count(), 1);
      await pane.getByRole('button', { name: 'Third', exact: true }).press('F2');
      await pane.getByRole('textbox', { name: 'Object name' }).fill('Cancelled');
      await pane.getByRole('textbox', { name: 'Object name' }).press('Escape');
      await pane.getByRole('button', { name: 'Third', exact: true }).press('Space');
      assert.equal(await editor.locator('[contenteditable="true"]').count(), 0);
      await pane.getByRole('button', { name: 'Third', exact: true }).dblclick();
      await pane.getByRole('textbox', { name: 'Object name' }).fill('Renamed');
      await pane.getByRole('textbox', { name: 'Object name' }).press('Enter');
      await saved();
      assert.equal(
        getShapeName((await state()).find((shape) => getShapeId(shape) === getShapeId(third))),
        'Renamed',
      );
      assert.equal(await editor.locator('.hit').count(), 2);
      assert.equal(await editor.locator('.paint').getByText('Third', { exact: true }).count(), 1);
      await pane.getByRole('button', { name: 'Hide object: Renamed', exact: true }).click();
      await saved();
      assert.equal(
        isShapeHidden((await state()).find((shape) => getShapeName(shape) === 'Renamed')),
        true,
      );
      assert.equal(await editor.locator('.hit').count(), 1);
      assert.equal(await editor.locator('.paint').getByText('Third', { exact: true }).count(), 0);
      await pane.getByRole('button', { name: 'Show object: Renamed', exact: true }).press('Meta+z');
      await saved();
      assert.equal(
        isShapeHidden((await state()).find((shape) => getShapeName(shape) === 'Renamed')),
        false,
      );
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      pane = editor.getByRole('region', { name: '選択ウィンドウ', exact: true });
      await page.screenshot({ path: '/tmp/pptx-selection-pane.png', fullPage: true });
      await pane.getByRole('button', { name: 'すべて非表示', exact: true }).click();
      await saved();
      assert.ok((await state()).every(isShapeHidden));
      assert.equal(await editor.locator('.hit').count(), 0);
      await pane.getByRole('button', { name: 'すべて表示', exact: true }).click();
      await saved();
      assert.ok((await state()).every((shape) => !isShapeHidden(shape)));
      await pane.getByRole('button', { name: '選択ウィンドウを閉じる', exact: true }).click();
      assert.equal(await pane.count(), 0);
      await page.reload();
      await editor.locator('button[aria-haspopup="menu"]').filter({ hasText: '配置' }).click();
      await editor
        .getByRole('menuitemcheckbox', { name: '選択ウィンドウ...', exact: true })
        .click();
      await editor
        .getByRole('region', { name: '選択ウィンドウ', exact: true })
        .getByRole('button', { name: 'Renamed', exact: true })
        .waitFor();
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'selection pane scrolls during a drag and saves the offscreen stacking order',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-selection-scroll-'));
    let preview, browser;
    try {
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      const count = 45;
      for (let i = 0; i < count; i++) {
        renameShape(
          addSlideTextBox(slide, {
            x: inches(1),
            y: inches(1),
            w: inches(2),
            h: inches(1),
            text: `Object ${i}`,
          }),
          `Object ${i}`,
        );
      }
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
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      await saved();
      await editor.getByRole('button', { name: 'Arrange', exact: true }).click();
      await editor
        .getByRole('menuitemcheckbox', { name: 'Selection Pane...', exact: true })
        .click();
      const pane = editor.getByRole('region', { name: 'Selection Pane', exact: true });
      const list = pane.locator('.objects');
      const first = pane.getByRole('button', { name: 'Object 44', exact: true });
      const origin = await first.boundingBox();
      const bounds = await list.boundingBox();
      const x = origin.x + origin.width / 2;
      await page.mouse.move(x, origin.y + origin.height / 2);
      await page.mouse.down();
      await page.mouse.move(x, origin.y + origin.height + 10, { steps: 5 });
      await page.mouse.move(x, bounds.y + bounds.height - 3, { steps: 8 });
      await list.evaluate(
        (element) =>
          new Promise((resolve, reject) => {
            const deadline = performance.now() + 10000;
            const check = () => {
              if (element.scrollTop + element.clientHeight >= element.scrollHeight - 2) resolve();
              else if (performance.now() >= deadline)
                reject(new Error('Drag did not scroll to the last object'));
              else requestAnimationFrame(check);
            };
            check();
          }),
      );
      const last = await pane.getByRole('button', { name: 'Object 0', exact: true }).boundingBox();
      await page.mouse.move(x, last.y + last.height - 2, { steps: 3 });
      await page.mouse.up();
      await saved();
      assert.equal((await pane.locator('.name').allTextContents()).at(-1), 'Object 44');
      const loaded = await loadPresentation(
        new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
      );
      assert.equal(getShapeName(getSlideShapes(getSlides(loaded)[0])[0]), 'Object 44');
      const stoppedAt = await list.evaluate((element) => element.scrollTop);
      await list
        .evaluate(
          (element) =>
            new Promise((resolve) =>
              requestAnimationFrame(() => requestAnimationFrame(() => resolve(element.scrollTop))),
            ),
        )
        .then((value) => assert.equal(value, stoppedAt));
      await pane.getByRole('button', { name: 'Object 44', exact: true }).press('Meta+z');
      await saved();
      assert.equal((await pane.locator('.name').allTextContents())[0], 'Object 44');
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

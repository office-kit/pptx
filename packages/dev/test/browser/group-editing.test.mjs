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
  getShapeId,
  getGroupChildren,
  getShapeRotation,
  setShapeRotation,
  getShapeBounds,
  getShapeText,
  getSlides,
  getSlideShapes,
  savePresentation,
  loadPresentation,
  inches,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'edit a rotated group child in place, undo, leave the group and reload in both languages',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-group-edit-'));
    let preview, browser, page;
    try {
      const pres = createPresentation(),
        slide = addBlankSlide(pres);
      const first = addSlideTextBox(slide, {
        x: inches(2),
        y: inches(2),
        w: inches(2),
        h: inches(1),
        text: '日本語',
      });
      const second = addSlideTextBox(slide, {
        x: inches(5),
        y: inches(2),
        w: inches(2),
        h: inches(1),
        text: 'English',
      });
      const group = groupShapes([first, second]);
      setShapeRotation(group, 90);
      const original = getShapeBounds(first),
        sibling = getShapeBounds(second),
        id = getShapeId(first);
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
      await editor.locator('.hit').dblclick();
      await editor.getByText('Editing group', { exact: true }).waitFor();
      assert.equal(await editor.locator('.hit').count(), 2);
      const bounds = await editor.locator('.hit.selected').boundingBox();
      await page.mouse.move(bounds.x + bounds.width / 2, bounds.y + bounds.height / 2);
      await page.mouse.down();
      await page.mouse.move(bounds.x + bounds.width / 2 + 40, bounds.y + bounds.height / 2, {
        steps: 8,
      });
      await page.mouse.up();
      await saved();
      let shapes = await state();
      const moved = getShapeBounds(shapes.find((s) => getShapeId(s) === id));
      assert.ok(Math.abs(moved.x - original.x) < 5);
      assert.ok(moved.y < original.y - 100000);
      assert.deepEqual(
        getShapeBounds(shapes.find((s) => getShapeId(s) === getShapeId(second))),
        sibling,
      );
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(getShapeBounds((await state()).find((s) => getShapeId(s) === id)), original);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      await page.keyboard.press('ArrowRight');
      await saved();
      assert.ok(getShapeBounds((await state()).find((s) => getShapeId(s) === id)).y < moved.y);
      const beforeAlign = getShapeBounds((await state()).find((s) => getShapeId(s) === id));
      const childBox = await editor.locator('.hit.selected').boundingBox();
      const stageBox = await editor.locator('.stage').boundingBox();
      await page.mouse.move(childBox.x + childBox.width / 2, childBox.y + childBox.height / 2);
      await page.mouse.down();
      await page.mouse.move(stageBox.x + childBox.width / 2 + 4, childBox.y + childBox.height / 2, {
        steps: 8,
      });
      const guide = editor.locator('.guide.v').first();
      await guide.waitFor();
      assert.ok(Math.abs((await guide.boundingBox()).x - stageBox.x) < 2);
      await page.mouse.up();
      await saved();
      assert.ok(
        Math.abs(getShapeBounds((await state()).find((s) => getShapeId(s) === id)).y - inches(6)) <
          5,
      );
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      await editor.getByRole('tab', { name: 'Size & Properties' }).click();
      await editor.getByRole('button', { name: 'Align left', exact: true }).click();
      await saved();
      const aligned = getShapeBounds((await state()).find((s) => getShapeId(s) === id));
      // The parent is rotated 90°: the visible left edge is 6 inches minus local y.
      assert.ok(Math.abs(aligned.y - inches(6)) < 5);
      assert.equal(aligned.x, beforeAlign.x);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(
        getShapeBounds((await state()).find((s) => getShapeId(s) === id)),
        beforeAlign,
      );
      // Duplicating a child must retain its inherited rotation without copying its sibling.
      await page.keyboard.press('ControlOrMeta+d');
      await saved();
      shapes = await state();
      assert.deepEqual(shapes.map(getShapeText).filter(Boolean), ['日本語', 'English', '日本語']);
      const copiedGroup = shapes.find(
        (s) => getShapeId(s) !== getShapeId(group) && getGroupChildren(s).length,
      );
      assert.ok(copiedGroup);
      assert.equal(getShapeRotation(copiedGroup), 90);
      assert.equal(getGroupChildren(copiedGroup).length, 1);
      await page.keyboard.press('ControlOrMeta+z');
      await saved();
      await editor.getByText('Editing group', { exact: true }).waitFor();
      await page.keyboard.press('Delete');
      await saved();
      assert.deepEqual((await state()).map(getShapeText).filter(Boolean), ['English']);
      await page.keyboard.press('ControlOrMeta+z');
      await saved();
      assert.deepEqual((await state()).map(getShapeText).filter(Boolean), ['日本語', 'English']);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await editor.getByText('グループを編集中', { exact: true }).waitFor();
      await editor.getByRole('button', { name: 'グループの編集を終了', exact: true }).click();
      assert.equal(await editor.locator('.hit').count(), 1);
      await page.reload();
      await saved();
      await editor.locator('.hit').dblclick();
      await editor.getByText('グループを編集中', { exact: true }).waitFor();
      assert.deepEqual((await state()).map(getShapeText).filter(Boolean), ['日本語', 'English']);
      for (const language of ['en', 'ja']) {
        await editor.locator('.lang select').selectOption(language);
        ja = language === 'ja';
        await editor.locator('.hit.selected').click();
        await page.keyboard.press('ControlOrMeta+a');
        assert.equal(await editor.locator('.hit.selected').count(), 2);
        const beforeGroup = (await state()).map((shape) => [
          getShapeId(shape),
          getShapeBounds(shape),
        ]);
        await page.keyboard.press('ControlOrMeta+g');
        await saved();
        shapes = await state();
        const outer = shapes.find((shape) => getShapeId(shape) === getShapeId(group));
        const inner = getGroupChildren(outer)[0];
        assert.equal(getGroupChildren(outer).length, 1);
        assert.deepEqual(getGroupChildren(inner).map(getShapeId), [id, getShapeId(second)]);
        assert.equal(getShapeRotation(outer), 90);
        assert.equal(await editor.locator('.hit').count(), 1);
        await page.keyboard.press('ControlOrMeta+Shift+g');
        await saved();
        assert.equal(await editor.locator('.hit.selected').count(), 2);
        assert.deepEqual(
          (await state()).map((shape) => [getShapeId(shape), getShapeBounds(shape)]),
          beforeGroup,
        );
        await page.keyboard.press('ControlOrMeta+z');
        await saved();
        assert.equal(await editor.locator('.hit.selected').count(), 1);
        await page.keyboard.press('ControlOrMeta+Shift+z');
        await saved();
        assert.equal(await editor.locator('.hit.selected').count(), 2);
        await page.reload();
        await saved();
        await editor.locator('.hit').dblclick();
        assert.equal(await editor.locator('.hit').count(), 2);
      }
      await page.screenshot({ path: '/tmp/pptx-group-edit-ja.png', fullPage: true });
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-group-edit-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

// Resetting a placeholder that lives inside a group.
//
// A group is an arrangement: it scales and turns what is inside it. It does
// not decide what font the text is in. So "reset to the layout" reaches a
// grouped placeholder's formatting and appearance, and leaves its place in the
// group alone — which is the one thing the layout has nothing to say about.
//
// Driven from the ribbon rather than the API, in both languages, because the
// point is that the command a person presses reaches these shapes at all.

import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  addSlide,
  addSlideTextBox,
  createPresentation,
  findSlideLayout,
  getGroupChildren,
  getShapeBoundsResolved,
  getShapeText,
  getShapeXmlString,
  getSlideShapes,
  getSlides,
  groupShapes,
  inches,
  loadPresentation,
  savePresentation,
  setShapeFill,
  setShapePosition,
  setShapeRotation,
  setShapeSize,
  setShapeText,
  setShapeTextFormat,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test('the ribbon resets a grouped placeholder without moving it', { timeout: 120000 }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-reset-grouped-'));
  let preview, browser, page;
  try {
    const pres = createPresentation();
    // Two slides: the first holds the group, the second is left alone so the
    // command's scope can be checked as well.
    for (const index of [0, 1]) {
      const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Title and Content') });
      for (const shape of getSlideShapes(slide)) {
        setShapeText(shape, `日本語 / English ${index}`);
        setShapeTextFormat(shape, { bold: true, size: 44, color: '#FF0000' });
        setShapeFill(shape, '#00FF00');
        const bounds = getShapeBoundsResolved(pres, shape);
        setShapePosition(shape, bounds.x, bounds.y);
        setShapeSize(shape, bounds.w, bounds.h);
      }
      addSlideTextBox(slide, {
        x: inches(1),
        y: inches(5),
        w: inches(2),
        h: inches(1),
        text: 'Decoration',
      });
      if (index === 0) {
        const group = groupShapes([...getSlideShapes(slide)]);
        setShapeRotation(group, 21);
      }
    }
    const placed = getSlides(pres).map((slide) =>
      getSlideShapes(slide).map((shape) => getShapeBoundsResolved(pres, shape)),
    );

    const source = join(dir, 'source.pptx');
    await writeFile(source, await savePresentation(pres));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';` +
        `export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
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
    await editor.getByText('Saved to this project', { exact: true }).waitFor();

    /** What the deck on disk says about every shape of every slide. */
    const state = async () => {
      const deck = await loadPresentation(
        new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
      );
      return getSlides(deck).map((slide) =>
        getSlideShapes(slide).map((shape) => {
          // A group's XML holds its children's, so asking it about their
          // formatting would answer for them. It has none of its own.
          const own = getGroupChildren(shape).length > 0 ? null : getShapeXmlString(shape);
          return {
            text: getShapeText(shape),
            bold: own === null ? null : own.includes('b="1"'),
            painted: own === null ? null : own.includes('00FF00'),
            bounds: getShapeBoundsResolved(deck, shape),
          };
        }),
      );
    };
    /** Waits for the deck itself to satisfy `done`, not for the save label. */
    const settles = async (done, what) => {
      const deadline = Date.now() + 20000;
      let last;
      for (;;) {
        last = await state();
        if (done(last)) return last;
        if (Date.now() > deadline) {
          assert.fail(`${what}: the deck never got there, last was ${JSON.stringify(last)}`);
        }
        await new Promise((resolve) => setTimeout(resolve, 100));
      }
    };
    const reset = (name) =>
      editor
        .getByRole('region', { name: ja ? 'スライドの設定' : 'Slide options', exact: true })
        .getByRole('button', { name, exact: true });

    const before = await state();
    // Slide one: a group holding a title, a body and a plain text box.
    assert.deepEqual(
      before[0].map((shape) => shape.text),
      ['', '日本語 / English 0', '日本語 / English 0', 'Decoration'],
    );
    assert.deepEqual(
      before[0].map((shape) => shape.bold),
      [null, true, true, false],
    );
    assert.deepEqual(
      before.map((slide) => slide.map((shape) => shape.bounds)),
      placed,
    );

    // --- Text formatting reaches inside the group --------------------------
    await editor.locator('.thumb-row').nth(0).click();
    await reset('Reset placeholder text formatting').click();
    const afterText = await settles(
      (now) => now[0].every((shape) => shape.bold !== true),
      'reset the grouped placeholders’ text formatting',
    );
    // The fills are a separate command, so they are still there.
    assert.deepEqual(
      afterText[0].map((shape) => shape.painted),
      [null, true, true, false],
    );
    // Nothing moved, on either slide, and the other slide is untouched.
    assert.deepEqual(
      afterText.map((slide) => slide.map((shape) => shape.bounds)),
      placed,
    );
    assert.deepEqual(afterText[1], before[1]);

    await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
    await settles(
      (now) => JSON.stringify(now) === JSON.stringify(before),
      'undo the formatting reset',
    );

    // --- And so does the whole layout reset, in Japanese -------------------
    await editor.locator('.lang select').selectOption('ja');
    ja = true;
    await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
    await reset('レイアウトをリセット').click();
    const afterLayout = await settles(
      (now) => now[0].every((shape) => shape.bold !== true && shape.painted !== true),
      'reset the whole layout',
    );
    // The text itself survived, and so did every shape's place in the group.
    assert.deepEqual(
      afterLayout[0].map((shape) => shape.text),
      ['', '日本語 / English 0', '日本語 / English 0', 'Decoration'],
    );
    assert.deepEqual(
      afterLayout.map((slide) => slide.map((shape) => shape.bounds)),
      placed,
    );
    assert.deepEqual(afterLayout[1], before[1]);

    // --- One step of history, and the file says the same after a reload ----
    await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
    await settles((now) => JSON.stringify(now) === JSON.stringify(before), 'undo the layout reset');
    await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
    await settles(
      (now) => JSON.stringify(now) === JSON.stringify(afterLayout),
      'redo the layout reset',
    );
    await page.reload();
    await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
    await page
      .frameLocator('#editor-frame')
      .getByText('このプロジェクトに保存済み', { exact: true })
      .waitFor();
    assert.deepEqual(await state(), afterLayout);

    assert.deepEqual(errors, []);
  } catch (error) {
    await page?.screenshot({ path: '/tmp/pptx-reset-grouped-failure.png', fullPage: true });
    throw error;
  } finally {
    await browser?.close();
    await preview?.close();
    await rm(dir, { recursive: true, force: true });
  }
});

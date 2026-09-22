// Two things the inline editor used to ignore while the renderer honoured
// them: `<a:bodyPr vert=…>` (the caret has to follow the reading direction of
// the painted glyphs) and `<a:normAutofit/>` (text shrunk to fit must not jump
// back to its authored size the moment the caret appears).
//
// Both are checked against the rendered SVG rather than against constants, so
// the editor is compared with the renderer, not with a copy of its rules.

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
  savePresentation,
  setShapeTextAutoFit,
  setShapeTextDirection,
  inches,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const LINES = Array.from({ length: 8 }, (_, i) => `Line ${i + 1}`).join('\n');

const openEditor = async (build, dir) => {
  const pres = createPresentation(),
    slide = addBlankSlide(pres);
  build(slide);
  const source = join(dir, 'source.pptx'),
    file = join(dir, 'deck.tsx');
  await writeFile(source, await savePresentation(pres));
  await writeFile(
    file,
    `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
  );
  const preview = await startPreview(file);
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
  await page.goto(preview.url);
  await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
  const editor = page.frameLocator('#editor-frame');
  await editor.getByText('Saved to this project', { exact: true }).waitFor();
  return { preview, browser, page, editor };
};

for (const { vert, writingMode } of [
  { vert: 'vert', writingMode: 'vertical-rl' },
  { vert: 'wordArtVert', writingMode: 'vertical-rl' },
  { vert: 'vert270', writingMode: 'vertical-lr' },
]) {
  test(
    `inline editing reads ${vert} the way the preview paints it`,
    { timeout: 60000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-vertical-text-'));
      let preview, browser;
      try {
        const opened = await openEditor((slide) => {
          const box = addSlideTextBox(slide, {
            x: inches(2),
            y: inches(1),
            w: inches(3),
            h: inches(4),
            text: '日本語 English',
          });
          setShapeTextDirection(box, vert);
        }, dir);
        ({ preview, browser } = opened);
        const { page, editor } = opened;
        await editor.locator('.hit').first().dblclick();
        const input = editor.locator('.inline-edit');
        await input.waitFor();
        const layout = await input.evaluate((node) => {
          const region = document.querySelector('.paint foreignObject div');
          return {
            editing: {
              writingMode: getComputedStyle(node).writingMode,
              orientation: getComputedStyle(node).textOrientation,
            },
            rendered: {
              writingMode: getComputedStyle(region).writingMode,
              orientation: getComputedStyle(region).textOrientation,
            },
            // The half turn vert270 reads bottom-to-top with lives on the box
            // transform in the editor and on the text container in the preview,
            // so compare the direction the glyphs actually run instead.
            editingBox: node.getBoundingClientRect().width > node.getBoundingClientRect().height,
          };
        });
        assert.equal(layout.editing.writingMode, writingMode, JSON.stringify(layout));
        assert.equal(
          layout.editing.writingMode,
          layout.rendered.writingMode,
          JSON.stringify(layout),
        );
        assert.equal(
          layout.editing.orientation,
          layout.rendered.orientation,
          JSON.stringify(layout),
        );
        assert.equal(layout.editingBox, false, 'a vertical box stays taller than it is wide');

        await input.fill('縦書き Edited');
        await page.keyboard.press('ControlOrMeta+Enter');
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        await editor.locator('.lang select').selectOption('ja');
        await page.reload();
        await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
        await editor.locator('.hit').first().dblclick();
        assert.equal(await input.innerText(), '縦書き Edited');
        assert.equal(
          await input.evaluate((node) => getComputedStyle(node).writingMode),
          writingMode,
          'the saved deck still reads the same way',
        );
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );
}

test(
  'inline editing shrinks autofit text by what the preview shrank it',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-autofit-text-'));
    let preview, browser;
    try {
      const opened = await openEditor((slide) => {
        // Eight lines in a 2×1 in box: far past what fits, so normAutofit has
        // something to do. The second box is the unshrunk reference.
        const shrunk = addSlideTextBox(slide, {
          x: inches(1),
          y: inches(1),
          w: inches(2),
          h: inches(1),
          text: LINES,
        });
        setShapeTextAutoFit(shrunk, 'normal');
        addSlideTextBox(slide, {
          x: inches(5),
          y: inches(1),
          w: inches(6),
          h: inches(4),
          text: LINES,
        });
      }, dir);
      ({ preview, browser } = opened);
      const { page, editor } = opened;
      const input = editor.locator('.inline-edit');
      // Editing px live at canvas zoom while the painted SVG's px live at slide
      // scale, so the two are compared as the ratio between the shrunk box and
      // the unshrunk reference beside it — which is the shrink factor itself.
      const editingFont = async (index) => {
        await editor.locator('.hit').nth(index).dblclick();
        await input.waitFor();
        return await input.evaluate((node) =>
          parseFloat(
            getComputedStyle(node.querySelector('[data-text-paragraph] span') ?? node).fontSize,
          ),
        );
      };
      const renderedFonts = () =>
        page.evaluate(() =>
          [
            ...document
              .querySelector('#editor-frame')
              .contentDocument.querySelectorAll('.paint foreignObject'),
          ].map((region) =>
            parseFloat(
              getComputedStyle(region.querySelector('p span') ?? region.querySelector('p'))
                .fontSize,
            ),
          ),
        );
      const shrunkEditing = await editingFont(0);
      const referenceEditing = await editingFont(1);
      await page.keyboard.press('Escape');
      const rendered = await renderedFonts();
      const paintedRatio = rendered[0] / rendered[1];
      // The renderer shrank the small box, and the editor by the same factor.
      assert.ok(paintedRatio < 0.9, JSON.stringify({ rendered }));
      assert.ok(
        Math.abs(shrunkEditing / referenceEditing - paintedRatio) < 0.02,
        JSON.stringify({ shrunkEditing, referenceEditing, rendered }),
      );

      await editor.locator('.hit').nth(0).dblclick();
      await input.waitFor();
      await input.fill('縮小 Edited');
      await page.keyboard.press('ControlOrMeta+Enter');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.lang select').selectOption('ja');
      await page.reload();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      const afterEditing = await editingFont(0);
      assert.equal(await input.innerText(), '縮小 Edited');
      await page.keyboard.press('Escape');
      // One short line fits, so the shrink is gone on both sides — the factor
      // follows the text rather than being frozen into the deck.
      const after = await renderedFonts();
      assert.ok(Math.abs(after[0] / after[1] - 1) < 0.001, JSON.stringify({ after }));
      assert.ok(afterEditing > shrunkEditing, JSON.stringify({ shrunkEditing, afterEditing }));
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

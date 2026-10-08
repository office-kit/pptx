// The format painter: pick up an object's formatting and put it on another,
// or repaint one text selection with the formatting of another — the gesture
// the reference desktop app and Google Slides both put on a paintbrush.
//
// Checked through the saved deck rather than the canvas, so a green run means
// the .pptx really carries the formatting.

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
  getParagraphPropertiesEffective,
  getShapeFill,
  getShapeRunFormatEffective,
  getShapeStroke,
  getShapeText,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
  savePresentation,
  setParagraphAlignment,
  setShapeFill,
  setShapeStroke,
  setShapeTextFormat,
} from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';

const WORDS = {
  en: (key) => key,
  ja: (key) =>
    ({
      'Saved to this project': 'このプロジェクトに保存済み',
      'Undo (Ctrl+Z)': '元に戻す (Ctrl+Z)',
    })[key],
};

const deckWith = async (dir) => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const source = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(3),
    h: inches(1),
    text: '書式もと Source',
  });
  setShapeFill(source, '#FF0000');
  setShapeStroke(source, { color: '#00FF00', widthEmu: 38100 });
  setShapeTextFormat(source, { bold: true, size: 30 });
  setParagraphAlignment(source, 0, 'right');
  addSlideTextBox(slide, {
    x: inches(6),
    y: inches(1),
    w: inches(3),
    h: inches(1),
    text: '書式さき Target',
  });
  const file = join(dir, 'deck.tsx');
  const source_pptx = join(dir, 'source.pptx');
  await writeFile(source_pptx, await savePresentation(pres));
  await writeFile(
    file,
    `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source_pptx)})} />;`,
  );
  return file;
};

const savedShapes = async (preview) => {
  const bytes = new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer());
  const pres = await loadPresentation(bytes);
  return { pres, shapes: getSlideShapes(getSlides(pres)[0]) };
};

for (const language of ['en', 'ja']) {
  test(
    `an object's formatting travels to another object (${language})`,
    { timeout: 60000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-format-painter-'));
      let preview, browser;
      const word = WORDS[language];
      try {
        preview = await startPreview(await deckWith(dir));
        browser = await chromium.launch({ headless: true });
        const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
        await page.goto(preview.url);
        const editor = page.frameLocator('#editor-frame');
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        if (language === 'ja') {
          await editor.locator('.lang select').selectOption('ja');
          await editor.getByText(word('Saved to this project'), { exact: true }).waitFor();
        }

        const before = await savedShapes(preview);

        // Pick the formatting up off the first object, put it on the second.
        await editor
          .locator('.hit')
          .first()
          .click({ position: { x: 2, y: 2 } });
        // Format ▸ Pick Up Object Style: ⇧⌘C in the English build, ⌥⌘C in the Japanese one.
        await page.keyboard.press(language === 'ja' ? 'Control+Alt+KeyC' : 'Control+Shift+KeyC');
        await editor
          .locator('.hit')
          .nth(1)
          .click({ position: { x: 2, y: 2 } });
        // Format ▸ Apply Object Style.
        await page.keyboard.press('Control+Shift+KeyV');
        await editor.getByText(word('Saved to this project'), { exact: true }).waitFor();

        const after = await savedShapes(preview);
        const target = after.shapes[1];
        assert.equal(getShapeText(target), '書式さき Target');
        assert.deepEqual(getShapeFill(target), { kind: 'solid', color: '#FF0000' });
        assert.deepEqual(getShapeStroke(target), {
          kind: 'solid',
          color: '#00FF00',
          widthEmu: 38100,
        });
        assert.equal(getShapeRunFormatEffective(after.pres, target, 0, 0).bold, true);
        assert.equal(getShapeRunFormatEffective(after.pres, target, 0, 0).size, 30);
        assert.equal(getParagraphPropertiesEffective(after.pres, target, 0).align, 'right');

        // One undo step puts the target back where it started.
        await editor.getByTitle(word('Undo (Ctrl+Z)'), { exact: true }).click();
        await editor.getByText(word('Saved to this project'), { exact: true }).waitFor();
        const undone = await savedShapes(preview);
        assert.deepEqual(getShapeFill(undone.shapes[1]), getShapeFill(before.shapes[1]));
        assert.deepEqual(
          getShapeRunFormatEffective(undone.pres, undone.shapes[1], 0, 0),
          getShapeRunFormatEffective(before.pres, before.shapes[1], 0, 0),
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
  'a text selection is repainted with another selection’s formatting',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-format-painter-text-'));
    let preview, browser;
    try {
      preview = await startPreview(await deckWith(dir));
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const before = await savedShapes(preview);

      // Copy from inside the source's text…
      await editor.locator('.hit').first().dblclick();
      const input = editor.locator('.canvas-shell .inline-edit');
      await input.waitFor();
      await input.press('ControlOrMeta+a');
      await input.press('ControlOrMeta+Shift+c');
      await page.keyboard.press('Escape');

      // …and paste it over the target's text.
      await editor.locator('.hit').nth(1).dblclick();
      await input.waitFor();
      await input.press('ControlOrMeta+a');
      const beforePasteRevision = (await waitForState(preview.url, () => true)).revision;
      await input.press('ControlOrMeta+Shift+v');
      await page.keyboard.press('ControlOrMeta+Enter');
      await waitForState(preview.url, (state) => state.revision !== beforePasteRevision);
      await editor.getByText('Saved to this project', { exact: true }).waitFor();

      const { pres, shapes } = await savedShapes(preview);
      assert.equal(getShapeRunFormatEffective(pres, shapes[1], 0, 0).bold, true);
      assert.equal(getShapeRunFormatEffective(pres, shapes[1], 0, 0).size, 30);
      assert.equal(getParagraphPropertiesEffective(pres, shapes[1], 0).align, 'right');
      // A text pickup carries no paint, so the target keeps its own fill.
      assert.deepEqual(getShapeFill(shapes[1]), getShapeFill(before.shapes[1]));
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'format painter shortcuts use the native selection before select delivery',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-format-painter-selection-race-'));
    let preview, browser;
    try {
      preview = await startPreview(await deckWith(dir));
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();

      const dispatchPainterKey = async (locator, code) =>
        locator.evaluate((node, keyCode) => {
          // Suppress the pending select event so the shortcut runs against a
          // deliberately stale reactive textRange while the native selection
          // is already current.
          document.addEventListener(
            'selectionchange',
            (event) => event.stopImmediatePropagation(),
            {
              capture: true,
              once: true,
            },
          );
          const range = document.createRange();
          range.selectNodeContents(node);
          const selection = window.getSelection();
          selection.removeAllRanges();
          selection.addRange(range);
          node.dispatchEvent(
            new KeyboardEvent('keydown', {
              bubbles: true,
              cancelable: true,
              // ⇧⌘C / ⇧⌘V: Pick Up and Apply Object Style in the reference desktop app's English build.
              ctrlKey: true,
              shiftKey: true,
              key: keyCode.slice(-1),
              code: keyCode,
            }),
          );
        }, code);

      await editor.locator('.hit').first().dblclick();
      const input = editor.locator('.canvas-shell .inline-edit');
      await input.waitFor();
      await dispatchPainterKey(input, 'KeyC');
      await page.keyboard.press('Escape');

      await editor.locator('.hit').nth(1).dblclick();
      await input.waitFor();
      const beforePasteRevision = (await waitForState(preview.url, () => true)).revision;
      await dispatchPainterKey(input, 'KeyV');
      await page.keyboard.press('ControlOrMeta+Enter');
      await waitForState(preview.url, (state) => state.revision !== beforePasteRevision);
      const { pres, shapes } = await savedShapes(preview);
      assert.equal(getShapeRunFormatEffective(pres, shapes[1], 0, 0).bold, true);
      assert.equal(getShapeRunFormatEffective(pres, shapes[1], 0, 0).size, 30);
      assert.equal(getParagraphPropertiesEffective(pres, shapes[1], 0).align, 'right');
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

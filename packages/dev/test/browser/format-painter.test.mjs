// The format painter: pick up an object's formatting and put it on another,
// or repaint one text selection with the formatting of another — the gesture
// PowerPoint and Google Slides both put on a paintbrush.
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
import { startPreview } from '../helpers/server.mjs';

const WORDS = {
  en: (key) => key,
  ja: (key) =>
    ({
      'Saved to this project': 'このプロジェクトに保存済み',
      'Copy formatting': '書式のコピー',
      'Paste formatting': '書式の貼り付け',
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
        await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
        const editor = page.frameLocator('#editor-frame');
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        if (language === 'ja') {
          await editor.locator('.lang select').selectOption('ja');
          await editor.getByText(word('Saved to this project'), { exact: true }).waitFor();
        }

        const before = await savedShapes(preview);

        // Pick the formatting up off the first object, put it on the second.
        await editor.locator('.hit').first().click();
        await editor.locator('.hit').first().click({ button: 'right' });
        await editor.getByRole('menuitem', { name: word('Copy formatting') }).click();
        await editor.locator('.hit').nth(1).click();
        await editor.locator('.hit').nth(1).click({ button: 'right' });
        await editor.getByRole('menuitem', { name: word('Paste formatting') }).click();
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
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const before = await savedShapes(preview);

      // Copy from inside the source's text…
      await editor.locator('.hit').first().dblclick();
      const input = editor.locator('.inline-edit');
      await input.waitFor();
      await page.keyboard.press('ControlOrMeta+a');
      await editor
        .locator('.canvas-shell > .text-format-bar')
        .getByRole('button', { name: 'Copy formatting', exact: true })
        .click();
      await page.keyboard.press('Escape');

      // …and paste it over the target's text.
      await editor.locator('.hit').nth(1).dblclick();
      await input.waitFor();
      await page.keyboard.press('ControlOrMeta+a');
      await editor
        .locator('.canvas-shell > .text-format-bar')
        .getByRole('button', { name: 'Paste formatting', exact: true })
        .click();
      await page.keyboard.press('ControlOrMeta+Enter');
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

// Deck-wide slide numbers, the way Google Slides presents them: one switch
// that numbers every slide, and a number that follows the slide's position
// rather than repeating whatever text was typed once.
//
// Checked through the saved deck (the field is really in the .pptx) and
// through the canvas (the live number the author sees while editing).

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
  getShapeParagraphElements,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
  savePresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const WORDS = {
  en: (key) => key,
  ja: (key) =>
    ({
      'Saved to this project': 'このプロジェクトに保存済み',
      'Slide numbers': 'スライド番号',
      'Undo (Ctrl+Z)': '元に戻す (Ctrl+Z)',
    })[key],
};

const deckWith = async (dir) => {
  const pres = createPresentation();
  for (const text of ['一枚目 First', '二枚目 Second']) {
    const slide = addBlankSlide(pres);
    addSlideTextBox(slide, { x: inches(1), y: inches(1), w: inches(4), h: inches(1), text });
  }
  const file = join(dir, 'deck.tsx');
  const source = join(dir, 'source.pptx');
  await writeFile(source, await savePresentation(pres));
  await writeFile(
    file,
    `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
  );
  return file;
};

// Every slide's slide-number fields, as the saved .pptx carries them.
const savedFields = async (preview) => {
  const bytes = new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer());
  const pres = await loadPresentation(bytes);
  return getSlides(pres).map((slide) =>
    getSlideShapes(slide).flatMap((shape) =>
      getShapeParagraphElements(shape, 0).filter(
        (element) => element.kind === 'fld' && element.type === 'slidenum',
      ),
    ),
  );
};

for (const language of ['en', 'ja']) {
  test(`one switch numbers every slide (${language})`, { timeout: 60000 }, async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-slide-numbers-'));
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

      assert.deepEqual(await savedFields(preview), [[], []]);

      const toggle = editor.getByLabel(word('Slide numbers'), { exact: true });
      await toggle.check();
      await editor.getByText(word('Saved to this project'), { exact: true }).waitFor();

      const fields = await savedFields(preview);
      assert.equal(fields.length, 2);
      for (const slide of fields) assert.equal(slide.length, 1);

      // The canvas shows the slide's own number, not the cached text — the
      // second slide reads "2" although nothing ever typed a 2 into it.
      const canvas = editor.locator('.paint');
      const first = await canvas.innerHTML();
      assert.match(first, />1</);
      assert.doesNotMatch(first, />2</);
      await editor.locator('.thumb-row').nth(1).click();
      await editor.locator('.thumb-row[aria-current="true"]').nth(0).waitFor();
      const second = await canvas.innerHTML();
      assert.match(second, />2</);
      assert.doesNotMatch(second, />1</);

      // Off again removes the numbers and leaves the slides' own text.
      await toggle.uncheck();
      await editor.getByText(word('Saved to this project'), { exact: true }).waitFor();
      assert.deepEqual(await savedFields(preview), [[], []]);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  });
}

// The other field types have no live value a preview could compute, so they
// stay per-shape: Insert ▸ Insert field turns the selected box into one.
test('the ribbon turns a selected box into a date field', { timeout: 60000 }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-insert-field-'));
  let preview, browser;
  try {
    preview = await startPreview(await deckWith(dir));
    browser = await chromium.launch({ headless: true });
    const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
    await page.goto(preview.url);
    await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
    const editor = page.frameLocator('#editor-frame');
    await editor.getByText('Saved to this project', { exact: true }).waitFor();

    await editor.locator('.hit').first().click();
    await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
    await editor.locator('.ribbon').getByRole('button', { name: 'Insert field' }).click();
    const dialog = editor.getByRole('dialog');
    await dialog.locator('select').selectOption('datetime1');
    await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
    await editor.getByText('Saved to this project', { exact: true }).waitFor();

    const bytes = new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer());
    const pres = await loadPresentation(bytes);
    const shape = getSlideShapes(getSlides(pres)[0])[0];
    assert.deepEqual(
      getShapeParagraphElements(shape, 0).map((element) => [element.kind, element.type]),
      [['fld', 'datetime1']],
    );
  } finally {
    await browser?.close();
    await preview?.close();
    await rm(dir, { recursive: true, force: true });
  }
});

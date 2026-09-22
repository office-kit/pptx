// A run's own outline, shadow and glow: painted on the editing canvas, and
// carried into the saved .pptx by the format painter.

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
  getShapeRunFormatEffective,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
  savePresentation,
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
    })[key],
};

const deckWith = async (dir) => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const source = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(1.2),
    text: '袋文字 WordArt',
  });
  setShapeTextFormat(source, {
    size: 40,
    color: '#FFFFFF',
    outline: { color: '#FF0000', widthEmu: 19050 },
    shadow: { color: '#000000', blurEmu: 50800, offsetEmu: 38100, angleDeg: 45 },
    glow: { color: '#00FF00', radiusEmu: 63500 },
  });
  addSlideTextBox(slide, {
    x: inches(1),
    y: inches(3),
    w: inches(4),
    h: inches(1.2),
    text: '素の文字 Plain',
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
    `character effects paint on the canvas and travel with the format painter (${language})`,
    { timeout: 60000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-character-effects-'));
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

        // The canvas paints what the file states, before any edit.
        const painted = await editor.locator('.paint').innerHTML();
        assert.match(painted, /-webkit-text-stroke:2\.00px #FF0000/);
        assert.match(painted, /text-shadow:0 0 6\.67px #00FF00/);

        // The format painter carries them onto the plain box.
        await editor.locator('.hit').first().click();
        await editor.locator('.hit').first().click({ button: 'right' });
        await editor.getByRole('menuitem', { name: word('Copy formatting') }).click();
        await editor.locator('.hit').nth(1).click();
        await editor.locator('.hit').nth(1).click({ button: 'right' });
        await editor.getByRole('menuitem', { name: word('Paste formatting') }).click();
        await editor.getByText(word('Saved to this project'), { exact: true }).waitFor();

        const { pres, shapes } = await savedShapes(preview);
        const format = getShapeRunFormatEffective(pres, shapes[1], 0, 0);
        assert.deepEqual(format.outline, { color: '#FF0000', widthEmu: 19050 });
        assert.equal(format.glow?.color, '#00FF00');
        assert.equal(format.shadow?.offsetEmu, 38100);
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );
}

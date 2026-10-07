import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeRunFormat,
  getShapeStroke,
  getShapeStrokeGradient,
  getShapeStrokeSketch,
  getShapeTextFlat,
  getSlides,
  getSlideShapes,
  loadPresentation,
} from '@office-kit/pptx';
import { expandFormatSections } from '../helpers/format-pane.mjs';
import { startPreview } from '../helpers/server.mjs';

// Mac PowerPoint's Format Shape pane: Gradient line and Sketched style in
// Line, No fill / Picture or texture fill and Gradient line under Text Fill &
// Outline, and Keep text flat in 3-D Rotation.
const DECK = `import {Presentation,Slide,Shape} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Shape preset="rect" x={1} y={1} width={3} height={1.5} fill="#2E75B6" text="First" /></Slide></Presentation>`;

const open = async (lang) => {
  const dir = await mkdtemp(join(tmpdir(), 'office-format-line-paint-'));
  const file = join(dir, 'deck.tsx');
  await writeFile(file, DECK);
  const preview = await startPreview(file);
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({ viewport: { width: 1512, height: 1000 } });
  const errors = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto(preview.url);
  const editor = page.frameLocator('#editor-frame');
  const saved = () =>
    editor
      .getByText(lang === 'ja' ? 'このプロジェクトに保存済み' : 'Saved to this project', {
        exact: true,
      })
      .waitFor();
  await editor.getByText('Saved to this project', { exact: true }).waitFor();
  if (lang === 'ja') await editor.locator('.lang select').selectOption('ja');
  const shape = async () => {
    const deck = await loadPresentation(
      new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
    );
    return getSlideShapes(getSlides(deck)[0])[0];
  };
  await editor
    .locator('.hit')
    .first()
    .click({ button: 'right', position: { x: 4, y: 4 } });
  await editor
    .getByRole('menuitem', {
      name: lang === 'ja' ? '図形の書式設定...' : 'Format Shape...',
      exact: true,
    })
    .click();
  const close = async () => {
    await browser.close();
    await preview.close();
    await rm(dir, { recursive: true, force: true });
  };
  return { editor, saved, shape, errors, close };
};

test(
  'Line: Gradient line edits the outline gradient and Sketched style writes the sketch',
  { timeout: 120000 },
  async () => {
    const { editor, saved, shape, errors, close } = await open('en');
    try {
      await expandFormatSections(editor, 'Line');
      const lineTypes = editor.getByRole('group', { name: 'Line type', exact: true });
      await lineTypes.getByRole('radio', { name: 'Gradient line', exact: true }).check();
      await saved();
      assert.equal(getShapeStroke(await shape()).kind, 'gradient');
      assert.equal(getShapeStrokeGradient(await shape()).stops.length, 4);
      // The fill's gradient controls drive the line.
      const section = editor.locator('details[data-section="line"]');
      const angle = section.getByRole('spinbutton', { name: 'Gradient angle', exact: true });
      await angle.fill('30');
      await angle.press('Enter');
      await saved();
      assert.equal(getShapeStrokeGradient(await shape()).angleDeg, 30);
      await section.getByRole('button', { name: 'Remove gradient stop', exact: true }).click();
      await saved();
      assert.equal(getShapeStrokeGradient(await shape()).stops.length, 3);
      assert.equal(await section.getByRole('combobox', { name: 'Compound type' }).count(), 1);

      // Sketched style: None, Curved, Freehand, Scribble.
      const sketch = editor.getByRole('button', { name: 'Sketched style', exact: true });
      await sketch.click();
      const menu = editor.getByRole('menu', { name: 'Sketched style', exact: true });
      assert.deepEqual(
        await menu
          .getByRole('menuitemradio')
          .evaluateAll((items) =>
            items.map((item) => [
              item.getAttribute('aria-label'),
              item.getAttribute('aria-checked'),
            ]),
          ),
        [
          ['None', 'true'],
          ['Curved', 'false'],
          ['Freehand', 'false'],
          ['Scribble', 'false'],
        ],
      );
      await menu.getByRole('menuitemradio', { name: 'Freehand', exact: true }).click();
      await saved();
      assert.equal(getShapeStrokeSketch(await shape()), 'freehand');
      // The gradient survives the sketch.
      assert.equal(getShapeStroke(await shape()).kind, 'gradient');
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(getShapeStrokeSketch(await shape()), null);

      await lineTypes.getByRole('radio', { name: 'Solid line', exact: true }).check();
      await saved();
      assert.equal(getShapeStroke(await shape()).kind, 'solid');
      assert.equal(await section.getByRole('spinbutton', { name: 'Gradient angle' }).count(), 0);
      assert.deepEqual(errors, []);
    } finally {
      await close();
    }
  },
);

test(
  'Text Options: No fill, Picture fill, Gradient line and Keep text flat',
  { timeout: 120000 },
  async () => {
    const { editor, saved, shape, errors, close } = await open('en');
    try {
      await editor.getByRole('radio', { name: 'Text Options' }).click();
      await expandFormatSections(editor, 'Text Fill', 'Text Outline');
      const fills = editor.getByRole('group', { name: 'Text fill type', exact: true });
      await fills.getByRole('radio', { name: 'No fill', exact: true }).check();
      await saved();
      assert.deepEqual((await getShapeRunFormat(await shape(), 0, 0)).textFill, { kind: 'none' });
      await fills.getByRole('radio', { name: 'Picture or texture fill', exact: true }).check();
      // The texture PNG is encoded asynchronously, so the save badge from the
      // previous edit can still be showing; poll the saved deck instead.
      let picture;
      for (const deadline = Date.now() + 10000; Date.now() < deadline; ) {
        await saved();
        picture = (await getShapeRunFormat(await shape(), 0, 0)).textFill;
        if (picture?.kind === 'image') break;
        await new Promise((resolve) => setTimeout(resolve, 100));
      }
      assert.equal(picture.kind, 'image');
      // The default texture, a PNG.
      assert.deepEqual(Array.from(picture.bytes.subarray(1, 4)), [0x50, 0x4e, 0x47]);
      assert.equal(
        await editor
          .locator('details[data-section="textFill"]')
          .getByRole('button', { name: 'Insert...', exact: true })
          .count(),
        1,
      );

      const outlines = editor.getByRole('group', { name: 'Text outline type', exact: true });
      await outlines.getByRole('radio', { name: 'Gradient line', exact: true }).check();
      await saved();
      let outline = (await getShapeRunFormat(await shape(), 0, 0)).outline;
      assert.equal(outline.fill.kind, 'gradient');
      assert.equal(outline.widthEmu, 9525);
      const angle = editor
        .locator('details[data-section="textOutline"]')
        .getByRole('spinbutton', { name: 'Gradient angle', exact: true });
      await angle.fill('45');
      await angle.press('Enter');
      await saved();
      outline = (await getShapeRunFormat(await shape(), 0, 0)).outline;
      assert.equal(outline.fill.angleDeg, 45);
      assert.equal(outline.widthEmu, 9525);

      // Keep text flat (Text Effects ▸ 3-D Rotation); Soft Edges stays off for
      // text, as in PowerPoint.
      await editor.getByRole('tab', { name: 'Text Effects', exact: true }).click();
      await expandFormatSections(editor, 'Soft Edges', '3-D Rotation');
      assert.equal(
        await editor
          .locator('details[data-section="text-softEdges"]')
          .getByRole('button', { name: 'Presets', exact: true })
          .isDisabled(),
        true,
      );
      await editor
        .locator('details[data-section="text-3dRotation"]')
        .getByRole('checkbox', { name: 'Keep text flat', exact: true })
        .check();
      await saved();
      assert.equal(getShapeTextFlat(await shape()), true);
      assert.deepEqual(errors, []);
    } finally {
      await close();
    }
  },
);

test(
  'Japanese: gradient line, sketched style and text paint use PowerPoint wording',
  { timeout: 120000 },
  async () => {
    const { editor, saved, shape, errors, close } = await open('ja');
    try {
      await expandFormatSections(editor, '線');
      await editor
        .getByRole('group', { name: '線の種類', exact: true })
        .getByRole('radio', { name: '線 (グラデーション)', exact: true })
        .check();
      await saved();
      assert.equal(getShapeStroke(await shape()).kind, 'gradient');
      await editor.getByRole('button', { name: 'スケッチ スタイル', exact: true }).click();
      const items = editor
        .getByRole('menu', { name: 'スケッチ スタイル', exact: true })
        .getByRole('menuitemradio');
      // Mac PowerPoint's Japanese build names Freehand and Scribble alike.
      assert.deepEqual(
        await items.evaluateAll((nodes) => nodes.map((node) => node.getAttribute('aria-label'))),
        ['なし', '曲線', 'フリーハンド', 'フリーハンド'],
      );
      await items.nth(3).click();
      await saved();
      assert.equal(getShapeStrokeSketch(await shape()), 'scribble');

      await editor.getByRole('radio', { name: '文字のオプション' }).click();
      await expandFormatSections(editor, '文字の塗りつぶし');
      await editor
        .getByRole('group', { name: '文字の塗りつぶしの種類', exact: true })
        .getByRole('radio', { name: '塗りつぶしなし', exact: true })
        .check();
      await saved();
      assert.deepEqual((await getShapeRunFormat(await shape(), 0, 0)).textFill, { kind: 'none' });
      await editor.getByRole('tab', { name: '文字の効果', exact: true }).click();
      await expandFormatSections(editor, '3-D 回転');
      await editor.getByRole('checkbox', { name: 'テキストを立体表示しない', exact: true }).check();
      await saved();
      assert.equal(getShapeTextFlat(await shape()), true);
      assert.deepEqual(errors, []);
    } finally {
      await close();
    }
  },
);

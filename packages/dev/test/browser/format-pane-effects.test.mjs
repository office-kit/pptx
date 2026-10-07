import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShape3D,
  getShapeEffects,
  getShapeRunFormat,
  getSlides,
  getSlideShapes,
  loadPresentation,
} from '@office-kit/pptx';
import { expandFormatSections } from '../helpers/format-pane.mjs';
import { startPreview } from '../helpers/server.mjs';

// Mac PowerPoint's Format Shape pane, as captured on 2026-10-07: a Shape
// Options / Text Options switch, every section collapsed on first open, and
// the Effects category's Shadow, Reflection, Glow, Soft Edges, 3-D Format and
// 3-D Rotation sections.
const DECK = `import {Presentation,Slide,Shape} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Shape preset="rect" x={1} y={1} width={3} height={1.5} fill="#2E75B6" text="First" /><Shape preset="rect" x={5} y={1} width={3} height={1.5} fill="#2E75B6" text="Second" /></Slide></Presentation>`;
const EFFECTS = ['Shadow', 'Reflection', 'Glow', 'Soft Edges', '3-D Format', '3-D Rotation'];

const sectionState = (editor) =>
  editor
    .locator('details.pane-section:visible')
    .evaluateAll((nodes) =>
      nodes.map((node) => [node.querySelector(':scope > summary').textContent.trim(), node.open]),
    );

test(
  'Format Shape pane: options switch, collapsed sections and the native Effects sections',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-format-effects-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(file, DECK);
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1512, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const shapes = async () => {
        const deck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getSlideShapes(getSlides(deck)[0]);
      };
      const effects = async (index = 0) => {
        const deck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getShapeEffects(deck, getSlideShapes(getSlides(deck)[0])[index]);
      };
      const undo = async () => {
        await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
        await saved();
      };
      const section = (id) => editor.locator(`details[data-section="${id}"]`);
      const preset = async (id, button, name) => {
        await section(id).getByRole('button', { name: button, exact: true }).click();
        await editor.getByRole('menuitemradio', { name, exact: true }).click();
        await saved();
      };
      const spin = async (name, value) => {
        const input = editor.getByRole('spinbutton', { name, exact: true });
        await input.fill(String(value));
        await input.press('Enter');
        await saved();
      };
      await saved();
      await editor
        .locator('.hit')
        .first()
        .click({ button: 'right', position: { x: 4, y: 4 } });
      await editor.getByRole('menuitem', { name: 'Format Shape...', exact: true }).click();

      // Shape Options / Text Options above the categories; everything collapsed.
      const switcher = editor.getByRole('radiogroup', { name: 'Format Shape options' });
      assert.deepEqual(
        await switcher
          .getByRole('radio')
          .evaluateAll((nodes) =>
            nodes.map((node) => [node.textContent.trim(), node.getAttribute('aria-checked')]),
          ),
        [
          ['Shape Options', 'true'],
          ['Text Options', 'false'],
        ],
      );
      assert.deepEqual(await sectionState(editor), [
        ['Fill', false],
        ['Line', false],
      ]);
      await editor.getByRole('tab', { name: 'Effects', exact: true }).click();
      assert.deepEqual(
        await sectionState(editor),
        EFFECTS.map((name) => [name, false]),
      );

      // Shadow: a preset, then single fields, each one undo step.
      await expandFormatSections(editor, 'Shadow');
      await preset('shadow', 'Presets', 'Offset: Bottom Right');
      assert.deepEqual(
        (await effects()).map(({ kind, blurEmu, distEmu, angleDeg, opacity }) => ({
          kind,
          blurEmu,
          distEmu,
          angleDeg,
          opacity,
        })),
        [{ kind: 'outerShdw', blurEmu: 50800, distEmu: 38100, angleDeg: 45, opacity: 0.4 }],
      );
      // Mac PowerPoint shows this preset as 60%, 100%, 4 pt, 45°, 3 pt.
      for (const [name, value] of [
        ['Shadow transparency', '60'],
        ['Shadow size', '100'],
        ['Shadow blur', '4'],
        ['Shadow angle', '45'],
        ['Shadow distance', '3'],
      ])
        assert.equal(
          await editor.getByRole('spinbutton', { name, exact: true }).inputValue(),
          value,
        );
      await spin('Shadow distance', 6);
      assert.equal((await effects())[0].distEmu, 76200);
      await undo();
      assert.equal((await effects())[0].distEmu, 38100);
      await preset('shadow', 'Presets', 'Inside: Center');
      assert.deepEqual(
        (await effects()).map((effect) => effect.kind),
        ['innerShdw'],
      );
      assert.equal(
        await editor.getByRole('spinbutton', { name: 'Shadow size', exact: true }).isDisabled(),
        true,
      );
      await preset('shadow', 'Presets', 'None');
      assert.deepEqual(await effects(), []);

      // Reflection, Glow and Soft Edges.
      await expandFormatSections(editor, 'Reflection', 'Glow', 'Soft Edges');
      await preset('reflection', 'Presets', 'Half Reflection: 4 point offset');
      let [reflection] = await effects();
      assert.equal(reflection.kind, 'reflection');
      assert.equal(reflection.endPosition, 0.55);
      assert.equal(reflection.distEmu, 50800);
      await spin('Reflection size', 40);
      [reflection] = await effects();
      assert.equal(reflection.endPosition, 0.4);
      await preset('glow', 'Presets', 'Glow: 8 point; Accent color 2');
      await spin('Glow transparency', 25);
      const glow = (await effects()).find((effect) => effect.kind === 'glow');
      assert.equal(glow.radiusEmu, 101600);
      assert.equal(glow.opacity, 0.75);
      await preset('softEdges', 'Presets', '5 Point');
      await spin('Soft edge size', 10);
      assert.deepEqual(
        (await effects()).map((effect) => [effect.kind, effect.radiusEmu ?? null]),
        [
          ['glow', 101600],
          ['reflection', null],
          ['softEdge', 127000],
        ],
      );

      // 3-D Format and 3-D Rotation write the shape's own scene3d / sp3d.
      await expandFormatSections(editor, '3-D Format', '3-D Rotation');
      await preset('3dFormat', 'Top bevel', 'Circle');
      await spin('Depth size', 12);
      await preset('3dFormat', 'Material', 'Metal');
      await preset('3dRotation', 'Presets', 'Perspective Front');
      await spin('X Rotation', 30);
      await spin('Perspective', 60);
      const threeD = getShape3D((await shapes())[0]);
      assert.deepEqual(threeD.bevelTop, { widthEmu: 76200, heightEmu: 76200, preset: 'circle' });
      assert.equal(threeD.extrusionHeightEmu, 152400);
      assert.equal(threeD.material, 'metal');
      assert.equal(threeD.scene.camera, 'perspectiveFront');
      assert.deepEqual(threeD.scene.cameraRotation, {
        latitudeDeg: 0,
        longitudeDeg: 30,
        revolutionDeg: 0,
      });
      assert.equal(threeD.scene.fieldOfViewDeg, 60);
      await section('3dRotation')
        .getByRole('button', { name: 'Reset 3-D Rotation', exact: true })
        .click();
      await saved();
      assert.equal(getShape3D((await shapes())[0]).scene.camera, 'orthographicFront');

      // A multiple selection changes in one transaction.
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      await preset('glow', 'Presets', 'Glow: 18 point; Accent color 6');
      assert.deepEqual(
        [await effects(0), await effects(1)].map(
          (list) => list.find((effect) => effect.kind === 'glow')?.radiusEmu,
        ),
        [228600, 228600],
      );
      await undo();
      assert.deepEqual(
        [await effects(0), await effects(1)].map(
          (list) => list.find((effect) => effect.kind === 'glow')?.radiusEmu,
        ),
        [101600, undefined],
      );

      // Text Options: categories and sections for the text, collapsed.
      await editor.locator('.hit').nth(1).click();
      await switcher.getByRole('radio', { name: 'Text Options' }).click();
      assert.deepEqual(
        await editor
          .getByRole('tablist', { name: 'Format Shape' })
          .getByRole('tab')
          .evaluateAll((tabs) => tabs.map((tab) => tab.getAttribute('aria-label'))),
        ['Text Fill & Outline', 'Text Effects', 'Textbox'],
      );
      assert.deepEqual(await sectionState(editor), [
        ['Text Fill', false],
        ['Text Outline', false],
      ]);
      await expandFormatSections(editor, 'Text Outline');
      await editor
        .getByRole('group', { name: 'Text outline type' })
        .getByRole('radio', { name: 'Solid line' })
        .check();
      await saved();
      await spin('Text outline width', 2);
      const runFormat = () => shapes().then((list) => getShapeRunFormat(list[1], 0, 0));
      assert.equal((await runFormat()).outline.widthEmu, 25400);
      await editor.getByRole('tab', { name: 'Text Effects', exact: true }).click();
      assert.deepEqual(
        await sectionState(editor),
        EFFECTS.map((name) => [name, false]),
      );
      await expandFormatSections(editor, 'Shadow', 'Soft Edges');
      await preset('text-shadow', 'Presets', 'Offset: Bottom');
      assert.equal((await runFormat()).shadow.angleDeg, 90);
      assert.deepEqual(
        (await effects(1)).map((effect) => effect.kind),
        [],
      );
      assert.equal(
        await section('text-softEdges')
          .getByRole('button', { name: 'Presets', exact: true })
          .isDisabled(),
        true,
      );
      await editor.getByRole('tab', { name: 'Textbox', exact: true }).click();
      assert.deepEqual(await sectionState(editor), [['Text Box', false]]);

      // Japanese uses PowerPoint's wording; the switch keeps its state.
      await editor.locator('.lang select').selectOption('ja');
      assert.equal(
        await editor.getByRole('radio', { name: '文字のオプション' }).getAttribute('aria-checked'),
        'true',
      );
      await editor.getByRole('radio', { name: '図形のオプション' }).click();
      await editor.getByRole('tab', { name: '効果', exact: true }).click();
      assert.deepEqual(
        (await sectionState(editor)).map(([name]) => name),
        ['影', '反射', '光彩', 'ぼかし', '3-D 書式', '3-D 回転'],
      );
      assert.equal(
        await section('shadow').getByRole('button', { name: '標準スタイル', exact: true }).count(),
        1,
      );
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'Line section: Sketched style, 39 pt arrow buttons and no flip checkboxes',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-format-line-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(file, DECK);
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1512, height: 1000 } });
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor
        .locator('.hit')
        .first()
        .click({ button: 'right', position: { x: 4, y: 4 } });
      await editor.getByRole('menuitem', { name: 'Format Shape...', exact: true }).click();
      await expandFormatSections(editor, 'Line');
      await editor
        .getByRole('group', { name: 'Line type' })
        .getByRole('radio', { name: 'Solid line' })
        .check();
      const sketch = editor.getByRole('button', { name: 'Sketched style', exact: true });
      assert.equal(await sketch.isDisabled(), true);
      const labels = await editor
        .locator('[data-section="line"] .pane-fields')
        .evaluate((node) =>
          [...node.querySelectorAll('button[aria-haspopup], select')].map((control) =>
            control.getAttribute('aria-label'),
          ),
        );
      assert.deepEqual(labels.slice(labels.indexOf('Sketched style')), [
        'Sketched style',
        'Compound type',
        'Outline style',
        'Cap type',
        'Join type',
        'Begin Arrow type',
        'Begin Arrow size',
        'End Arrow type',
        'End Arrow size',
      ]);
      for (const name of ['Sketched style', 'Begin Arrow type', 'End Arrow size']) {
        const rect = await editor
          .getByRole('button', { name, exact: true })
          .evaluate((node) => node.getBoundingClientRect().toJSON());
        assert.deepEqual([rect.width, rect.height], [39, 26]);
      }
      await editor.getByRole('tab', { name: 'Size & Properties', exact: true }).click();
      await expandFormatSections(editor, 'Size', 'Position');
      assert.equal(await editor.getByRole('checkbox', { name: /^Flip/ }).count(), 0);
      await editor.locator('.lang select').selectOption('ja');
      await editor.getByRole('tab', { name: '塗りつぶしと線', exact: true }).click();
      assert.equal(
        await editor.getByRole('button', { name: 'スケッチ スタイル', exact: true }).count(),
        1,
      );
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

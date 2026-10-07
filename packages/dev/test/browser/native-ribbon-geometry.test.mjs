import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { startPreview } from '../helpers/server.mjs';

// Insert, Draw and Design measured from Mac PowerPoint 16 through the
// accessibility API in 1512 × 900 and 1200 × 900 pt windows (2026-10-07).
// CSS px equal Mac points; positions are window-relative.
const DECK = `import {Presentation,Slide,Shape} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Shape preset="rect" x={4} y={4} width={3} height={1.5} fill="#2E75B6" /></Slide></Presentation>`;

const INSERT = [
  'New Slide',
  'Table',
  'Pictures',
  'Screenshot',
  'Cameo',
  'Shapes',
  'Icons',
  '3D Models',
  'SmartArt',
  'Chart',
  'Zoom',
  'Link',
  'Action',
  'Comment',
  'Text Box',
  'Header & Footer',
  'WordArt',
  'Date & Time',
  'Slide Number',
  'Object',
  'Equation',
  'Symbol',
  'Video',
  'Audio',
];
// Native x of each Insert command at 1512 pt.
const INSERT_X = [
  9, 78, 147, 197, 280, 340, 390, 428, 478, 530, 599, 649, 687, 744, 818, 868, 922, 972, 1012, 1059,
  1117, 1168, 1230, 1280,
];
// Native grouping: Slides | Table | Pictures, Screenshot | Cameo | Shapes … Chart
// | Zoom, Link, Action | Comment | Text Box … Object | Equation, Symbol | Video, Audio.
const INSERT_GROUPS = [1, 1, 2, 1, 5, 3, 1, 6, 2, 2];
const TOLERANCE = 6;

const boxes = (locator) =>
  locator.evaluateAll((nodes) =>
    nodes.map((node) => {
      const rect = node.getBoundingClientRect();
      return {
        name: node.getAttribute('aria-label'),
        x: rect.x,
        y: rect.y,
        width: rect.width,
        height: rect.height,
      };
    }),
  );
const near = (actual, expected, label) =>
  assert.ok(Math.abs(actual - expected) <= TOLERANCE, `${label}: ${actual} (native ${expected})`);

test(
  'Insert, Draw and Design ribbons follow Mac PowerPoint geometry in English and Japanese',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-ribbon-geometry-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(file, DECK);
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1512, height: 900 } });
      await page.goto(preview.url + '/editor');
      await page.getByText('Saved to this project', { exact: true }).waitFor();
      const commands = (tab) =>
        page.locator(`.${tab} section button:not(.side, .pager, .theme, .pen)`);

      // Insert at 1512 pt: native order and groups, every command large,
      // 38 pt plain and 50 pt ▾ buttons filling the 72 pt row.
      await page.locator('#ribbon-tab-insert').click();
      assert.equal((await page.locator('.insert').boundingBox()).height, 72);
      let insert = await boxes(commands('insert'));
      assert.deepEqual(
        insert.map(({ name }) => name),
        INSERT,
      );
      assert.deepEqual(
        await page
          .locator('.insert > section')
          .evaluateAll((sections) =>
            sections.map((section) => section.querySelectorAll('button:not(.side)').length),
          ),
        INSERT_GROUPS,
      );
      insert.forEach(({ name, x, height }, index) => {
        near(x, INSERT_X[index], name);
        assert.equal(height, 72, name);
      });
      const width = (name) => insert.find((item) => item.name === name).width;
      for (const name of ['Icons', 'Link', 'Action']) assert.equal(width(name), 38, name);
      for (const name of [
        'Table',
        'Pictures',
        'Shapes',
        '3D Models',
        'Chart',
        'Zoom',
        'WordArt',
        'Video',
        'Audio',
      ])
        assert.equal(width(name), 50, name);
      // New Slide and Text Box carry their ▾ inside one 50 pt button.
      assert.equal(width('New Slide'), 50);
      assert.equal(width('Text Box'), 50);
      assert.equal(
        await page.locator('.insert [aria-label="New Slide"] .caption').textContent(),
        'New\nSlide',
      );
      assert.equal(
        await page.locator('.insert [aria-label="Header & Footer"] .caption').textContent(),
        'Header &\nFooter',
      );

      // Draw: Draw 38, Eraser ▾ 50, Lasso Select 38; a 180 × 60 pen gallery
      // 18 pt inside its group; Add ▾ 50; three 38 pt Ink buttons; the switch.
      await page.locator('#ribbon-tab-draw').click();
      const draw = await boxes(commands('draw'));
      assert.deepEqual(
        draw.map(({ name }) => [name, Math.round(draw.find((item) => item.name === name).width)]),
        [
          ['Draw', 38],
          ['Eraser', 50],
          ['Lasso Select', 38],
          ['Add', 50],
          ['Ink to Text', 38],
          ['Ink to Shape', 38],
          ['Ink to Math', 38],
        ],
      );
      const pens = await page.getByRole('radiogroup', { name: 'Pens' }).boundingBox();
      assert.deepEqual([pens.width, pens.height], [180, 60]);
      near(pens.x, 173, 'Pens');
      near(draw.find(({ name }) => name === 'Add').x, 372, 'Add');
      near(draw.find(({ name }) => name === 'Ink to Math').x, 517, 'Ink to Math');
      assert.equal((await page.locator('.draw .switch').boundingBox()).width, 62);

      // Design: the gallery fills ten 95 pt slots between 18 pt pagers, then
      // Variants, Colors, Fonts, Background Styles | Layout, Slide Size | Design Suggestions.
      await page.locator('#ribbon-tab-design').click();
      assert.equal((await page.getByRole('listbox', { name: 'Themes' }).boundingBox()).width, 950);
      const design = await boxes(commands('design'));
      assert.deepEqual(
        design.map(({ name }) => name),
        [
          'Variants',
          'Colors',
          'Fonts',
          'Background Styles',
          'Layout',
          'Slide Size',
          'Design Suggestions',
        ],
      );
      const native = {
        Variants: 1016,
        Colors: 1066,
        Fonts: 1116,
        'Background Styles': 1166,
        Layout: 1252,
        'Slide Size': 1302,
        'Design Suggestions': 1371,
      };
      for (const { name, x } of design) near(x, native[name], name);
      for (const name of ['Variants', 'Colors', 'Fonts', 'Layout', 'Slide Size'])
        assert.equal(design.find((item) => item.name === name).width, 50, name);
      assert.equal(design.find(({ name }) => name === 'Design Suggestions').width, 69);

      // 1200 pt: Insert stacks 3D Models / SmartArt / Chart as 22 pt rows and
      // Date & Time / Slide Number / Object as 24 × 22 icons; Design keeps
      // seven theme slots; Draw does not change.
      await page.setViewportSize({ width: 1200, height: 900 });
      await page.locator('#ribbon-tab-insert').click();
      insert = await boxes(commands('insert'));
      const at = (name) => insert.find((item) => item.name === name);
      for (const [name, y] of [
        ['3D Models', 0],
        ['SmartArt', 22],
        ['Chart', 44],
      ]) {
        assert.equal(at(name).height, 22, name);
        assert.equal(at(name).x, at('3D Models').x, name);
        assert.equal(at(name).y - at('3D Models').y, y, name);
      }
      for (const name of ['Date & Time', 'Slide Number', 'Object'])
        assert.deepEqual([at(name).width, at(name).height], [24, 22], name);
      near(at('Zoom').x, 543, 'Zoom at 1200');
      near(at('Audio').x, 1122, 'Audio at 1200');
      await page.locator('#ribbon-tab-design').click();
      assert.equal((await page.getByRole('listbox', { name: 'Themes' }).boundingBox()).width, 665);
      near(
        (await page.getByRole('button', { name: 'Variants', exact: true }).boundingBox()).x,
        731,
        'Variants at 1200',
      );

      // Japanese uses Mac PowerPoint's wording, broken after a particle or
      // where the script changes.
      await page.setViewportSize({ width: 1512, height: 900 });
      await page.locator('.lang select').selectOption('ja');
      await page.locator('#ribbon-tab-insert').click();
      assert.deepEqual((await boxes(commands('insert'))).map(({ name }) => name).slice(0, 4), [
        '新しいスライド',
        '表',
        '画像',
        'スクリーンショット',
      ]);
      assert.equal(
        await page.locator('.insert [aria-label="新しいスライド"] .caption').textContent(),
        '新しい\nスライド',
      );
      assert.equal(
        await page.locator('.insert [aria-label="ヘッダーとフッター"] .caption').textContent(),
        'ヘッダーと\nフッター',
      );
      await page.locator('#ribbon-tab-draw').click();
      assert.equal(
        await page.locator('.draw [aria-label="インクをテキストに変換"] .caption').textContent(),
        'インクを\nテキストに変換',
      );
      await page.locator('#ribbon-tab-design').click();
      assert.deepEqual(
        (await boxes(commands('design'))).map(({ name }) => name),
        [
          'バリエーション',
          '配色',
          'フォント',
          '背景のスタイル',
          'レイアウト',
          'スライドのサイズ',
          'デザイン アイデア',
        ],
      );
      await page.getByRole('button', { name: '配色', exact: true }).click();
      assert.deepEqual(
        await page
          .getByRole('group', { name: 'すべての色', exact: true })
          .getByRole('menuitemradio')
          .evaluateAll((nodes) => nodes.slice(3, 7).map((node) => node.getAttribute('aria-label'))),
        ['グレースケール', '暖かみのある青', '青', '青 II'],
      );
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

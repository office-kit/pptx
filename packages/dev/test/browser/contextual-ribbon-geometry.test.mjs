import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeImageOpacity,
  getTableCells,
  getTableCellBorders,
  getSlides,
  getSlideShapes,
  getTableDimensions,
  getTableStyleFlags,
  isTableShape,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

// Shape Format, Picture Format, Table Design and Table Layout as measured from
// Mac PowerPoint 16 through the accessibility API in 1512 × 900 and 1200 × 900
// pt windows (2026-10-07). CSS px equal Mac points; x is window-relative.
const PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const DECK = `import { Presentation, Slide, Image, Shape, Table } from '@office-kit/pptx-dsl';
const png = Uint8Array.from(atob('${PNG}'), (c) => c.charCodeAt(0));
export default (
  <Presentation>
    <Slide><Shape preset="rect" x={0.5} y={0.5} width={3} height={1.5} text="Shape" /><Image data={png} x={5} y={0.5} width={2} height={2} /><Table x={0.5} y={3.5} width={8} height={2} rows={[["A","B"],["C","D"]]} /></Slide>
  </Presentation>
);
`;
// Native group starts drift by a few points with label widths (the editor's
// font is not PowerPoint's), so positions are compared within this tolerance.
const TOLERANCE = 9;
// Native metrics were measured on macOS, where the editor's system-ui font is
// San Francisco like PowerPoint's. Elsewhere (CI's Linux fonts are wider) a
// control sized by its label may grow by up to WIDER_FONT, and the growth
// accumulates from left to right (Table Layout's labelled rows push its last
// groups about 30 pt right), so group starts may drift by POSITION_DRIFT.
// Heights and the sizes the CSS fixes (galleries, spin boxes, icon buttons)
// are checked the same way on every platform.
const MAC = process.platform === 'darwin';
const POSITION_DRIFT = 0.08;
const WIDER_FONT = 0.2;

const box = (locator) =>
  locator.evaluate((node) => {
    const rect = node.getBoundingClientRect();
    return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
  });

async function groups(panel) {
  return panel.locator('.ctx-ribbon > section').evaluateAll((nodes) =>
    nodes.map((node) => {
      // The group's content starts after its 10 pt margin (4 pt for the first).
      const first = node.firstElementChild.getBoundingClientRect().x;
      return [node.getAttribute('aria-label'), Math.round(first)];
    }),
  );
}

function assertStarts(actual, expected, name) {
  assert.deepEqual(
    actual.map(([label]) => label),
    expected.map(([label]) => label),
    `${name} group order`,
  );
  actual.forEach(([label, x], index) => {
    const native = expected[index][1];
    const tolerance = MAC ? TOLERANCE : Math.max(TOLERANCE, native * POSITION_DRIFT);
    assert.ok(
      Math.abs(x - native) <= tolerance,
      `${name} ${label} starts at ${x}, native ${native}`,
    );
  });
}

async function size(locator, width, height, name) {
  const rect = await box(locator);
  const labelled = MAC
    ? false
    : await locator.evaluate((node) => node.textContent.trim().length > 1);
  const widest = labelled ? width * (1 + WIDER_FONT) + 2 : width + 2;
  if (width !== null)
    assert.ok(
      rect.width >= width - 2 && rect.width <= widest,
      `${name} is ${rect.width} wide, native ${width}`,
    );
  if (height !== null)
    assert.ok(
      Math.abs(rect.height - height) <= 2,
      `${name} is ${rect.height} high, native ${height}`,
    );
}

test(
  'contextual tabs follow Mac PowerPoint geometry at 1512 and 1200 pt, in English and Japanese',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-contextual-geometry-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(file, DECK);
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1512, height: 900 } });
      await page.goto(preview.url + '/editor');
      await page.getByText('Saved to this project', { exact: true }).waitFor();
      const panel = page.locator('#ribbon-panel');
      const button = (name) => panel.getByRole('button', { name, exact: true });
      const changed = async (action) => {
        const before = await (await fetch(preview.url + '/editor/state')).json();
        await action();
        for (let i = 0; i < 200; i += 1) {
          const state = await (await fetch(preview.url + '/editor/state')).json();
          if (!state.building && state.revision !== before.revision) return;
          await new Promise((resolve) => setTimeout(resolve, 50));
        }
        throw new Error('The edit was not saved');
      };
      const open = async (hit, tab) => {
        await page.locator('.hit').nth(hit).click();
        await page.getByRole('tab', { name: tab, exact: true }).click();
      };

      // Shape Format (native groups at 10, 302, 621, 940, 997, 1312, 1456).
      await open(0, 'Shape Format');
      assertStarts(
        await groups(panel),
        [
          ['Insert Shapes', 10],
          ['Shape Styles', 302],
          ['WordArt Styles', 621],
          ['Accessibility', 940],
          ['Arrange', 997],
          ['Size', 1312],
          ['Format Pane', 1456],
        ],
        'Shape Format',
      );
      await size(button('Next Shapes gallery'), 18, 58, 'Next Shapes gallery');
      await size(button('Rectangle'), 20, 18, 'shape strip cell');
      for (const name of ['Text Box', 'Edit Shape', 'Merge Shapes'])
        await size(button(name), null, 22, name);
      assert.equal(await button('Merge Shapes').isDisabled(), true);
      await size(panel.locator('.ctx-paint', { hasText: 'Shape Fill' }), 50, null, 'Shape Fill');
      for (const name of ['Shape Effects', 'Text Effects', 'Group', 'Rotate'])
        await size(button(name), 38, 26, name);
      for (const name of ['Bring Forward', 'Reorder Objects', 'Align'])
        await size(button(name), 50, null, name);
      for (const name of ['Height', 'Width'])
        await size(panel.getByRole('spinbutton', { name, exact: true }), 78, 24, name);
      {
        const last = await box(button('Format Pane'));
        const ribbon = await box(panel.locator('.ctx-ribbon'));
        assert.ok(
          last.x + last.width <= ribbon.x + ribbon.width + 1,
          'every Shape Format group fits at 1512 pt',
        );
      }

      // Picture Format (native groups at 10, 96, 369, 754, 811, 1127, 1372, 1433).
      await open(1, 'Picture Format');
      assertStarts(
        await groups(panel),
        [
          ['Remove Background', 10],
          ['Adjust', 96],
          ['Picture Styles', 369],
          ['Accessibility', 754],
          ['Arrange', 811],
          ['Size', 1127],
          ['Format Pane', 1372],
          ['Animate as Background', 1433],
        ],
        'Picture Format',
      );
      for (const name of ['Remove Background', 'Animate as Background', 'Compress Pictures'])
        assert.equal(await button(name).isDisabled(), true, `${name} is unavailable`);
      for (const name of ['Color', 'Artistic Effects', 'Transparency', 'Picture Effects'])
        await size(button(name), null, 22, name);
      for (const name of ['Picture Quality', 'Change Picture', 'Reset Picture'])
        await size(button(name), 36, 22, name);
      await size(button('Crop'), 50, null, 'Crop');
      await size(button('Next Quick Styles gallery'), 18, 58, 'Picture Styles arrow');

      // Transparency ▸ 50% writes the picture's alpha.
      await button('Transparency').click();
      await changed(() =>
        panel.getByRole('menuitem', { name: 'Transparency: 50%', exact: true }).click(),
      );
      const saved = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return { pres, shapes: getSlideShapes(getSlides(pres)[0]) };
      };
      assert.equal(getShapeImageOpacity((await saved()).shapes[1]), 0.5);

      // Table Design (native groups at 10, 240, 898, 1003).
      await open(2, 'Table Design');
      assertStarts(
        await groups(panel),
        [
          ['Table Style Options', 10],
          ['Table Styles', 240],
          ['WordArt Styles', 898],
          ['Draw Borders', 1003],
        ],
        'Table Design',
      );
      assert.deepEqual(
        await panel
          .locator('section[aria-label="Table Style Options"]')
          .getByRole('checkbox')
          .evaluateAll((boxes) => boxes.map((input) => input.labels[0].textContent.trim())),
        ['Header Row', 'Total Row', 'Banded Rows', 'First Column', 'Last Column', 'Banded Columns'],
      );
      await size(panel.locator('.table-strip'), 518, null, 'Table Styles strip');
      for (const name of ['Borders', 'Effects']) await size(button(name), null, 22, name);
      for (const name of ['Pen Style', 'Pen Weight'])
        await size(panel.getByRole('combobox', { name, exact: true }), 120, 22, name);
      for (const name of ['Draw Table', 'Eraser'])
        assert.equal(await button(name).isDisabled(), true, `${name} is unavailable`);
      await changed(() =>
        panel.getByRole('checkbox', { name: 'Header Row', exact: true }).uncheck(),
      );
      // Borders ▸ All Borders draws the pen (1 pt solid) on every edge.
      await button('Borders').click();
      await changed(() =>
        panel.getByRole('menuitem', { name: 'All Borders', exact: true }).click(),
      );

      // Table Layout (native groups at 10, 130, 475, 570, 820, 1019, 1214, 1414).
      await page.getByRole('tab', { name: 'Table Layout', exact: true }).click();
      assertStarts(
        await groups(panel),
        [
          ['Table', 10],
          ['Rows & Columns', 130],
          ['Merge', 475],
          ['Cell Size', 570],
          ['Alignment', 820],
          ['Table Size', 1019],
          ['Arrange', 1214],
          ['Format Pane', 1414],
        ],
        'Table Layout',
      );
      for (const [name, width] of [
        ['Select', 50],
        ['Delete', 50],
        ['Insert Row Above', 60],
        ['Insert Row Below', 60],
        ['Insert Column Left', 68],
        ['Merge Cells', 38],
        ['Split Cells', 38],
        ['Text Direction', 52],
        ['Cell Margins', 50],
      ])
        await size(button(name), width, null, name);
      for (const name of ['Align Left', 'Center Text', 'Center Vertically', 'Align Bottom'])
        await size(button(name), 26, 26, name);
      for (const name of ['Distribute Rows', 'Distribute Columns'])
        await size(button(name), null, 26, name);
      for (const name of ['Align', 'Group', 'Rotate']) await size(button(name), null, 22, name);
      for (const name of ['Table Row Height', 'Table Column Width'])
        await size(panel.getByRole('spinbutton', { name, exact: true }), 78, 24, name);

      await changed(() => button('Insert Row Below').click());
      const { pres, shapes } = await saved();
      const table = shapes.find(isTableShape);
      assert.equal(getTableDimensions(table).rows, 3);
      assert.equal(getTableStyleFlags(table).firstRow, false);
      const borders = getTableCellBorders(pres, getTableCells(table)[0][0]);
      assert.equal(borders.top?.widthEmu, 12700);
      assert.equal(borders.right?.widthEmu, 12700);

      // 1200 pt: the compact layouts PowerPoint shows at this width.
      await page.setViewportSize({ width: 1200, height: 900 });
      await page.locator('.hit').nth(2).click();
      await page.getByRole('tab', { name: 'Table Layout', exact: true }).click();
      for (const name of ['Insert Row Below', 'Insert Column Left', 'Insert Column Right'])
        await size(button(name), null, 22, name);
      for (const name of ['Distribute Rows', 'Distribute Columns'])
        await size(button(name), 26, 26, name);
      await size(button('Arrange'), 50, null, 'collapsed Arrange');
      await page.getByRole('tab', { name: 'Table Design', exact: true }).click();
      for (const name of ['Borders', 'Effects']) await size(button(name), 36, 22, name);
      await open(0, 'Shape Format');
      await size(button('Shapes'), 50, null, 'Shapes');
      for (const name of ['Text Box', 'Edit Shape', 'Merge Shapes'])
        await size(button(name), 36, 22, name);
      await size(button('Arrange'), 50, null, 'collapsed Arrange');
      assertStarts(
        (await groups(panel)).slice(0, 3),
        [
          ['Insert Shapes', 10],
          ['Shape Styles', 115],
          ['WordArt Styles', 434],
        ],
        'Shape Format at 1200',
      );
      const fits = () =>
        panel.locator('.ctx-ribbon').evaluate((node) => node.scrollWidth <= node.clientWidth + 1);
      assert.ok(await fits(), 'Shape Format fits at 1200 pt');

      // Japanese uses Mac PowerPoint's own wording.
      await page.locator('.lang select').selectOption('ja');
      await page.locator('.hit').nth(2).click();
      const tabs = (await page.getByRole('tablist').getByRole('tab').allTextContents()).map(
        (name) => name.trim(),
      );
      assert.deepEqual(tabs.slice(-2), ['テーブル デザイン', 'テーブル レイアウト']);
      await page.getByRole('tab', { name: 'テーブル レイアウト', exact: true }).click();
      for (const name of [
        '上に行を挿入',
        '下に行を挿入',
        'セルの結合',
        '高さを揃える',
        'セル内の配置',
      ])
        assert.equal(await button(name).count(), 1, name);
      await page.getByRole('tab', { name: 'テーブル デザイン', exact: true }).click();
      assert.equal(
        await panel.getByRole('checkbox', { name: 'タイトル行', exact: true }).count(),
        1,
      );
      await page.locator('.hit').nth(1).click();
      await page.getByRole('tab', { name: '図の形式', exact: true }).click();
      for (const name of ['背景の削除', '修整', 'アート効果', '透明度', '図の効果', 'トリミング'])
        assert.equal(await button(name).count(), 1, name);
      await open(0, '図形の書式');
      for (const name of ['書式ウィンドウ', '代替テキスト', '図形の効果'])
        assert.equal(await button(name).count(), 1, name);
      assert.ok(await fits(), 'the Japanese Shape Format tab fits at 1200 pt');
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

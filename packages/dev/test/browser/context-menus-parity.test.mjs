import assert from 'node:assert/strict';
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getParagraphPropertiesEffective,
  getShapeKind,
  getShapeParagraphElements,
  getSlideShapes,
  getSlides,
  getTableCells,
  getTableCellSpan,
  getTableCellText,
  isTableShape,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';
import { installRichTextSelection } from '../helpers/rich-text.mjs';

// The reference desktop app's (Mac, 16) text, table-cell, picture and slide-background context
// menus (native AX captures, 2026-10-07): item order, separators, submenus and
// shortcut hints. The system-provided Continuity Camera and Services entries
// are not the reference desktop app's and are not reproduced.

const PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const DECK = `import { Presentation, Slide, Image, Table, Text } from '@office-kit/pptx-dsl';
const png = Uint8Array.from(atob('${PNG}'), (c) => c.charCodeAt(0));
export default (
  <Presentation>
    <Slide><Text x={1} y={1} width={8} height={2}>Before Target After</Text></Slide>
    <Slide><Table x={1} y={1} width={6} height={2} rows={[["A","B"],["C","D"]]} /></Slide>
    <Slide><Image data={png} x={2} y={2} width={2} height={2} /></Slide>
  </Presentation>
);
`;
const SHOTS = process.env.CONTEXT_MENU_SHOTS;

const TEXT_MENU = [
  'Cut ⌘X',
  'Copy ⌘C',
  'Paste ⌘V',
  '----',
  'Exit Edit Text',
  'Font... ⌘T',
  'Paragraph... ⌥⌘M',
  'Bullets ▸',
  'Numbering ▸',
  '----',
  'Thesaurus... ⌃⌥⌘R (disabled)',
  'Translate... (disabled)',
  '----',
  'Format Text Effects...',
  'Format Shape... ⇧⌘1',
  '----',
  'Lock',
  '----',
  'Hyperlink... ⌘K',
  '----',
  'New Comment ⇧⌘M',
];

const CELL_MENU = [
  'Cut ⌘X',
  'Copy ⌘C',
  'Paste ⌘V (disabled)',
  '----',
  'Font... ⌘T',
  'Paragraph... ⌥⌘M',
  'Bullets ▸',
  'Numbering ▸',
  '----',
  'Insert ▸',
  'Delete ▸',
  'Select ▸',
  '----',
  'Merge Cells (disabled)',
  'Split Cells... (disabled)',
  '----',
  'Thesaurus... ⌃⌥⌘R (disabled)',
  'Translate... (disabled)',
  '----',
  'Format Text Effects... (disabled)',
  'Format Shape... ⇧⌘1',
  '----',
  'Lock',
  '----',
  'Hyperlink... ⌘K',
  '----',
  'New Comment ⇧⌘M',
];

const PICTURE_MENU = [
  'Cut ⌘X',
  'Copy ⌘C',
  'Paste ⌘V (disabled)',
  '----',
  'Change Picture ▸',
  '----',
  'Reorder Objects (disabled)',
  'Reorder Overlapping Objects (disabled)',
  '----',
  'Group ▸',
  'Bring to Front ▸',
  'Send to Back ▸',
  'Lock',
  '----',
  'Hyperlink... ⌘K',
  '----',
  'Edit Picture (disabled)',
  'Save as Picture... (disabled)',
  '----',
  'View Alt Text...',
  'Crop ⇧C',
  'Auto Crop (disabled)',
  'Size and Position...',
  'Format Picture... ⇧⌘1',
  '----',
  'Action Settings...',
  '----',
  'New Comment ⇧⌘M',
];

const BACKGROUND_MENU = [
  'Cut ⌘X (disabled)',
  'Copy ⌘C (disabled)',
  'Paste ⌘V (disabled)',
  'Paste Special... ⌃⌘V (disabled)',
  '----',
  'New Slide ⇧⌘N',
  'Duplicate Slide ⇧⌘D',
  'Delete Slide',
  '----',
  'Hide Slide',
  '----',
  'Ruler',
  'Grid and Guides ▸',
  'Zoom...',
  '----',
  'Format Background...',
  '----',
  'Slide Show ⇧⌘↩',
  '----',
  'New Comment ⇧⌘M',
];

/** "Label accel ▸ (disabled)" per row, "----" per separator. */
function entries(menu) {
  return menu
    .locator(
      ':scope > .ctx-item, :scope > .branch > .ctx-item, :scope > .gallery > .tile, :scope > .ctx-sep',
    )
    .evaluateAll((nodes) =>
      nodes.map((node) => {
        if (node.classList.contains('ctx-sep')) return '----';
        const accel = node.querySelector('.accel')?.textContent;
        return [
          node.getAttribute('aria-label'),
          accel,
          node.getAttribute('aria-haspopup') ? '▸' : '',
          node.disabled ? '(disabled)' : '',
        ]
          .filter(Boolean)
          .join(' ');
      }),
    );
}

async function open(file) {
  const preview = await startPreview(file);
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ viewport: { width: 1500, height: 1000 } });
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  const page = await context.newPage();
  await installRichTextSelection(page);
  const errors = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto(preview.url);
  const editor = page.frameLocator('#editor-frame');
  await editor.getByText('Saved to this project', { exact: true }).waitFor();
  const deck = async () =>
    loadPresentation(new Uint8Array(await (await fetch(`${preview.url}/deck.pptx`)).arrayBuffer()));
  const changed = async (action) => {
    const before = (await waitForState(preview.url, () => true)).revision;
    await action();
    await waitForState(preview.url, (state) => state.revision !== before);
  };
  return { preview, browser, page, editor, errors, deck, changed };
}

async function shot(page, name) {
  if (!SHOTS) return;
  await mkdir(SHOTS, { recursive: true });
  await page.screenshot({ path: join(SHOTS, `${name}.png`) });
}

test(
  'the text-editing menu keeps the caret and selection and acts on the selected range',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-context-text-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(file, DECK);
    let session;
    try {
      session = await open(file);
      const { page, editor, errors, deck, changed } = session;
      const menu = editor.locator('.ctx:not(.submenu)');
      const input = editor.locator('.canvas-shell .inline-edit');
      const frame = await page.locator('#editor-frame').boundingBox();
      const selectTarget = () =>
        input.evaluate((node) => {
          node.focus({ preventScroll: true });
          window.selectEditorText(node, 7, 13);
          node.dispatchEvent(new Event('select', { bubbles: true }));
        });
      // Right-click inside the selected text, as a user would.
      const contextOnSelection = async () => {
        const point = await input.evaluate(() => {
          const rect = document.getSelection().getRangeAt(0).getBoundingClientRect();
          return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
        });
        await page.mouse.click(frame.x + point.x, frame.y + point.y, { button: 'right' });
        await menu.waitFor();
      };
      const state = () =>
        input.evaluate((node) => ({
          focused: node === node.ownerDocument.activeElement,
          selected: node.ownerDocument.getSelection().toString(),
        }));
      const clipboard = async (expected) => {
        const deadline = Date.now() + 10000;
        while ((await page.evaluate(() => navigator.clipboard.readText())) !== expected) {
          assert.ok(Date.now() < deadline, `clipboard did not become ${expected}`);
          await page.waitForTimeout(100);
        }
      };
      const runs = async () =>
        getShapeParagraphElements(getSlideShapes(getSlides(await deck())[0])[0], 0).map(
          (element) => [element.text, element.format?.bold === true],
        );

      await editor.locator('.hit').first().dblclick();
      await input.waitFor();
      await selectTarget();
      await contextOnSelection();
      assert.deepEqual(await entries(menu), TEXT_MENU);
      // The menu opens without taking focus: the caret and selection stay.
      assert.deepEqual(await state(), { focused: true, selected: 'Target' });
      await menu.getByRole('menuitem', { name: 'Bullets', exact: true }).hover();
      const bullets = editor.getByRole('menu', { name: 'Bullets', exact: true });
      assert.deepEqual(await entries(bullets), [
        'None',
        'Filled Round Bullets',
        'Hollow Round Bullets',
        'Filled Square Bullets',
        'Hollow Square Bullets',
        'Star Bullets',
        'Arrow Bullets',
        'Checkmark Bullets',
        '----',
        'Bullets and Numbering... (disabled)',
      ]);
      await shot(page, 'ours-text-edit-sub-bullets');
      await menu.getByRole('menuitem', { name: 'Numbering', exact: true }).hover();
      assert.deepEqual(
        (await entries(editor.getByRole('menu', { name: 'Numbering', exact: true }))).slice(0, 8),
        [
          'None',
          '1. 2. 3.',
          '1) 2) 3)',
          'I. II. III.',
          'A. B. C.',
          'a) b) c)',
          'a. b. c.',
          'i. ii. iii.',
        ],
      );
      assert.deepEqual(await state(), { focused: true, selected: 'Target' });
      // Escape closes the menu but keeps editing.
      await page.keyboard.press('Escape');
      assert.equal(await menu.count(), 0);
      assert.deepEqual(await state(), { focused: true, selected: 'Target' });

      // Copy puts the selected range on the clipboard without touching the text.
      await contextOnSelection();
      await menu.getByRole('menuitem', { name: 'Copy', exact: true }).click();
      await clipboard('Target');
      assert.deepEqual(await state(), { focused: true, selected: 'Target' });

      // Font... applies Bold to the selected range only.
      await contextOnSelection();
      await shot(page, 'ours-text-edit');
      await menu.getByRole('menuitem', { name: 'Font...', exact: true }).click();
      const dialog = editor.getByRole('dialog', { name: 'Font', exact: true });
      await dialog.waitFor();
      await dialog
        .locator('label', { hasText: 'Font style' })
        .locator('select')
        .selectOption('true:false');
      await changed(() => dialog.getByRole('button', { name: 'OK', exact: true }).click());
      assert.equal((await state()).focused, true);
      assert.deepEqual(await runs(), [
        ['Before ', false],
        ['Target', true],
        [' After', false],
      ]);

      // Paragraph... opens over the edit and hands focus back to the text.
      await selectTarget();
      await contextOnSelection();
      await menu.getByRole('menuitem', { name: 'Paragraph...', exact: true }).click();
      const paragraph = editor.getByRole('dialog', { name: 'Paragraph', exact: true });
      await paragraph.waitFor();
      await paragraph.getByRole('button', { name: 'Cancel', exact: true }).click();
      await paragraph.waitFor({ state: 'detached' });
      assert.equal((await state()).focused, true);

      // Format Text Effects... opens the pane on Text Options ▸ Text Effects.
      await contextOnSelection();
      await menu.getByRole('menuitem', { name: 'Format Text Effects...', exact: true }).click();
      assert.equal(
        await editor
          .getByRole('tab', { name: 'Text Effects', exact: true })
          .getAttribute('aria-selected'),
        'true',
      );
      assert.deepEqual(await state(), { focused: true, selected: 'Target' });

      // A gallery bullet applies to the paragraph being edited, still editing.
      await selectTarget();
      await contextOnSelection();
      await menu.getByRole('menuitem', { name: 'Bullets', exact: true }).hover();
      await changed(() =>
        editor.getByRole('menuitemradio', { name: 'Hollow Square Bullets', exact: true }).click(),
      );
      assert.equal((await state()).focused, true);
      assert.deepEqual(
        getParagraphPropertiesEffective(
          await deck(),
          getSlideShapes(getSlides(await deck())[0])[0],
          0,
        ).bullet,
        { char: '❑' },
      );

      // Cut removes the range after the clipboard write; Exit Edit Text ends editing.
      await selectTarget();
      await contextOnSelection();
      await page.evaluate(() => navigator.clipboard.writeText(''));
      await menu.getByRole('menuitem', { name: 'Cut', exact: true }).click();
      await clipboard('Target');
      await input.click({ button: 'right' });
      await changed(() =>
        menu.getByRole('menuitem', { name: 'Exit Edit Text', exact: true }).click(),
      );
      assert.equal(await input.count(), 0);
      assert.deepEqual(
        getShapeParagraphElements(getSlideShapes(getSlides(await deck())[0])[0], 0)
          .map((element) => element.text)
          .join(''),
        'Before  After',
      );
      assert.equal(await editor.locator('.hit.selected').count(), 1);

      // Japanese labels (unverified against native Japanese menus).
      await editor.locator('.lang select').selectOption('ja');
      await editor.locator('.hit').first().dblclick();
      await input.waitFor();
      await input.click({ button: 'right' });
      assert.deepEqual((await entries(menu)).slice(0, 9), [
        '切り取り ⌘X (disabled)',
        'コピー ⌘C (disabled)',
        '貼り付け ⌘V',
        '----',
        'テキストの編集を終了',
        'フォント... ⌘T',
        '段落... ⌥⌘M',
        '箇条書き ▸',
        '段落番号 ▸',
      ]);
      await page.keyboard.press('Escape');
      await page.keyboard.press('Escape');
      assert.deepEqual(errors, []);
    } finally {
      await session?.browser.close();
      await session?.preview.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'the table-cell menu inserts, selects, merges and deletes rows and columns',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-context-cell-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(file, DECK);
    let session;
    try {
      session = await open(file);
      const { page, editor, errors, deck, changed } = session;
      const menu = editor.locator('.ctx:not(.submenu)');
      const table = async () => getSlideShapes(getSlides(await deck())[1]).find(isTableShape);
      const values = async () =>
        getTableCells(await table()).map((row) => row.map(getTableCellText));
      await editor.getByRole('button', { name: 'Slide 2', exact: true }).first().click();
      const hit = editor.locator('.hit').first();
      await hit.waitFor();
      // Row heights grow with inserted rows, so measure the table each time.
      const cellContext = async (row, col) => {
        const box = await hit.boundingBox();
        const grid = await values();
        await page.mouse.click(
          box.x + (box.width * (col + 0.5)) / grid[0].length,
          box.y + (box.height * (row + 0.5)) / grid.length,
          { button: 'right' },
        );
      };
      await cellContext(0, 0);
      assert.deepEqual(await entries(menu), CELL_MENU);
      await shot(page, 'ours-table-cell');
      const submenu = async (name) => {
        await menu.getByRole('menuitem', { name, exact: true }).hover();
        return entries(editor.getByRole('menu', { name, exact: true }));
      };
      assert.deepEqual(await submenu('Insert'), [
        'Insert Columns to the Left',
        'Insert Columns to the Right',
        'Insert Rows Above',
        'Insert Rows Below',
      ]);
      assert.deepEqual(await submenu('Delete'), ['Delete Columns', 'Delete Rows', 'Delete Table']);
      assert.deepEqual(await submenu('Select'), ['Select Table', 'Select Column', 'Select Row']);

      await menu.getByRole('menuitem', { name: 'Insert', exact: true }).hover();
      await changed(() =>
        editor.getByRole('menuitem', { name: 'Insert Rows Above', exact: true }).click(),
      );
      assert.deepEqual(await values(), [
        ['', ''],
        ['A', 'B'],
        ['C', 'D'],
      ]);
      // The selection follows the original cell.
      await cellContext(1, 0);
      await menu.getByRole('menuitem', { name: 'Select', exact: true }).hover();
      await editor.getByRole('menuitem', { name: 'Select Row', exact: true }).click();
      await cellContext(1, 0);
      await changed(() => menu.getByRole('menuitem', { name: 'Merge Cells', exact: true }).click());
      assert.equal(getTableCellSpan(getTableCells(await table())[1][0]).gridSpan, 2);
      // Rows cannot be changed while cells are merged; Split Cells... unmerges.
      await cellContext(1, 0);
      await menu.getByRole('menuitem', { name: 'Delete', exact: true }).hover();
      assert.equal(
        await editor.getByRole('menuitem', { name: 'Delete Rows', exact: true }).isDisabled(),
        true,
      );
      await changed(() =>
        menu.getByRole('menuitem', { name: 'Split Cells...', exact: true }).click(),
      );
      assert.equal(getTableCellSpan(getTableCells(await table())[1][0]).gridSpan, 1);
      await cellContext(0, 0);
      await menu.getByRole('menuitem', { name: 'Delete', exact: true }).hover();
      await changed(() =>
        editor.getByRole('menuitem', { name: 'Delete Rows', exact: true }).click(),
      );
      assert.equal((await values()).length, 2);

      // While a cell's text is edited, the cell menu acts on the text.
      await editor.locator('.lang select').selectOption('ja');
      const box = await hit.boundingBox();
      await hit.dblclick({ position: { x: box.width / 4, y: box.height / 4 } });
      const input = editor.locator('.canvas-shell .inline-edit');
      await input.waitFor();
      await input.click({ button: 'right' });
      assert.equal(await input.evaluate((node) => node === node.ownerDocument.activeElement), true);
      assert.deepEqual((await entries(menu)).slice(9, 16), [
        '挿入 ▸',
        '削除 ▸',
        '選択 ▸',
        '----',
        'セルの結合 (disabled)',
        'セルの分割... (disabled)',
        '----',
      ]);
      await menu.getByRole('menuitem', { name: '挿入', exact: true }).hover();
      await changed(() =>
        editor.getByRole('menuitem', { name: '右に列を挿入', exact: true }).click(),
      );
      assert.equal((await values())[0].length, 3);
      assert.deepEqual(errors, []);
    } finally {
      await session?.browser.close();
      await session?.preview.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'the picture and slide-background menus follow the reference desktop app',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-context-picture-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(file, DECK);
    let session;
    try {
      session = await open(file);
      const { page, editor, errors, deck } = session;
      const menu = editor.locator('.ctx:not(.submenu)');
      await editor.getByRole('button', { name: 'Slide 3', exact: true }).first().click();
      const hit = editor.locator('.hit').first();
      await hit.click({ button: 'right' });
      assert.deepEqual(await entries(menu), PICTURE_MENU);
      await shot(page, 'ours-picture');
      await menu.getByRole('menuitem', { name: 'Change Picture', exact: true }).hover();
      assert.deepEqual(await entries(editor.getByRole('menu', { name: 'Change Picture' })), [
        'From a File...',
        'From Stock Images... (disabled)',
        'From Online Sources... (disabled)',
        'From Brand Images... (disabled)',
        'From Icons... (disabled)',
        'From Clipboard... (disabled)',
      ]);
      await menu.getByRole('menuitem', { name: 'Group', exact: true }).hover();
      assert.deepEqual(await entries(editor.getByRole('menu', { name: 'Group', exact: true })), [
        'Group ⌥⌘G (disabled)',
        'Regroup ⌥⌘J (disabled)',
        '----',
        'Ungroup ⌥⇧⌘G (disabled)',
      ]);
      await shot(page, 'ours-picture-sub-group');
      await menu.getByRole('menuitem', { name: 'Crop', exact: true }).click();
      await editor.getByRole('dialog').first().waitFor();
      await page.keyboard.press('Escape');
      assert.equal(getShapeKind(getSlideShapes(getSlides(await deck())[2])[0]), 'picture');

      // The slide's own menu.
      await editor.locator('.stage').click({ button: 'right', position: { x: 600, y: 400 } });
      assert.deepEqual(await entries(menu), BACKGROUND_MENU);
      await shot(page, 'ours-slide-background');
      await menu.getByRole('menuitem', { name: 'Grid and Guides', exact: true }).hover();
      assert.deepEqual(
        await editor
          .getByRole('menu', { name: 'Grid and Guides', exact: true })
          .locator(':scope > .ctx-item, :scope > .ctx-sep')
          .evaluateAll((nodes) =>
            nodes.map((node) =>
              node.classList.contains('ctx-sep')
                ? '----'
                : `${node.getAttribute('aria-label')}${node.getAttribute('aria-checked') === 'true' ? ' ✓' : ''}`,
            ),
          ),
        [
          'Add Vertical Guide',
          'Add Horizontal Guide',
          'Delete',
          '----',
          'Smart Guides ✓',
          'Guides',
          'Gridlines',
          '----',
          'Snap to Grid',
          '----',
          'Grid Options...',
        ],
      );
      await shot(page, 'ours-slide-background-sub-grid-and-guides');
      await editor.getByRole('menuitemcheckbox', { name: 'Gridlines', exact: true }).click();
      assert.equal(await editor.locator('.grid-dots').count(), 1);
      await editor.locator('.stage').click({ button: 'right', position: { x: 600, y: 400 } });
      await menu.getByRole('menuitemcheckbox', { name: 'Ruler', exact: true }).click();
      await editor.locator('.rulers').first().waitFor();
      await editor.locator('.lang select').selectOption('ja');
      await editor.locator('.stage').click({ button: 'right', position: { x: 600, y: 400 } });
      assert.deepEqual((await entries(menu)).slice(9, 14), [
        '非表示スライドに設定',
        '----',
        'ルーラー',
        'グリッドとガイド ▸',
        'ズーム...',
      ]);
      await menu.getByRole('menuitem', { name: '背景の書式設定...', exact: true }).click();
      await editor.getByRole('region', { name: '背景の書式設定', exact: true }).waitFor();
      assert.deepEqual(errors, []);
    } finally {
      await session?.browser.close();
      await session?.preview.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

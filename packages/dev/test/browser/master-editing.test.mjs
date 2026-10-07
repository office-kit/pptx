import assert from 'node:assert/strict';
import { copyFile, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getHandoutMasterPlaceholders,
  getHandoutSlidesPerPage,
  getNotesMasterPlaceholders,
  getNotesPageSize,
  getSlideLayoutName,
  getSlideLayoutPlaceholders,
  getSlideMasterLayouts,
  getSlideMasterName,
  getSlideMasterPartNames,
  getSlideMasterPlaceholders,
  isSlideLayoutBackgroundGraphicsHidden,
  isSlideMasterPreserved,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';

// The Slide, Handout and Notes Master tabs' editing commands, each checked
// against the deck the editor saves.
const DECK = `import {readFileSync} from 'node:fs';
import {Presentation,Slide,Fill} from '@office-kit/pptx-dsl';
const source = new Uint8Array(readFileSync(new URL('./template.pptx', import.meta.url)));
export default <Presentation source={source} mode="compose">
  <Slide layout={{name:'Title Only'}}><Fill target={{placeholder:{type:'title'}}}>One</Fill></Slide>
  <Slide layout={{name:'Title and Content'}}><Fill target={{placeholder:{type:'title'}}}>Two</Fill></Slide>
</Presentation>;`;

const LABELS = {
  en: {
    view: 'View',
    slideMaster: 'Slide Master',
    handoutMaster: 'Handout Master',
    notesMaster: 'Notes Master',
    pane: 'Slide Master Pane',
    insertSlideMaster: 'Insert Slide Master',
    insertLayout: 'Insert Layout',
    delete: 'Delete',
    rename: 'Rename',
    renameMaster: 'Rename Master',
    preserve: 'Preserve',
    masterLayout: 'Master Layout',
    insertPlaceholder: 'Insert Placeholder',
    picture: 'Picture',
    title: 'Title',
    footers: 'Footers',
    hideBackground: 'Hide Background Graphics',
    date: 'Date',
    ok: 'OK',
    handoutOrientation: 'Handout Orientation',
    landscape: 'Landscape',
    fourSlides: '4 Slides',
    header: 'Header',
    body: 'Body',
  },
  ja: {
    view: '表示',
    slideMaster: 'スライド マスター',
    handoutMaster: '配布資料マスター',
    notesMaster: 'ノート マスター',
    pane: 'スライド マスター ウィンドウ',
    insertSlideMaster: 'スライド マスターの挿入',
    insertLayout: 'レイアウトの挿入',
    delete: '削除',
    rename: '名前の変更',
    renameMaster: 'マスター名の変更',
    preserve: '保持',
    masterLayout: 'マスターのレイアウト',
    insertPlaceholder: 'プレースホルダーの挿入',
    picture: '図',
    title: 'タイトル',
    footers: 'フッター',
    hideBackground: '背景グラフィックを表示しない',
    date: '日付',
    ok: 'OK',
    handoutOrientation: '配布資料の方向',
    landscape: '横',
    fourSlides: '4 枚',
    header: 'ヘッダー',
    body: '本文',
  },
};

for (const locale of ['en', 'ja'])
  test(
    `master views edit masters, layouts and page masters (${locale})`,
    { timeout: 180000 },
    async () => {
      const L = LABELS[locale];
      const dir = await mkdtemp(join(tmpdir(), 'office-master-editing-'));
      let preview, browser;
      try {
        await copyFile(
          new URL('../../../../test/fixtures/minimal/one-text-slide.pptx', import.meta.url),
          join(dir, 'template.pptx'),
        );
        const file = join(dir, 'deck.tsx');
        await writeFile(file, DECK);
        preview = await startPreview(file);
        browser = await chromium.launch({ headless: true });
        const page = await browser.newPage({ viewport: { width: 1512, height: 900 } });
        const errors = [];
        page.on('pageerror', (error) => errors.push(error.message));
        await page.addInitScript(
          (language) => localStorage.setItem('ok-editor-locale', language),
          locale,
        );
        await page.goto(preview.url + '/editor');
        await page.locator('.nav [data-slide-index="1"]').waitFor();

        const deck = async () =>
          loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          );
        const changed = async (action) => {
          const before = (await waitForState(preview.url, () => true)).revision;
          await action();
          await waitForState(preview.url, (state) => state.revision !== before);
        };
        const ribbon = page.locator('#ribbon-panel');
        const ribbonButton = (name) => ribbon.getByRole('button', { name, exact: true }).first();
        const checkbox = (name) => ribbon.getByRole('checkbox', { name, exact: true });
        const enter = async (name) => {
          await page.getByRole('tab', { name: L.view, exact: true }).click();
          await ribbonButton(name).click();
        };

        // Slide Master view opens on the current slide's layout.
        await enter(L.slideMaster);
        const pane = page.getByRole('navigation', { name: L.pane, exact: true });
        await pane.waitFor();
        const [firstMaster] = getSlideMasterPartNames(await deck());

        // Insert Slide Master adds a master with PowerPoint's eleven layouts and selects it.
        await changed(() => ribbonButton(L.insertSlideMaster).click());
        let pres = await deck();
        const masters = getSlideMasterPartNames(pres);
        assert.equal(masters.length, 2);
        const added = masters[1];
        assert.equal(getSlideMasterLayouts(pres, added).length, 11);
        assert.equal(await pane.locator('.cell.master').count(), 2);
        assert.equal(await pane.locator('.cell.master[aria-current="true"]').count(), 1);

        // Rename and Preserve act on the selected master.
        await ribbonButton(L.rename).click();
        const dialog = page.getByRole('dialog', { name: L.renameMaster, exact: true });
        await dialog.locator('input').first().fill('Brand');
        await changed(() => dialog.locator('button[type="submit"]').click());
        await changed(() => ribbonButton(L.preserve).click());
        pres = await deck();
        assert.equal(getSlideMasterName(pres, added), 'Brand');
        assert.equal(isSlideMasterPreserved(pres, added), false);
        assert.equal(await ribbonButton(L.preserve).getAttribute('aria-pressed'), 'false');

        // Master Layout drops the master's date placeholder.
        await ribbonButton(L.masterLayout).click();
        const masterLayout = ribbon.getByRole('dialog', { name: L.masterLayout, exact: true });
        await masterLayout.getByRole('checkbox', { name: L.date, exact: true }).uncheck();
        await changed(() => masterLayout.getByRole('button', { name: L.ok, exact: true }).click());
        assert.deepEqual(
          getSlideMasterPlaceholders(await deck(), added).map((placeholder) => placeholder.type),
          ['title', 'body', 'ftr', 'sldNum'],
        );

        // Insert Layout adds a Custom Layout after the master's last layout.
        await changed(() => ribbonButton(L.insertLayout).click());
        pres = await deck();
        let layouts = getSlideMasterLayouts(pres, added);
        assert.equal(layouts.length, 12);
        const custom = layouts[11];
        assert.equal(getSlideLayoutName(custom), 'Custom Layout');
        assert.equal(
          await pane.locator('.cell.layout[aria-current="true"]').getAttribute('aria-label'),
          'Custom Layout',
        );

        // The layout commands: Insert Placeholder, Title, Footers, Hide Background Graphics.
        await ribbonButton(L.insertPlaceholder).click();
        await changed(() => ribbon.getByRole('menuitem', { name: L.picture, exact: true }).click());
        await changed(() => checkbox(L.footers).uncheck());
        await changed(() => checkbox(L.title).uncheck());
        await changed(() => checkbox(L.hideBackground).check());
        layouts = getSlideMasterLayouts(await deck(), added);
        const edited = layouts[11];
        assert.deepEqual(
          getSlideLayoutPlaceholders(edited).map((placeholder) => placeholder.type),
          ['pic'],
        );
        assert.equal(isSlideLayoutBackgroundGraphicsHidden(edited), true);

        // Delete removes the unused layout, then the unused master; the master
        // that slides use cannot be deleted.
        await changed(() => ribbonButton(L.delete).click());
        assert.equal(getSlideMasterLayouts(await deck(), added).length, 11);
        await pane.locator('.cell.master').nth(1).click();
        await changed(() => ribbonButton(L.delete).click());
        assert.deepEqual(getSlideMasterPartNames(await deck()), [firstMaster]);
        assert.equal(await pane.locator('.cell.master').count(), 1);
        assert.equal(await ribbonButton(L.delete).isDisabled(), true);

        // Undo brings the master back; Insert Slide Master is one undo step.
        await page.keyboard.press('Meta+z');
        await page.waitForFunction(
          () => document.querySelectorAll('.master-pane .cell.master').length === 2,
        );

        // Handout Master: four slides per page, landscape, no header.
        await enter(L.handoutMaster);
        const handout = page.getByRole('region', { name: L.handoutMaster, exact: true });
        await handout.waitFor();
        await changed(() => ribbonButton(L.fourSlides).click());
        assert.equal(await handout.locator('[data-role="slide"]').count(), 4);
        assert.equal(await ribbonButton(L.fourSlides).getAttribute('aria-pressed'), 'true');
        await ribbonButton(L.handoutOrientation).click();
        await changed(() =>
          ribbon.getByRole('menuitemradio', { name: L.landscape, exact: true }).click(),
        );
        await changed(() => checkbox(L.header).uncheck());
        pres = await deck();
        assert.equal(getHandoutSlidesPerPage(pres), 4);
        const size = getNotesPageSize(pres);
        assert.ok(size.width > size.height);
        assert.deepEqual(
          getHandoutMasterPlaceholders(pres).map((placeholder) => placeholder.type),
          ['dt', 'ftr', 'sldNum'],
        );
        assert.equal(await handout.locator('[data-role="hdr"]').count(), 0);
        const sheet = await handout.locator('.page').evaluate((node) => {
          const rect = node.getBoundingClientRect();
          return rect.width / rect.height;
        });
        assert.ok(Math.abs(sheet - 10 / 7.5) < 0.01, `landscape page ${sheet}`);

        // Notes Master: the body placeholder comes off.
        await enter(L.notesMaster);
        const notes = page.getByRole('region', { name: L.notesMaster, exact: true });
        await notes.waitFor();
        assert.equal(await notes.locator('[data-role="body"]').count(), 1);
        await changed(() => checkbox(L.body).uncheck());
        assert.equal(await notes.locator('[data-role="body"]').count(), 0);
        assert.deepEqual(
          getNotesMasterPlaceholders(await deck()).map((placeholder) => placeholder.type),
          ['hdr', 'dt', 'sldImg', 'ftr', 'sldNum'],
        );
        assert.deepEqual(errors, []);
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );

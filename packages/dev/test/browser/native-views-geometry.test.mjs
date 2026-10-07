import assert from 'node:assert/strict';
import { copyFile, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { findSlideLayout, getSlideLayoutPlaceholders, loadPresentation } from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';

// Mac PowerPoint 16's views, measured through the accessibility API and from
// window screenshots in a 1512 × 900 pt window (2026-10-07). CSS px equal Mac
// points. Sizes the CSS fixes are exact everywhere; positions that follow
// font-dependent captions get a tolerance off macOS, where CI's Linux fonts
// are wider (as in native-ribbon-geometry.test.mjs).
const MAC = process.platform === 'darwin';
const near = (actual, expected, label, tolerance = MAC ? 2 : 6) =>
  assert.ok(Math.abs(actual - expected) <= tolerance, `${label}: ${actual} (native ${expected})`);

// The native capture's deck: titles on Title Only slides and a three-level
// body on a Title and Content slide, 16:9.
const DECK = `import {readFileSync} from 'node:fs';
import {getSlideShapes,setParagraphLevel,setSlideSize,SLIDE_SIZE_16_9} from '@office-kit/pptx';
import {Presentation,Slide,Fill,Shape,Raw} from '@office-kit/pptx-dsl';
const source = new Uint8Array(readFileSync(new URL('./template.pptx', import.meta.url)));
const title = (text) => <Fill target={{placeholder:{type:'title'}}}>{text}</Fill>;
const only = {name:'Title Only'};
export default <Presentation source={source} mode="compose">
  <Raw scope="presentation" apply={({presentation}) => setSlideSize(presentation, SLIDE_SIZE_16_9, {content:'fit'})} />
  <Slide layout={only}>{title('Shape slide')}<Shape preset="rect" x={2} y={3} width={4} height={2} fill="#2E5F7F" text="Sample text in shape" /></Slide>
  <Slide layout={{name:'Title and Content'}}>{title('Outline title')}<Fill target={{placeholder:{idx:1}}}>{'First level\\nSecond level\\nThird level\\nBack to first'}</Fill><Raw scope="slide" apply={({slide}) => { const body = getSlideShapes(slide)[1]; setParagraphLevel(body, 1, 1); setParagraphLevel(body, 2, 2); }} /></Slide>
  <Slide layout={only}>{title('Indents')}</Slide>
  <Slide layout={only}>{title('Rotated')}</Slide>
  <Slide layout={only}>{title('Vertical')}</Slide>
  <Slide layout={only}>{title('Picture')}</Slide>
</Presentation>;`;

const LABELS = {
  en: {
    view: 'View',
    outline: 'Outline View',
    notesPage: 'Notes Page',
    slideMaster: 'Slide Master',
    handoutMaster: 'Handout Master',
    notesMaster: 'Notes Master',
    close: 'Close Master',
    switcher: ['Normal', 'Slide Sorter', 'Reading View', 'Slide Show'],
    notes: 'Notes',
    comments: 'Comments',
    slide: (n) => `Slide ${n} of 6`,
    notesPosition: (n) => `Notes ${n} of 6`,
    masterHelp: (name) => `Currently in ${name} View.`,
    slideMasterTabs: [
      'Slide Master',
      'Home',
      'Insert',
      'Draw',
      'Transitions',
      'Animations',
      'Review',
      'View',
    ],
    pageMasterTabs: (first) => [first, 'Home', 'Insert', 'Draw', 'Review', 'View'],
    rename: 'Rename',
    delete: 'Delete',
    preserve: 'Preserve',
    insertSlideMaster: 'Insert Slide Master',
    sixSlides: '6 Slides',
    placeholder: 'Placeholder',
    readingTitle: 'Slide Show - [deck]',
    readingView: 'Reading View',
    slideLabel: 'Slide',
  },
  ja: {
    view: '表示',
    outline: 'アウトライン表示',
    notesPage: 'ノート',
    slideMaster: 'スライド マスター',
    handoutMaster: '配布資料マスター',
    notesMaster: 'ノート マスター',
    close: 'マスターを閉じる',
    switcher: ['標準', 'スライド一覧', '閲覧表示', 'スライド ショー'],
    notes: 'ノート',
    comments: 'コメント',
    slide: (n) => `スライド ${n} / 6`,
    notesPosition: (n) => `ノート ${n} / 6`,
    masterHelp: (name) => `現在のモード: ${name}表示`,
    slideMasterTabs: [
      'スライド マスター',
      'ホーム',
      '挿入',
      '描画',
      '画面切り替え',
      'アニメーション',
      '校閲',
      '表示',
    ],
    pageMasterTabs: (first) => [first, 'ホーム', '挿入', '描画', '校閲', '表示'],
    rename: '名前の変更',
    delete: '削除',
    preserve: '保持',
    insertSlideMaster: 'スライド マスターの挿入',
    sixSlides: '6 枚',
    placeholder: 'プレースホルダー',
    readingTitle: 'スライド ショー - [deck]',
    readingView: '閲覧表示',
    slideLabel: 'スライド',
  },
};

const box = (locator) =>
  locator.evaluate((node) => {
    const rect = node.getBoundingClientRect();
    return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
  });

for (const locale of ['en', 'ja'])
  test(
    `views, status bar switcher and outline follow Mac PowerPoint (${locale})`,
    { timeout: 180000 },
    async () => {
      const L = LABELS[locale];
      const dir = await mkdtemp(join(tmpdir(), 'office-native-views-'));
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
        await page.locator('.statusbar').waitFor();
        await page.locator('.nav [data-slide-index="5"]').waitFor();

        const bar = page.locator('.statusbar');
        const switcher = bar.getByRole('group', {
          name: locale === 'en' ? 'Presentation views' : 'プレゼンテーションの表示',
        });
        const viewTab = () => page.getByRole('tab', { name: L.view, exact: true }).click();
        const ribbonButton = (name) =>
          page.locator('#ribbon-panel').getByRole('button', { name, exact: true }).first();
        const enter = async (name) => {
          await viewTab();
          await ribbonButton(name).click();
        };
        // The switcher: exactly four segments, 37 + 36 + 36 + 36 = 147 pt.
        const segments = async () =>
          switcher.getByRole('button').evaluateAll((nodes) =>
            nodes.map((node) => ({
              name: node.getAttribute('aria-label'),
              pressed: node.getAttribute('aria-pressed') === 'true',
              width: node.getBoundingClientRect().width,
            })),
          );
        const pressed = async () => (await segments()).map((segment) => segment.pressed);
        const position = () => bar.locator('.position').textContent();
        const hasPaneButtons = async () =>
          (await bar.getByRole('button', { name: L.notes, exact: true }).count()) +
          (await bar.getByRole('button', { name: L.comments, exact: true }).count());
        const zoomText = async () =>
          (
            await bar
              .getByTitle(locale === 'en' ? 'Zoom...' : 'ズーム...', { exact: true })
              .innerText()
          ).trim();
        const tabs = () =>
          page
            .getByRole('tab')
            .evaluateAll((nodes) => nodes.map((node) => node.textContent.trim()));

        // Normal.
        const normalSegments = await segments();
        assert.deepEqual(
          normalSegments.map((segment) => segment.name),
          L.switcher,
        );
        assert.deepEqual(
          normalSegments.map((segment) => segment.width),
          [37, 36, 36, 36],
        );
        assert.equal((await box(switcher)).width, 147);
        assert.deepEqual(await pressed(), [true, false, false, false]);
        assert.equal(await position(), L.slide(1));
        assert.equal(await hasPaneButtons(), 2);

        // Outline View counts as Normal. Titles start 36 pt in, bullets step
        // 11.5 pt per level with the text 6 pt after them, formatted text is
        // drawn at a third of its size (the 44 pt title at 14.7 pt), and slides
        // without body text are 27 pt apart.
        await enter(L.outline);
        const outline = page.getByRole('navigation', { name: L.outline, exact: true });
        await outline.waitFor();
        assert.deepEqual(await pressed(), [true, false, false, false]);
        assert.equal(await hasPaneButtons(), 2);
        const paneLeft = (await box(outline)).x;
        const titles = await outline
          .locator('[data-outline-title]')
          .evaluateAll((nodes) => nodes.map((node) => node.getBoundingClientRect().left));
        assert.equal(titles.length, 6);
        for (const left of titles) assert.equal(left - paneLeft, 36);
        const titleSize = await outline
          .locator('[data-outline-title] span')
          .first()
          .evaluate((node) => parseFloat(getComputedStyle(node).fontSize));
        near(titleSize, 44 / 3, 'formatted title size', 0.1);
        const body = await outline
          .locator('[data-outline-slide="1"] [data-outline-paragraph]:not([data-outline-title])')
          .evaluateAll((nodes) =>
            nodes.map((node) => ({
              padding: parseFloat(getComputedStyle(node).paddingLeft),
              marker: parseFloat(getComputedStyle(node, '::before').left),
            })),
          );
        assert.deepEqual(
          body.map((paragraph) => paragraph.padding),
          [6, 17.5, 29, 6],
        );
        assert.deepEqual(
          body.map((paragraph) => paragraph.marker),
          [0, 11.5, 23, 0],
        );
        const rows = await outline
          .locator('.outline-slide')
          .evaluateAll((nodes) => nodes.map((node) => node.getBoundingClientRect().top));
        assert.deepEqual(
          rows.slice(3).map((top, index) => top - rows[index + 2]),
          [27, 27, 27],
        );

        // Show Formatting is on by default, in Mac PowerPoint's outline menu.
        if (locale === 'en') {
          await outline
            .locator('[data-outline-slide="1"] [role="textbox"]')
            .nth(1)
            .click({ button: 'right' });
          const menu = page.getByRole('menu').first();
          const items = await menu.evaluate((node) =>
            [
              ...node.querySelectorAll(
                ':scope > .ctx-item, :scope > .ctx-sep, :scope > * > .ctx-item[aria-haspopup]',
              ),
            ].map((item) =>
              item.classList.contains('ctx-sep') ? '|' : item.getAttribute('aria-label'),
            ),
          );
          assert.deepEqual(items, [
            'Cut',
            'Copy',
            'Paste',
            '|',
            'New Slide',
            'Duplicate Slide',
            'Delete Slide',
            '|',
            'Collapse',
            'Expand',
            '|',
            'Promote',
            'Demote',
            'Move Up',
            'Move Down',
            '|',
            'Thesaurus...',
            'Translate...',
            '|',
            'Show Formatting',
            '|',
            'Hyperlink...',
          ]);
          assert.equal(
            await menu
              .getByRole('menuitemcheckbox', { name: 'Show Formatting', exact: true })
              .getAttribute('aria-checked'),
            'true',
          );
          await page.keyboard.press('Escape');
        }

        // Slide Sorter at 80%: 200 pt thumbnails in 206 × 147 pt cells on a
        // 245 pt pitch, six to a row in this window, centred.
        await switcher.getByRole('button', { name: L.switcher[1], exact: true }).click();
        const sorter = page.locator('.nav.sorter');
        await sorter.waitFor();
        assert.deepEqual(await pressed(), [false, true, false, false]);
        assert.equal(await hasPaneButtons(), 0);
        assert.equal(await zoomText(), '80%');
        const list = await box(sorter);
        const cells = await sorter.locator('.thumb-row').evaluateAll((nodes) =>
          nodes.map((node) => {
            const rect = node.getBoundingClientRect();
            const image = node.querySelector('.thumb').getBoundingClientRect();
            return {
              x: rect.x,
              y: rect.y,
              width: rect.width,
              height: rect.height,
              imageX: image.x - rect.x,
              imageY: image.y - rect.y,
              imageWidth: image.width,
              imageHeight: image.height,
            };
          }),
        );
        assert.equal(cells.length, 6);
        assert.equal(new Set(cells.map((cell) => cell.y)).size, 1, 'one row of six');
        for (const cell of cells) {
          assert.deepEqual(
            [cell.width, cell.imageX, cell.imageY, cell.imageWidth],
            [206, 3, 3, 200],
          );
          near(cell.height, 147, 'sorter cell height', 1);
          near(cell.imageHeight, 112.5, 'sorter thumbnail height', 1);
        }
        assert.deepEqual(
          cells.slice(1).map((cell, index) => cell.x - cells[index].x),
          [245, 245, 245, 245, 245],
        );
        // The row is centred beside the scrollbar gutter: macOS overlay
        // scrollbars reserve none, Linux Chromium's classic ones 15 px.
        near(cells[0].x - list.x, 32, 'first sorter cell', MAC ? 3 : 8);
        near(cells[0].y - list.y, 5, 'first sorter row', 1);

        // Notes Page: a portrait page fitted with 22 pt around it, the slide
        // image where the default notes master puts it, and no segment selected.
        await enter(L.notesPage);
        const notesPage = page
          .getByRole('region', { name: L.notesPage, exact: true })
          .and(page.locator('.page-area'));
        const sheet = await box(notesPage.locator('.page'));
        assert.deepEqual(await pressed(), [false, false, false, false]);
        assert.equal(await position(), L.notesPosition(locale === 'en' ? 2 : 1));
        assert.equal(await hasPaneButtons(), 0);
        const zoom = sheet.width / 540;
        assert.equal(await zoomText(), `${Math.round(zoom * 100)}%`);
        near(zoom * 100, 96, 'notes page fit zoom', MAC ? 2 : 4);
        near(sheet.height, 720 * zoom, 'page height', 1);
        const image = await box(notesPage.locator('.slide'));
        near(image.x - sheet.x, 54 * zoom, 'slide image x', 1);
        near(image.y - sheet.y, 90 * zoom, 'slide image y', 1);
        near(image.width, 432 * zoom, 'slide image width', 1);

        // Slide Master: the Slide Master tab replaces Design, Slide Show and
        // Record; the pane shows the master (193 pt image at 20, 7 in a 119 pt
        // cell) and its eleven layouts (124 pt images at x = 90 in 75 pt cells,
        // 87 pt apart) with the current slide's layout selected.
        await page.locator('#ribbon-tab-view').click();
        await viewTab();
        await page.keyboard.press('Meta+1');
        await page.locator('.nav [data-slide-index="1"]').click();
        await enter(L.slideMaster);
        const pane = page.getByRole('navigation', {
          name: locale === 'en' ? 'Slide Master Pane' : 'スライド マスター ウィンドウ',
          exact: true,
        });
        await pane.waitFor();
        assert.deepEqual(await tabs(), L.slideMasterTabs);
        assert.equal(
          await page
            .getByRole('tab', { name: L.slideMaster, exact: true })
            .getAttribute('aria-selected'),
          'true',
        );
        assert.deepEqual(await pressed(), [false, false, false, false]);
        assert.equal(await position(), L.slideMaster);
        assert.equal(
          await bar.locator('.position').getAttribute('title'),
          L.masterHelp(L.slideMaster),
        );
        assert.equal(await hasPaneButtons(), 0);
        const paneBox = await box(pane);
        assert.equal(paneBox.x, 0);
        const master = await pane.locator('.cell.master').evaluate((node) => {
          const rect = node.getBoundingClientRect();
          const image = node.querySelector('.image').getBoundingClientRect();
          return {
            width: rect.width,
            height: rect.height,
            imageX: image.x - rect.x,
            imageY: image.y - rect.y,
            imageWidth: image.width,
          };
        });
        assert.deepEqual(
          [master.width, master.imageX, master.imageY, master.imageWidth],
          [232, 20, 7, 193],
        );
        near(master.height, 119, 'master cell height', 1);
        const layouts = await pane.locator('.cell.layout').evaluateAll((nodes) =>
          nodes.map((node) => {
            const rect = node.getBoundingClientRect();
            const image = node.querySelector('.image').getBoundingClientRect();
            return {
              y: rect.y,
              height: rect.height,
              imageX: image.x - rect.x,
              imageY: image.y - rect.y,
              imageWidth: image.width,
              current: node.getAttribute('aria-current') === 'true',
              name: node.getAttribute('aria-label'),
            };
          }),
        );
        assert.equal(layouts.length, 11);
        for (const layout of layouts) {
          assert.deepEqual([layout.imageX, layout.imageWidth], [90, 124]);
          near(layout.imageY, 2, 'layout image y', 1);
          near(layout.height, 75, 'layout cell height', 1);
        }
        for (let index = 1; index < layouts.length; index += 1)
          near(layouts[index].y - layouts[index - 1].y, 87, 'layout pitch', 1);
        assert.deepEqual(
          layouts.filter((layout) => layout.current).map((layout) => layout.name),
          ['Title and Content'],
        );
        // As natively, a layout that slides use cannot be deleted, and
        // Preserve belongs to the master.
        assert.equal(await ribbonButton(L.insertSlideMaster).isEnabled(), true);
        const deleteButton = ribbonButton(L.delete);
        assert.equal(await deleteButton.isDisabled(), true);
        assert.ok((await deleteButton.getAttribute('title')).length > 0);
        assert.equal(await ribbonButton(L.preserve).isDisabled(), true);
        assert.equal(await ribbonButton(L.rename).isEnabled(), true);
        // Placeholders move with the arrow keys and the layout keeps the change.
        const before = getSlideLayoutPlaceholders(
          findSlideLayout(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
            'Title and Content',
          ),
        );
        const revision = (await waitForState(preview.url, () => true)).revision;
        await page
          .getByRole('button', { name: `${L.placeholder} 1`, exact: true })
          .press('ArrowRight');
        await waitForState(preview.url, (state) => state.revision !== revision);
        const after = getSlideLayoutPlaceholders(
          findSlideLayout(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
            'Title and Content',
          ),
        );
        assert.equal(after[0].bounds.x - (before[0].bounds?.x ?? after[0].bounds.x - 18288), 18288);
        // The master is renamed and preserved from the same column.
        await pane.locator('.cell.master').click();
        assert.equal(await ribbonButton(L.rename).isEnabled(), true);
        assert.equal(await ribbonButton(L.preserve).isEnabled(), true);
        await ribbonButton(L.close).click();
        await pane.waitFor({ state: 'detached' });
        assert.equal(await page.getByRole('tab', { name: L.slideMaster, exact: true }).count(), 0);
        assert.deepEqual(await pressed(), [true, false, false, false]);

        // Handout Master: six slide frames between the four corner placeholders.
        await enter(L.handoutMaster);
        const handout = page.getByRole('region', { name: L.handoutMaster, exact: true });
        await handout.waitFor();
        assert.deepEqual(await tabs(), L.pageMasterTabs(L.handoutMaster));
        assert.deepEqual(await pressed(), [false, false, false, false]);
        assert.equal(await position(), L.handoutMaster);
        assert.equal(
          await bar.locator('.position').getAttribute('title'),
          L.masterHelp(L.handoutMaster),
        );
        // Measure once the fitted page has stopped moving: the page refits as
        // the ribbon switches to the Handout Master tab.
        const { handoutPage, frames } = await handout.evaluate(async (area) => {
          const measure = () => {
            const sheet = area.querySelector('.page').getBoundingClientRect();
            return {
              handoutPage: { x: sheet.x, y: sheet.y, width: sheet.width },
              frames: [...area.querySelectorAll('[data-role="slide"]')].map((node) => {
                const rect = node.getBoundingClientRect();
                return [rect.x, rect.y, rect.width, rect.height];
              }),
            };
          };
          const frame = () => new Promise((resolve) => requestAnimationFrame(resolve));
          let last = '';
          for (;;) {
            await frame();
            const current = JSON.stringify(measure());
            if (current === last) return JSON.parse(current);
            last = current;
          }
        });
        const handoutZoom = handoutPage.width / 540;
        assert.equal(frames.length, 6);
        const expected = [88.5, 297.5, 506].flatMap((y) => [38.5, 281].map((x) => [x, y]));
        frames.forEach(([x, y, width, height], index) => {
          near((x - handoutPage.x) / handoutZoom, expected[index][0], `frame ${index} x`, 1);
          near((y - handoutPage.y) / handoutZoom, expected[index][1], `frame ${index} y`, 1);
          near(width / handoutZoom, 221, `frame ${index} width`, 1);
          near(height / handoutZoom, 125, `frame ${index} height`, 1);
        });
        assert.deepEqual(
          await handout
            .locator('[data-role]')
            .evaluateAll((nodes) =>
              nodes.map((node) => node.dataset.role).filter((role) => role !== 'slide'),
            ),
          ['hdr', 'dt', 'ftr', 'sldNum'],
        );
        const six = ribbonButton(L.sixSlides);
        assert.equal(await six.getAttribute('aria-pressed'), 'true');
        assert.equal(await six.isEnabled(), true);

        // Notes Master: the slide image above five body levels.
        await enter(L.notesMaster);
        const notesMaster = page.getByRole('region', { name: L.notesMaster, exact: true });
        await notesMaster.waitFor();
        assert.deepEqual(await tabs(), L.pageMasterTabs(L.notesMaster));
        assert.equal(await position(), L.notesMaster);
        assert.deepEqual(await pressed(), [false, false, false, false]);
        assert.equal(await notesMaster.locator('[data-role="body"] p').count(), 5);
        assert.equal(await page.locator('#ribbon-panel input[type="checkbox"]').count(), 6);
        await ribbonButton(L.close).click();
        await notesMaster.waitFor({ state: 'detached' });

        // ⌘3 Notes Page, ⌥⌘1 Slide Master, ⌘1 Normal.
        await page.keyboard.press('Meta+3');
        await page
          .getByRole('region', { name: L.notesPage, exact: true })
          .and(page.locator('.page-area'))
          .waitFor();
        await page.keyboard.press('Alt+Meta+1');
        await pane.waitFor();
        await page.keyboard.press('Meta+1');
        await pane.waitFor({ state: 'detached' });

        // Reading View without a host viewer: the full window, a 32 pt title
        // bar and the slide below it; Esc returns on the slide shown last.
        await switcher.getByRole('button', { name: L.switcher[2], exact: true }).click();
        const reading = page.getByRole('dialog', { name: L.readingView, exact: true });
        await reading.waitFor();
        assert.deepEqual(await box(reading), { x: 0, y: 0, width: 1512, height: 900 });
        assert.equal((await box(reading.locator('.title-bar'))).height, 32);
        assert.equal((await reading.locator('.title').textContent()).trim(), L.readingTitle);
        const readingSlide = reading.locator('.slide');
        assert.equal(await readingSlide.getAttribute('aria-label'), `${L.slideLabel} 2`);
        assert.equal((await box(readingSlide)).width, 1512);
        await page.keyboard.press('ArrowRight');
        assert.equal(await readingSlide.getAttribute('aria-label'), `${L.slideLabel} 3`);
        await page.keyboard.press('Escape');
        await reading.waitFor({ state: 'detached' });
        assert.equal(await position(), L.slide(3));
        assert.deepEqual(errors, []);
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );

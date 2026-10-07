import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getSlideAnimations, getSlides, loadPresentation } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

// The Entrance gallery's and the Exit Effects menu's presets that the library
// writes can be chosen, together with every item of their Effect Options, in
// English and Japanese. The rest stay listed but unavailable.
const DECK = `import {Presentation,Slide,Shape} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Shape preset="rect" x={1} y={1} width={3} height={1} text="One" /><Shape preset="rect" x={5} y={1} width={3} height={1} text="Two" /></Slide></Presentation>`;

// Name → effect written and the number of Effect Options items above Sequence.
const ENTRANCE = [
  ['Appear', 'appear', 0],
  ['Blinds', 'blindsIn', 2],
  ['Checkerboard', 'checkerboardIn', 2],
  ['Dissolve In', 'dissolveIn', 0],
  ['Fly In', 'flyIn', 8],
  ['Peek In', 'peekIn', 4],
  ['Random Bars', 'randomBarsIn', 2],
  ['Shape', 'shapeIn', 6],
  ['Split', 'splitIn', 4],
  ['Strips', 'stripsIn', 4],
  ['Wedge', 'wedgeIn', 0],
  ['Wheel', 'wheelIn', 5],
  ['Wipe', 'wipeIn', 4],
  ['Fade', 'fadeIn', 0],
  ['Zoom', 'zoomIn', 0],
];

const EXIT = [
  ['Disappear', 'クリア', 'disappear', 0],
  ['Blinds', 'ブラインド', 'blindsOut', 2],
  ['Checkerboard', 'チェッカーボード', 'checkerboardOut', 2],
  ['Dissolve Out', 'ディゾルブアウト', 'dissolveOut', 0],
  ['Fly Out', 'スライドアウト', 'flyOut', 8],
  ['Peek Out', 'クロール アウト', 'peekOut', 4],
  ['Random Bars', 'ランダム ストライプ', 'randomBarsOut', 2],
  ['Shape', '図形', 'shapeOut', 6],
  ['Split', 'スプリット', 'splitOut', 4],
  ['Strips', 'ストリップス', 'stripsOut', 4],
  ['Wedge', 'くさび形', 'wedgeOut', 0],
  ['Wheel', 'ホイール', 'wheelOut', 5],
  ['Wipe', 'ワイプ', 'wipeOut', 4],
  ['Contract', 'コントラクト', null, 0],
  ['Fade', 'フェード', 'fadeOut', 0],
  ['Swivel', 'ターン', null, 0],
  ['Zoom', 'ズーム', 'zoomOut', 0],
  ['Center Revolve', 'センター リボルブ', null, 0],
  ['Boomerang', 'ブーメラン', null, 0],
  ['Bounce', 'バウンド', null, 0],
  ['Credits', 'クレジット タイトル', null, 0],
  ['Curve Down', 'カーブ (下)', null, 0],
  ['Drop', 'ドロップ', null, 0],
  ['Flip', 'フリップ', null, 0],
  ['Float', 'フロート', null, 0],
  ['Pinwheel', '風車', null, 0],
  ['Spiral Out', 'スパイラル アウト', null, 0],
  ['Basic Swivel', 'ターン (基本)', null, 0],
  ['Whip', 'ホイップ', null, 0],
];

const SEQUENCE = [
  { name: 'As One Object', disabled: false },
  { name: 'All at Once', disabled: true },
  { name: 'By Paragraph', disabled: false },
];

test(
  'every entrance and exit preset the library writes, and its Effect Options, can be chosen',
  { timeout: 300000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-animations-gallery-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(file, DECK);
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1512, height: 900 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url + '/editor');
      await page.getByText('Saved to this project', { exact: true }).waitFor();
      const panel = page.locator('#ribbon-panel');
      const animations = async () =>
        getSlideAnimations(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        );
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
      const menuItems = () =>
        page
          .getByRole('menu')
          .locator('[role^=menuitem]')
          .evaluateAll((nodes) =>
            nodes.map((node) => ({
              name: node.textContent.replace('✓', '').trim(),
              checked: node.getAttribute('aria-checked') === 'true',
              disabled: node.disabled,
            })),
          );
      const button = (name) => panel.getByRole('button', { name, exact: true });
      const effectOptions = button('Effect Options');

      // Every item above Sequence writes something the menu then shows as
      // checked; Sequence keeps PowerPoint's three, All at Once unavailable.
      const exerciseOptions = async (label, count) => {
        await effectOptions.click();
        const items = await menuItems();
        assert.equal(items.length, count + SEQUENCE.length, `${label} options`);
        assert.deepEqual(
          items.slice(count).map(({ name, disabled }) => ({ name, disabled })),
          SEQUENCE,
          `${label} Sequence`,
        );
        const own = items.slice(0, count);
        assert.ok(
          own.every((item) => !item.disabled),
          `${label} options are all enabled`,
        );
        for (const [at, item] of own.entries()) {
          if (at > 0) await effectOptions.click();
          await changed(() =>
            page.getByRole('menuitemradio', { name: item.name, exact: true }).click(),
          );
          await effectOptions.click();
          const checked = (await menuItems())
            .slice(0, count)
            .filter((entry) => entry.checked)
            .map((entry) => entry.name);
          // Shape has two sections, so one item of each is checked.
          assert.ok(checked.includes(item.name), `${label} ▸ ${item.name}: ${checked}`);
          // Toggled shut rather than dismissed with Escape, which would also
          // clear the selection the ribbon is showing.
          await effectOptions.click();
        }
      };

      await page.locator('.hit').first().click();
      await page.getByRole('tab', { name: 'Animations', exact: true }).click();

      // A fresh shape gets the preset's own default duration.
      const tile = async (name) => {
        const radio = panel.getByRole('radio', { name, exact: true });
        while ((await radio.count()) === 0) await button('Next Entrance Effects gallery').click();
        return radio;
      };
      await changed(async () => (await tile('Shape')).click());
      assert.deepEqual(
        (await animations()).map((step) => [step.effect, step.durationMs]),
        [['shapeIn', 2000]],
      );
      // Paging forward from the gallery's start again.
      while (!(await panel.getByRole('radio', { name: 'Appear', exact: true }).count()))
        await button('Previous Entrance Effects gallery').click();

      for (const [en, effect, count] of ENTRANCE) {
        const radio = await tile(en);
        assert.equal(await radio.isDisabled(), false, en);
        await changed(() => radio.click());
        const [step] = await animations();
        assert.equal(step.effect, effect, en);
        assert.equal(await radio.getAttribute('aria-checked'), 'true', en);
        await exerciseOptions(en, count);
      }

      for (const [en, , effect, count] of EXIT) {
        await button('Exit Effects').click();
        const item = page.getByRole('menuitemradio', { name: en, exact: true });
        assert.equal(await item.isDisabled(), effect === null, en);
        if (effect === null) {
          await button('Exit Effects').click();
          continue;
        }
        await changed(() => item.click());
        assert.equal((await animations())[0].effect, effect, en);
        await button('Exit Effects').click();
        assert.equal(
          await page
            .getByRole('menuitemradio', { name: en, exact: true })
            .getAttribute('aria-checked'),
          'true',
          `${en} is checked`,
        );
        await button('Exit Effects').click();
        await exerciseOptions(`${en} (exit)`, count);
      }

      // An exit's directions say where the shape goes.
      await changed(async () => {
        await button('Exit Effects').click();
        await page.getByRole('menuitemradio', { name: 'Fly Out', exact: true }).click();
      });
      await effectOptions.click();
      assert.deepEqual(
        (await menuItems()).slice(0, 8).map((item) => item.name),
        [
          'To Bottom',
          'To Bottom-Left',
          'To Left',
          'To Top-Left',
          'To Top',
          'To Top-Right',
          'To Right',
          'To Bottom-Right',
        ],
      );
      await effectOptions.click();

      // Japanese: the Exit Effects menu in PowerPoint's names, the same items
      // available, and Effect Options in Japanese.
      await page.locator('.lang select').selectOption('ja');
      await page.getByRole('tab', { name: 'アニメーション', exact: true }).click();
      await panel.getByRole('button', { name: '終了効果', exact: true }).click();
      const exitItems = await menuItems();
      assert.deepEqual(
        exitItems.map(({ name, disabled }) => ({ name, disabled })),
        EXIT.map(([, ja, effect]) => ({ name: ja, disabled: effect === null })),
      );
      await panel.getByRole('button', { name: '終了効果', exact: true }).click();
      await panel.getByRole('button', { name: '効果のオプション', exact: true }).click();
      assert.deepEqual(
        (await menuItems()).slice(0, 8).map((item) => item.name),
        ['下へ', '左下へ', '左へ', '左上へ', '上へ', '右上へ', '右へ', '右下へ'],
      );
      assert.ok((await menuItems()).slice(0, 8).every((item) => !item.disabled));
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

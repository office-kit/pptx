import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getSlideAnimations, getSlides, loadPresentation } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

// Every preset of the Entrance and Emphasis galleries and of the Exit Effects
// gallery can be chosen and writes its own effect, every item of its Effect
// Options too, and Sequence offers PowerPoint's three builds — in English and
// Japanese.
const DECK = `import {Presentation,Slide,Shape} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Shape preset="rect" x={1} y={1} width={3} height={1} text="One" /><Shape preset="rect" x={5} y={1} width={3} height={1} text="Two" /></Slide></Presentation>`;

// Name → effect written and the number of Effect Options items above Sequence,
// in the galleries' order.
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
  ['Expand', 'expandIn', 0],
  ['Fade', 'fadeIn', 0],
  ['Swivel', 'swivelIn', 0],
  ['Zoom', 'zoomIn', 0],
  ['Center Revolve', 'centerRevolveIn', 0],
  ['Float In', 'floatIn', 0],
  ['Grow & Turn', 'growTurnIn', 0],
  ['Rise Up', 'riseUpIn', 0],
  ['Spinner', 'spinnerIn', 0],
  ['Basic Zoom', 'basicZoomIn', 0],
  ['Stretch', 'stretchIn', 0],
  ['Boomerang', 'boomerangIn', 0],
  ['Bounce', 'bounceIn', 0],
  ['Credits', 'creditsIn', 0],
  ['Curve Up', 'curveUpIn', 0],
  ['Drop', 'dropIn', 0],
  ['Flip', 'flipIn', 0],
  ['Float', 'floatingIn', 0],
  ['Pinwheel', 'pinwheelIn', 0],
  ['Spiral In', 'spiralIn', 0],
  ['Basic Swivel', 'basicSwivelIn', 0],
  ['Whip', 'whipIn', 0],
];

// The colour effects offer the palette: ten theme colours, five rows of their
// tints and shades, and ten standard colours.
const PALETTE = 70;

const EMPHASIS = [
  ['Fill Color', 'fillColor', PALETTE],
  ['Font Color', 'fontColor', PALETTE],
  ['Grow/Shrink', 'growShrink', 7],
  ['Line Color', 'lineColor', PALETTE],
  ['Spin', 'spin', 6],
  ['Transparency', 'transparency', 4],
  ['Bold Flash', 'boldFlash'],
  ['Brush Color', 'brushColor', PALETTE],
  ['Complementary Color', 'complementaryColor'],
  ['Complementary Color 2', 'complementaryColor2'],
  ['Contrasting Color', 'contrastingColor'],
  ['Darken', 'darken'],
  ['Desaturate', 'desaturate'],
  ['Lighten', 'lighten'],
  ['Object Color', 'objectColor', PALETTE],
  ['Pulse', 'pulse'],
  ['Underline', 'underline'],
  ['Color Pulse', 'colorPulse', PALETTE],
  ['Grow With Color', 'growWithColor', PALETTE],
  ['Shimmer', 'shimmer'],
  ['Teeter', 'teeter'],
  ['Blink', 'blink'],
  ['Bold Reveal', 'boldReveal'],
  ['Wave', 'wave'],
];
// PowerPoint writes no build for these, so they have no Sequence either.
const UNBUILT = new Set(['fillColor', 'lineColor']);

const EXIT = [
  ['Blinds', 'ブラインド', 'blindsOut', 2],
  ['Checkerboard', 'チェッカーボード', 'checkerboardOut', 2],
  ['Disappear', 'クリア', 'disappear', 0],
  ['Dissolve Out', 'ディゾルブアウト', 'dissolveOut', 0],
  ['Fly Out', 'スライドアウト', 'flyOut', 8],
  ['Peek Out', 'ピークアウト', 'peekOut', 4],
  ['Random Bars', 'ランダムストライプ', 'randomBarsOut', 2],
  ['Shape', '図形', 'shapeOut', 6],
  ['Split', 'スプリット', 'splitOut', 4],
  ['Strips', 'ストリップ', 'stripsOut', 4],
  ['Wedge', 'くさび形', 'wedgeOut', 0],
  ['Wheel', 'ホイール', 'wheelOut', 5],
  ['Wipe', 'ワイプ', 'wipeOut', 4],
  ['Contract', 'コントラクト', 'contractOut', 0],
  ['Fade', 'フェード', 'fadeOut', 0],
  ['Swivel', 'ターン', 'swivelOut', 0],
  ['Zoom', 'ズーム', 'zoomOut', 0],
  ['Center Revolve', 'リボルブ', 'centerRevolveOut', 0],
  ['Collapse', 'コラプス', 'collapseOut', 0],
  ['Float Out', 'フロートアウト', 'floatOut', 0],
  ['Shrink & Turn', '縮小および回転', 'shrinkTurnOut', 0],
  ['Sink Down', 'シンク', 'sinkDownOut', 0],
  ['Spinner', 'スピナー', 'spinnerOut', 0],
  ['Basic Zoom', 'ベーシック ズーム', 'basicZoomOut', 0],
  ['Stretchy', 'ゴム', 'stretchyOut', 0],
  ['Boomerang', 'ブーメラン', 'boomerangOut', 0],
  ['Bounce', 'バウンド', 'bounceOut', 0],
  ['Credits', 'クレジット タイトル', 'creditsOut', 0],
  ['Curve Down', 'カーブ (下)', 'curveDownOut', 0],
  ['Drop', 'ドロップ', 'dropOut', 0],
  ['Flip', 'フリップ', 'flipOut', 0],
  ['Float', 'フロート', 'floatingOut', 0],
  ['Pinwheel', 'ピンウィール', 'pinwheelOut', 0],
  ['Spiral Out', 'スパイラルアウト', 'spiralOut', 0],
  ['Basic Swivel', 'ベーシック ターン', 'basicSwivelOut', 0],
  ['Whip', 'ホイップ', 'whipOut', 0],
];

const SEQUENCE = ['As One Object', 'All at Once', 'By Paragraph'];
const SEQUENCE_JA = ['1 つのオブジェクトとして', 'すべて同時', '段落別'];

test(
  'every gallery preset, its Effect Options and its Sequence can be chosen',
  { timeout: 600000 },
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
              // A palette swatch has no text, only its colour's name.
              name: (node.getAttribute('aria-label') ?? node.textContent).replace('✓', '').trim(),
              checked: node.getAttribute('aria-checked') === 'true',
              disabled: node.disabled,
            })),
          );
      const button = (name) => panel.getByRole('button', { name, exact: true });
      const effectOptions = button('Effect Options');

      // Every item above Sequence writes something the menu then shows as
      // checked; Sequence follows with PowerPoint's three builds.
      const exerciseOptions = async (label, effect, count) => {
        const sequence = UNBUILT.has(effect) ? [] : SEQUENCE;
        if (count === 0 && sequence.length === 0) {
          assert.equal(await effectOptions.isDisabled(), true, `${label} has no options`);
          return;
        }
        await effectOptions.click();
        const items = await menuItems();
        assert.deepEqual(
          items.slice(count).map(({ name, disabled }) => ({ name, disabled })),
          sequence.map((name) => ({ name, disabled: false })),
          `${label} Sequence`,
        );
        assert.equal(items.length, count + sequence.length, `${label} options`);
        const own = items.slice(0, count);
        assert.ok(
          own.every((item) => !item.disabled),
          `${label} options are all enabled`,
        );
        await effectOptions.click();
        // Every swatch of a palette writes the same way, so a theme colour, a
        // tint and a standard colour stand for the seventy.
        const exercised = count === PALETTE ? [own[5], own[25], own[count - 1]] : own;
        for (const item of exercised) {
          await effectOptions.click();
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

      const tile = async (gallery, name) => {
        const radio = panel
          .getByRole('radiogroup', { name: gallery, exact: true })
          .getByRole('radio', { name, exact: true });
        // Back to the gallery's start, then forward a page at a time.
        const previous = button(`Previous ${gallery} gallery`);
        while ((await radio.count()) === 0 && (await previous.count()) > 0) await previous.click();
        while ((await radio.count()) === 0) await button(`Next ${gallery} gallery`).click();
        return radio;
      };

      // A fresh shape gets the preset's own default duration.
      await changed(async () => (await tile('Entrance Effects', 'Shape')).click());
      assert.deepEqual(
        (await animations()).map((step) => [step.effect, step.durationMs]),
        [['shapeIn', 2000]],
      );

      for (const [gallery, list] of [
        ['Entrance Effects', ENTRANCE],
        ['Emphasis Effects', EMPHASIS],
      ]) {
        for (const [en, effect, count = 0] of list) {
          const radio = await tile(gallery, en);
          assert.equal(await radio.isDisabled(), false, en);
          await changed(() => radio.click());
          const steps = await animations();
          assert.deepEqual(
            steps.map((step) => step.effect),
            [effect],
            en,
          );
          assert.equal(await radio.getAttribute('aria-checked'), 'true', en);
          await exerciseOptions(en, effect, count);
        }
      }

      const exitGallery = () => panel.getByRole('dialog', { name: 'Exit Effects', exact: true });
      for (const [en, , effect, count] of EXIT) {
        await button('Exit Effects').click();
        const item = exitGallery().getByRole('radio', { name: en, exact: true });
        assert.equal(await item.isDisabled(), false, en);
        await changed(() => item.click());
        assert.equal((await animations())[0].effect, effect, en);
        await button('Exit Effects').click();
        assert.equal(
          await exitGallery()
            .getByRole('radio', { name: en, exact: true })
            .getAttribute('aria-checked'),
          'true',
          `${en} is checked`,
        );
        await button('Exit Effects').click();
        await exerciseOptions(`${en} (exit)`, effect, count);
      }

      // Sequence: each build is written and then shown as checked.
      await changed(async () => (await tile('Entrance Effects', 'Fly In')).click());
      for (const [name, build] of [
        ['All at Once', 'allAtOnce'],
        ['By Paragraph', 'byParagraph'],
        ['As One Object', 'asOneObject'],
      ]) {
        await effectOptions.click();
        await changed(() => page.getByRole('menuitemradio', { name, exact: true }).click());
        assert.deepEqual(
          (await animations()).map((step) => [step.effect, step.build]),
          [['flyIn', build]],
          name,
        );
        await effectOptions.click();
        const checked = (await menuItems()).filter((entry) => entry.checked).map((e) => e.name);
        assert.ok(checked.includes(name), `${name}: ${checked}`);
        await effectOptions.click();
      }

      // An exit's directions say where the shape goes.
      await changed(async () => {
        await button('Exit Effects').click();
        await exitGallery().getByRole('radio', { name: 'Fly Out', exact: true }).click();
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

      // Japanese: the Exit Effects gallery in PowerPoint's names and group
      // headings, all available, and Effect Options (Sequence included) in
      // Japanese.
      await page.locator('.lang select').selectOption('ja');
      await page.getByRole('tab', { name: 'アニメーション', exact: true }).click();
      await panel.getByRole('button', { name: '終了効果', exact: true }).click();
      const jaGallery = panel.getByRole('dialog', { name: '終了効果', exact: true });
      assert.deepEqual(
        await jaGallery
          .getByRole('radiogroup')
          .evaluateAll((groups) => groups.map((group) => group.getAttribute('aria-label'))),
        ['基本', '弱', '中', 'はなやか'],
      );
      assert.deepEqual(
        await jaGallery.getByRole('radio').evaluateAll((radios) =>
          radios.map((radio) => ({
            name: radio.getAttribute('aria-label'),
            disabled: radio.disabled,
          })),
        ),
        EXIT.map(([, ja]) => ({ name: ja, disabled: false })),
      );
      await panel.getByRole('button', { name: '終了効果', exact: true }).click();
      await panel.getByRole('button', { name: '効果のオプション', exact: true }).click();
      const options = await menuItems();
      assert.deepEqual(
        options.slice(0, 8).map((item) => item.name),
        ['下へ', '左下へ', '左へ', '左上へ', '上へ', '右上へ', '右へ', '右下へ'],
      );
      assert.deepEqual(
        options.slice(8).map((item) => item.name),
        SEQUENCE_JA,
      );
      assert.ok(options.every((item) => !item.disabled));
      await changed(() =>
        page.getByRole('menuitemradio', { name: 'すべて同時', exact: true }).click(),
      );
      assert.equal((await animations())[0].build, 'allAtOnce');
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

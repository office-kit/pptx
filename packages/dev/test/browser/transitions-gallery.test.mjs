import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getSlides, getSlideTransition, loadPresentation } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

// Every one of the reference desktop app's (Mac, 16) 49 transitions can be chosen, and every
// Effect Options item of every one of them, in English and Japanese. Each tile
// writes the effect the reference desktop app writes for it, with its own default duration
// (site/test/transition-defaults.test.mjs compares the XML with the reference desktop app's).
const DECK = `import {Presentation,Slide,Shape} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Shape preset="rect" x={1} y={1} width={3} height={1} text="One" /></Slide><Slide><Shape preset="rect" x={1} y={1} width={3} height={1} text="Two" /></Slide></Presentation>`;

// Tile → what it writes: the effect element (and the option that names the
// tile), its duration and the number of Effect Options items.
const TILES = [
  ['None', 'なし', { effect: 'none' }, null, 0],
  ['Morph', '変形', { effect: 'morph', morphOption: 'byObject' }, 2000, 3],
  ['Fade', 'フェード', { effect: 'fade' }, 700, 2],
  ['Push', 'プッシュ', { effect: 'push', direction: 'u' }, 1000, 4],
  ['Wipe', 'ワイプ', { effect: 'wipe' }, 1000, 8],
  ['Split', 'スプリット', { effect: 'split', orientation: 'vert' }, 1500, 4],
  ['Reveal', '出現', { effect: 'reveal' }, 3400, 4],
  ['Cut', 'カット', { effect: 'cut' }, 100, 2],
  ['Random Bars', 'ランダムストライプ', { effect: 'randomBar', direction: 'vert' }, 1000, 2],
  ['Shape', '図形', { effect: 'circle' }, 800, 5],
  ['Uncover', 'アンカバー', { effect: 'pull' }, 750, 8],
  ['Cover', 'カバー', { effect: 'cover' }, 1000, 8],
  ['Flash', 'フラッシュ', { effect: 'flash' }, 1000, 0],
  ['Fall Over', 'フォール オーバー', { effect: 'prstTrans', preset: 'fallOver' }, 2000, 2],
  ['Drape', '垂れ幕', { effect: 'prstTrans', preset: 'drape' }, 2000, 2],
  ['Curtains', 'カーテン', { effect: 'prstTrans', preset: 'curtains' }, 6000, 0],
  ['Wind', '風', { effect: 'prstTrans', preset: 'wind' }, 2000, 2],
  ['Prestige', 'プレステージ', { effect: 'prstTrans', preset: 'prestige' }, 2000, 0],
  ['Fracture', '破砕', { effect: 'prstTrans', preset: 'fracture' }, 2000, 0],
  ['Crush', 'クシャクシャ', { effect: 'prstTrans', preset: 'crush' }, 2000, 0],
  ['Peel Off', 'ピール オフ', { effect: 'prstTrans', preset: 'peelOff' }, 1250, 2],
  ['Page Curl', 'ページ カール', { effect: 'prstTrans', preset: 'pageCurlDouble' }, 1250, 4],
  ['Airplane', '飛行機', { effect: 'prstTrans', preset: 'airplane' }, 1250, 2],
  ['Origami', '折り紙', { effect: 'prstTrans', preset: 'origami' }, 3250, 2],
  ['Dissolve', 'ディゾルブ', { effect: 'dissolve' }, 1200, 0],
  ['Checkerboard', 'チェッカーボード', { effect: 'checker' }, 2500, 2],
  ['Blinds', 'ブラインド', { effect: 'blinds', direction: 'vert' }, 1600, 2],
  ['Clock', '時計', { effect: 'wheel', spokes: 1 }, 1000, 3],
  ['Ripple', 'さざ波', { effect: 'ripple' }, 1400, 5],
  ['Honeycomb', 'ハチの巣', { effect: 'honeycomb' }, 4400, 0],
  ['Glitter', 'キラキラ', { effect: 'glitter', pattern: 'hexagon' }, 3900, 8],
  ['Vortex', '渦巻き', { effect: 'vortex', direction: 'r' }, 4000, 4],
  ['Shred', '細分', { effect: 'shred' }, 3000, 4],
  ['Switch', '切り替え', { effect: 'switch', direction: 'r' }, 1250, 2],
  ['Flip', 'フリップ', { effect: 'flip', direction: 'r' }, 1250, 2],
  ['Gallery', 'ギャラリー', { effect: 'gallery', direction: 'l' }, 1600, 2],
  ['Cube', 'キューブ', { effect: 'prism' }, 1200, 4],
  ['Doors', '扉', { effect: 'doors', direction: 'vert' }, 1400, 2],
  ['Box', 'ボックス', { effect: 'prism', isInverted: true }, 1600, 4],
  ['Comb', 'コーム', { effect: 'comb' }, 1000, 2],
  ['Zoom', 'ズーム', { effect: 'warp', direction: 'in' }, 900, 3],
  ['Random', 'ランダム', { effect: 'random' }, 1500, 0],
  ['Pan', 'パン', { effect: 'pan', direction: 'u' }, 1300, 4],
  ['Ferris Wheel', '観覧車', { effect: 'ferris', direction: 'l' }, 2000, 2],
  ['Conveyor', 'コンベヤ', { effect: 'conveyor', direction: 'l' }, 1600, 2],
  ['Rotate', '回転', { effect: 'prism', isContent: true }, 2000, 4],
  ['Window', '窓', { effect: 'window', direction: 'vert' }, 1500, 2],
  ['Orbit', 'オービット', { effect: 'prism', isContent: true, isInverted: true }, 1600, 4],
  ['Fly Through', 'フライスルー', { effect: 'flythrough' }, 800, 4],
];

// How the library writes a duration (the reference desktop app's rule, see
// NATIVE_PARITY.md): `spd` is the fastest speed at least that long, left
// out when fast, and `p14:dur` only when the speed does not already say it.
const SPEED_MS = { fast: 500, med: 750, slow: 1000 };
const timingRead = (durationMs) => {
  const speed = durationMs <= 500 ? 'fast' : durationMs <= 750 ? 'med' : 'slow';
  return {
    ...(speed === 'fast' ? {} : { speed }),
    ...(SPEED_MS[speed] === durationMs ? {} : { durationMs }),
  };
};

test(
  'every native transition and Effect Options item can be chosen, in English and Japanese',
  { timeout: 300000 },
  async () => {
    assert.equal(TILES.length, 49);
    const dir = await mkdtemp(join(tmpdir(), 'office-transitions-gallery-'));
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
      const deck = async () =>
        loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
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
      // Pages the gallery forward until the tile is in view; the tiles are
      // visited in gallery order, so forward is always the way.
      const tileNamed = async (gallery, name) => {
        const radio = panel.getByRole('radio', { name, exact: true });
        while ((await radio.count()) === 0)
          await panel.getByRole('button', { name: `Next ${gallery} gallery` }).click();
        return radio;
      };
      const effectOptions = panel.getByRole('button', { name: 'Effect Options', exact: true });

      await page.getByRole('tab', { name: 'Transitions', exact: true }).click();
      for (const [en, , written, durationMs, optionCount] of TILES) {
        const radio = await tileNamed('Transition Styles', en);
        assert.equal(await radio.isDisabled(), false, en);
        // Choosing None on a slide without a transition changes nothing.
        if (en === 'None') continue;
        await changed(() => radio.click());
        const transition = getSlideTransition(getSlides(await deck())[0]);
        assert.deepEqual(transition, { ...written, ...timingRead(durationMs) }, en);
        assert.equal(await radio.getAttribute('aria-checked'), 'true', en);
        assert.equal(
          await panel.getByRole('spinbutton', { name: 'Duration:' }).inputValue(),
          (durationMs / 1000).toFixed(2),
          en,
        );
        assert.equal(await effectOptions.isDisabled(), optionCount === 0, `${en} Effect Options`);
        if (optionCount === 0) continue;
        await effectOptions.click();
        const items = await menuItems();
        assert.equal(items.length, optionCount, `${en} options`);
        assert.ok(
          items.every((item) => !item.disabled),
          `${en} options are all enabled`,
        );
        assert.equal(items.filter((item) => item.checked).length, 1, `${en}: one option checked`);
        // Every item writes something the menu then shows as checked.
        for (const [at, item] of items.entries()) {
          if (at > 0) await effectOptions.click();
          await page.getByRole('menuitemradio', { name: item.name, exact: true }).click();
          await effectOptions.click();
          const after = await menuItems();
          assert.deepEqual(
            after.filter((entry) => entry.checked).map((entry) => entry.name),
            [item.name],
            `${en} ▸ ${item.name}`,
          );
          await page.keyboard.press('Escape');
        }
        // The applied effect stays on its tile, with its duration kept.
        assert.equal(
          await panel.getByRole('spinbutton', { name: 'Duration:' }).inputValue(),
          (durationMs / 1000).toFixed(2),
          `${en} keeps its duration`,
        );
        assert.equal(await radio.getAttribute('aria-checked'), 'true', `${en} stays checked`);
      }

      // A 2010+ effect is saved with the option chosen for it.
      await changed(async () => {
        const radio = panel.getByRole('radio', { name: 'Vortex', exact: true });
        while ((await radio.count()) === 0)
          await panel.getByRole('button', { name: 'Previous Transition Styles gallery' }).click();
        await radio.click();
        await effectOptions.click();
        await page.getByRole('menuitemradio', { name: 'From Bottom', exact: true }).click();
      });
      assert.deepEqual(getSlideTransition(getSlides(await deck())[0]), {
        effect: 'vortex',
        direction: 'u',
        speed: 'slow',
        durationMs: 4000,
      });

      // Japanese: the same 49 tiles, all enabled, with the reference desktop app's (Mac) names,
      // and Effect Options in Japanese.
      await page.locator('.lang select').selectOption('ja');
      await page.getByRole('tab', { name: '画面切り替え', exact: true }).click();
      const gallery = panel.getByRole('radiogroup', { name: '画面切り替え効果', exact: true });
      const previous = panel.getByRole('button', { name: /画面切り替え効果 ギャラリーの前の行/ });
      while (!(await gallery.getByRole('radio', { name: 'なし', exact: true }).count()))
        await previous.click();
      const seen = [];
      for (;;) {
        for (const radio of await gallery.getByRole('radio').all()) {
          const name = await radio.getAttribute('aria-label');
          if (!seen.includes(name)) seen.push(name);
          assert.equal(await radio.isDisabled(), false, name);
        }
        const next = panel.getByRole('button', { name: /画面切り替え効果 ギャラリーの次の行/ });
        if (await next.isDisabled()) break;
        await next.click();
      }
      assert.deepEqual(
        seen,
        TILES.map(([, ja]) => ja),
      );
      // The slide still has Vortex, whose options read in Japanese.
      await panel.getByRole('button', { name: '効果のオプション', exact: true }).click();
      assert.deepEqual(
        (await menuItems()).map((item) => item.name),
        ['左から', '上から', '右から', '下から'],
      );
      assert.ok((await menuItems()).every((item) => !item.disabled));
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

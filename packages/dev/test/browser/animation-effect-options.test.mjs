import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getSlideAnimations, getSlides, loadPresentation } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

// The emphasis effects' own Effect Options (Spin's direction and amount,
// Grow/Shrink's direction and size, Transparency's amount, the colour of the
// colour effects) in the ribbon menu and the animation pane, in English and
// Japanese; a new preset taking its own duration; and the Emphasis Effects
// gallery a narrow ribbon collapses to.
const DECK = `import {Presentation,Slide,Shape} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Shape preset="rect" x={1} y={1} width={3} height={1} text="One" /></Slide></Presentation>`;

const open = async (width = 1512) => {
  const dir = await mkdtemp(join(tmpdir(), 'office-effect-options-'));
  const file = join(dir, 'deck.tsx');
  await writeFile(file, DECK);
  const preview = await startPreview(file);
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({ viewport: { width, height: 900 } });
  const errors = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto(preview.url + '/editor');
  await page.getByText('Saved to this project', { exact: true }).waitFor();
  const panel = page.locator('#ribbon-panel');
  const steps = async () =>
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
  const button = (name) => panel.getByRole('button', { name, exact: true });
  const menuItems = () =>
    page
      .getByRole('menu')
      .locator('[role^=menuitem]')
      .evaluateAll((nodes) =>
        nodes.map((node) => ({
          name: (node.getAttribute('aria-label') ?? node.textContent).replace('✓', '').trim(),
          checked: node.getAttribute('aria-checked') === 'true',
        })),
      );
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
  const close = async () => {
    await browser.close();
    await preview.close();
    await rm(dir, { recursive: true, force: true });
  };
  await page.locator('.hit').first().click();
  await page.getByRole('tab', { name: 'Animations', exact: true }).click();
  return { page, panel, errors, steps, changed, button, menuItems, tile, close };
};

test(
  'the emphasis effects offer their own Effect Options and write them',
  { timeout: 300000 },
  async () => {
    const ui = await open();
    const { page, panel, steps, changed, button, menuItems, tile } = ui;
    try {
      const choose = async (name) => {
        await button('Effect Options').click();
        await changed(() => page.getByRole('menuitemradio', { name, exact: true }).click());
      };
      const checked = async () => {
        await button('Effect Options').click();
        const names = (await menuItems()).filter((item) => item.checked).map((item) => item.name);
        await button('Effect Options').click();
        return names;
      };

      await changed(async () => (await tile('Emphasis Effects', 'Spin')).click());
      await button('Effect Options').click();
      assert.deepEqual(
        (await menuItems()).map((item) => item.name),
        [
          'Clockwise',
          'Counterclockwise',
          'Quarter Spin',
          'Half Spin',
          'Full Spin',
          'Two Spins',
          'As One Object',
          'All at Once',
          'By Paragraph',
        ],
      );
      assert.deepEqual(
        await panel
          .getByRole('menu', { name: 'Effect Options' })
          .locator('.heading')
          .allTextContents(),
        ['Direction', 'Amount', 'Sequence'],
      );
      await button('Effect Options').click();
      await choose('Counterclockwise');
      await choose('Quarter Spin');
      let [step] = await steps();
      assert.deepEqual(
        [step.effect, step.spinDirection, step.spinDegrees],
        ['spin', 'counterclockwise', 90],
      );
      assert.deepEqual(await checked(), ['Counterclockwise', 'Quarter Spin', 'As One Object']);

      await changed(async () => (await tile('Emphasis Effects', 'Grow/Shrink')).click());
      await choose('Horizontal');
      await choose('Huge');
      [step] = await steps();
      assert.deepEqual([step.scaleDirection, step.scalePercent], ['horizontal', 400]);
      assert.deepEqual(await checked(), ['Horizontal', 'Huge', 'As One Object']);

      await changed(async () => (await tile('Emphasis Effects', 'Transparency')).click());
      await choose('75%');
      [step] = await steps();
      assert.equal(step.transparencyPercent, 75);

      // A colour effect's options are the palette: a standard colour, then a
      // tint of a theme colour, each written and shown as the one chosen.
      await changed(async () => (await tile('Emphasis Effects', 'Fill Color')).click());
      await button('Effect Options').click();
      const palette = page.getByRole('menu', { name: 'Effect Options' });
      assert.deepEqual(
        await palette
          .getByRole('group')
          .evaluateAll((groups) => groups.map((g) => g.getAttribute('aria-label'))),
        ['Theme Colors', 'Standard Colors'],
      );
      await button('Effect Options').click();
      await choose('Red');
      [step] = await steps();
      assert.equal(step.color, '#FF0000');
      assert.deepEqual(await checked(), ['Red']);
      await choose('Accent 1, Lighter 80%');
      [step] = await steps();
      assert.deepEqual(step.color, {
        color: 'scheme:accent1',
        colorTransforms: [
          { kind: 'lumMod', value: 0.2 },
          { kind: 'lumOff', value: 0.8 },
        ],
      });
      assert.deepEqual(await checked(), ['Accent 1, Lighter 80%']);
      assert.deepEqual(ui.errors, []);
    } finally {
      await ui.close();
    }
  },
);

test(
  'a new preset takes its own duration; an option keeps the one set',
  { timeout: 300000 },
  async () => {
    const ui = await open();
    const { panel, steps, changed, button, tile } = ui;
    try {
      const duration = panel.getByRole('spinbutton', { name: 'Duration', exact: true });
      await changed(async () => (await tile('Entrance Effects', 'Fade')).click());
      assert.equal(await duration.inputValue(), '0.50');
      await changed(async () => {
        await duration.fill('3');
        await duration.press('Enter');
      });
      assert.equal((await steps())[0].durationMs, 3000);
      // Shape's default is two seconds, whatever the effect before it ran for.
      await changed(async () => (await tile('Entrance Effects', 'Shape')).click());
      assert.equal((await steps())[0].durationMs, 2000);
      assert.equal(await duration.inputValue(), '2.00');
      await changed(async () => {
        await duration.fill('4');
        await duration.press('Enter');
      });
      // An option of the same preset keeps it.
      await button('Effect Options').click();
      await changed(() =>
        ui.page.getByRole('menuitemradio', { name: 'Diamond', exact: true }).click(),
      );
      assert.deepEqual(
        (await steps()).map((step) => [step.effect, step.shape, step.durationMs]),
        [['shapeIn', 'diamond', 4000]],
      );
      // So does an exit picked from the gallery: its own default.
      await button('Exit Effects').click();
      await changed(() =>
        panel
          .getByRole('dialog', { name: 'Exit Effects', exact: true })
          .getByRole('radio', { name: 'Bounce', exact: true })
          .click(),
      );
      assert.equal((await steps())[0].durationMs, 2000);
      assert.deepEqual(ui.errors, []);
    } finally {
      await ui.close();
    }
  },
);

test('the animation pane offers the same options', { timeout: 300000 }, async () => {
  const ui = await open();
  const { page, steps, changed, tile } = ui;
  try {
    await changed(async () => (await tile('Emphasis Effects', 'Spin')).click());
    await ui.button('Animation Pane').click();
    const pane = page.locator('section[data-animation-pane]');
    await changed(() =>
      pane
        .getByRole('combobox', { name: 'Amount 1', exact: true })
        .selectOption({ label: 'Half Spin' }),
    );
    await changed(() =>
      pane
        .getByRole('combobox', { name: 'Direction 1', exact: true })
        .selectOption({ label: 'Counterclockwise' }),
    );
    let [step] = await steps();
    assert.deepEqual([step.spinDirection, step.spinDegrees], ['counterclockwise', 180]);

    await changed(async () => (await tile('Emphasis Effects', 'Line Color')).click());
    await pane.getByRole('button', { name: 'Color 1', exact: true }).click();
    await changed(() => page.getByRole('menuitemradio', { name: 'Blue', exact: true }).click());
    [step] = await steps();
    assert.deepEqual([step.effect, step.color], ['lineColor', '#0070C0']);
    assert.deepEqual(ui.errors, []);
  } finally {
    await ui.close();
  }
});

test('Japanese: the emphasis options and the gallery headings', { timeout: 300000 }, async () => {
  const ui = await open();
  const { page, panel, changed, menuItems, tile } = ui;
  try {
    await changed(async () => (await tile('Emphasis Effects', 'Grow/Shrink')).click());
    await page.locator('.lang select').selectOption('ja');
    await page.getByRole('tab', { name: 'アニメーション', exact: true }).click();
    await panel.getByRole('button', { name: '効果のオプション', exact: true }).click();
    assert.deepEqual(
      (await menuItems()).map((item) => item.name),
      [
        '両方向',
        '水平方向',
        '垂直方向',
        '極小',
        '小',
        '大',
        '特大',
        '1 つのオブジェクトとして',
        'すべて同時',
        '段落別',
      ],
    );
    assert.deepEqual(
      await panel
        .getByRole('menu', { name: '効果のオプション' })
        .locator('.heading')
        .allTextContents(),
      ['方向', '量', '連続'],
    );
    await panel.getByRole('button', { name: '効果のオプション', exact: true }).click();
    // Spin's, too.
    await changed(async () =>
      panel
        .getByRole('radiogroup', { name: '強調効果', exact: true })
        .getByRole('radio', { name: 'スピン', exact: true })
        .click(),
    );
    await panel.getByRole('button', { name: '効果のオプション', exact: true }).click();
    assert.deepEqual(
      (await menuItems()).slice(0, 6).map((item) => item.name),
      ['時計回り', '反時計回り', '4 分の 1 回転', '半回転', '1 回転', '2 回転'],
    );
    assert.deepEqual(ui.errors, []);
  } finally {
    await ui.close();
  }
});

test(
  'a narrow ribbon collapses Emphasis Effects into the same gallery popover',
  { timeout: 300000 },
  async () => {
    const ui = await open(1200);
    const { panel, steps, changed, button } = ui;
    try {
      await button('Emphasis Effects').click();
      const gallery = panel.getByRole('dialog', { name: 'Emphasis Effects', exact: true });
      assert.deepEqual(
        await gallery
          .getByRole('radiogroup')
          .evaluateAll((groups) => groups.map((g) => g.getAttribute('aria-label'))),
        ['Basic', 'Subtle', 'Moderate', 'Exciting'],
      );
      assert.equal(await gallery.getByRole('radio').count(), 24);
      // PowerPoint's 387 pt popover of 76 × 90 pt tiles, five to a row.
      const box = await gallery.boundingBox();
      assert.equal(Math.round(box.width), 387);
      const tiles = await gallery.getByRole('radio').evaluateAll((radios) =>
        radios.slice(0, 6).map((radio) => {
          const r = radio.getBoundingClientRect();
          return [Math.round(r.width), Math.round(r.height), Math.round(r.top)];
        }),
      );
      assert.deepEqual(
        tiles.map(([w, h]) => [w, h]),
        Array.from({ length: 6 }, () => [76, 90]),
      );
      assert.equal(new Set(tiles.slice(0, 5).map(([, , top]) => top)).size, 1, 'five to a row');
      await changed(() => gallery.getByRole('radio', { name: 'Teeter', exact: true }).click());
      assert.equal((await steps())[0].effect, 'teeter');
      assert.deepEqual(ui.errors, []);
    } finally {
      await ui.close();
    }
  },
);

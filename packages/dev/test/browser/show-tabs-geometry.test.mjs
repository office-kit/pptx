import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getSlides, getSlideTransition, loadPresentation } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

// Transitions, Animations, Slide Show, Record, Review and View measured from
// Mac PowerPoint 16 through the accessibility API in 1512 × 900 and
// 1200 × 900 pt windows (2026-10-07). CSS px equal Mac points; button widths
// are the AX widths, as for the Home ribbon in native-geometry.test.mjs.
const DECK = `import {Presentation,Slide,Shape} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Shape preset="rect" x={4} y={4} width={3} height={1.5} text="One" /></Slide><Slide><Shape preset="rect" x={1} y={1} width={3} height={1} text="Two" /></Slide></Presentation>`;

const WIDTHS = {
  Transitions: { Preview: 46, 'Effect Options': 50, 'Apply To All': 38 },
  Animations: {
    Preview: 50,
    'Exit Effects': 50,
    'Path Animation': 57,
    'Effect Options': 50,
    'Animation Pane': 57,
    Trigger: 50,
    'Animation Painter': 57,
  },
  'Slide Show': {
    'Play from Start': 54,
    'Play from Current Slide': 73,
    'Presenter View': 55,
    'Custom Show': 50,
    'Rehearse with Coach': 63,
    'Set Up Slide Show': 62,
    'Hide Slide': 38,
    'Rehearse Timings': 53,
    Record: 50,
  },
  Record: {
    Cameo: 42,
    'From Beginning': 57,
    'From Current Slide': 73,
    'Clear Recording': 58,
    'Reset to Cameo': 50,
  },
  Review: {
    Spelling: 47,
    Thesaurus: 59,
    'Check Accessibility': 70,
    Translate: 53,
    Language: 56,
    'Mark All as Read': 47,
    'Show Changes': 51,
    'New Comment': 55,
    Delete: 50,
    Previous: 50,
    Next: 38,
    'Show Comments': 61,
    'Always Open Read-Only': 71,
    'Restrict Permission': 62,
    'Hide Ink': 50,
  },
  View: {
    Normal: 43,
    'Outline View': 43,
    'Slide Sorter': 38,
    'Notes Page': 38,
    'Reading View': 47,
    'Slide Master': 41,
    'Handout Master': 49,
    'Notes Master': 41,
    Notes: 38,
    Zoom: 38,
    'Fit to Window': 47,
    Macros: 43,
  },
};

// The command order, with `|` where PowerPoint draws a rule between groups.
const ORDER = {
  'Slide Show': [
    'Play from Start',
    'Play from Current Slide',
    'Presenter View',
    'Custom Show',
    '|',
    'Rehearse with Coach',
    '|',
    'Set Up Slide Show',
    'Hide Slide',
    'Rehearse Timings',
    'Record',
    '|',
    'Subtitle Settings',
  ],
  Record: [
    'Cameo',
    '|',
    'From Beginning',
    'From Current Slide',
    '|',
    'Clear Recording',
    'Reset to Cameo',
    '|',
    'Learn More',
  ],
  Review: [
    'Spelling',
    'Thesaurus',
    '|',
    'Check Accessibility',
    '|',
    'Translate',
    'Language',
    '|',
    'Mark All as Read',
    'Show Changes',
    '|',
    'New Comment',
    'Delete',
    'Previous',
    'Next',
    'Show Comments',
    '|',
    'Always Open Read-Only',
    'Restrict Permission',
    '|',
    'Hide Ink',
  ],
  View: [
    'Normal',
    'Outline View',
    'Slide Sorter',
    'Notes Page',
    'Reading View',
    '|',
    'Slide Master',
    'Handout Master',
    'Notes Master',
    '|',
    'Notes',
    '|',
    'Zoom',
    'Fit to Window',
    '|',
    'Macros',
  ],
};

const box = (locator) =>
  locator.evaluate((node) => {
    const rect = node.getBoundingClientRect();
    return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
  });

test(
  'Transitions, Animations, Slide Show, Record, Review and View follow Mac PowerPoint',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-show-tabs-'));
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
      const tab = (name) => page.getByRole('tab', { name, exact: true }).click();
      const button = (name) => panel.getByRole('button', { name, exact: true });
      const order = () =>
        panel
          .locator(':scope > div > section')
          .evaluateAll((sections) =>
            sections.flatMap((section, index) => [
              ...(index ? ['|'] : []),
              ...[...section.querySelectorAll('button, a')]
                .filter((node) => !node.closest('[role=menu], [role=radiogroup], .gallery'))
                .map((node) => node.textContent.replace('⌄', '').trim()),
            ]),
          );
      const tiles = (name) =>
        panel
          .getByRole('radiogroup', { name, exact: true })
          .getByRole('radio')
          .evaluateAll((nodes) => nodes.map((node) => node.getAttribute('aria-label')));
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

      for (const [name, widths] of Object.entries(WIDTHS)) {
        await tab(name);
        assert.equal((await box(panel.locator(':scope > div'))).height, 72, name);
        for (const [label, width] of Object.entries(widths))
          assert.equal((await box(button(label))).width, width, `${name} ▸ ${label}`);
        if (ORDER[name]) assert.deepEqual(await order(), ORDER[name], name);
      }

      // Transitions with no effect: Preview and Effect Options are disabled,
      // timing stays editable, and the gallery shows ten 92 × 56 pt tiles.
      await tab('Transitions');
      assert.equal(await button('Preview').isDisabled(), true);
      assert.equal(await button('Effect Options').isDisabled(), true);
      assert.equal(await panel.getByRole('spinbutton', { name: 'Duration:' }).isDisabled(), false);
      assert.deepEqual(await tiles('Transition Styles'), [
        'None',
        'Morph',
        'Fade',
        'Push',
        'Wipe',
        'Split',
        'Reveal',
        'Cut',
        'Random Bars',
        'Shape',
      ]);
      const none = await box(panel.getByRole('radio', { name: 'None', exact: true }));
      assert.deepEqual([none.width, none.height], [92, 56]);
      assert.equal(await panel.getByRole('radio', { name: 'Morph' }).isDisabled(), false);
      const duration = await box(panel.getByRole('spinbutton', { name: 'Duration:' }));
      assert.deepEqual([duration.width, duration.height], [78, 24]);
      const sound = await box(panel.getByRole('combobox', { name: 'Sound:' }));
      assert.deepEqual([sound.width, sound.height], [102, 26]);
      await panel.getByRole('button', { name: 'Next Transition Styles gallery' }).click();
      assert.equal((await tiles('Transition Styles'))[0], 'Uncover');

      // Push's Effect Options are PowerPoint's four directions.
      await panel.getByRole('button', { name: 'Previous Transition Styles gallery' }).click();
      await panel.getByRole('radio', { name: 'Push', exact: true }).click();
      await button('Effect Options').click();
      assert.deepEqual(await menuItems(), [
        { name: 'From Bottom', checked: true, disabled: false },
        { name: 'From Left', checked: false, disabled: false },
        { name: 'From Right', checked: false, disabled: false },
        { name: 'From Top', checked: false, disabled: false },
      ]);
      await page.getByRole('menuitemradio', { name: 'From Left' }).click();
      for (let i = 0; i < 100; i += 1) {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        if (getSlideTransition(getSlides(pres)[0])?.direction === 'r') break;
        assert.ok(i < 99, 'From Left was not saved as push dir="r"');
        await new Promise((resolve) => setTimeout(resolve, 50));
      }

      // Animations with a shape that has no effect: Effect Options, Trigger,
      // Animation Painter, Start and Duration are disabled; both galleries
      // show five 64 pt tiles.
      await page.locator('.hit').first().click();
      await tab('Animations');
      for (const name of ['Effect Options', 'Trigger', 'Animation Painter'])
        assert.equal(await button(name).isDisabled(), true, name);
      assert.equal(await panel.getByRole('combobox', { name: 'Start' }).isDisabled(), true);
      assert.deepEqual(await tiles('Entrance Effects'), [
        'Appear',
        'Blinds',
        'Checkerboard',
        'Dissolve In',
        'Fly In',
      ]);
      assert.deepEqual(await tiles('Emphasis Effects'), [
        'Fill Color',
        'Font Color',
        'Grow/Shrink',
        'Line Color',
        'Spin',
      ]);
      assert.equal((await box(panel.getByRole('radio', { name: 'Appear' }))).width, 64);
      assert.equal(
        (await box(panel.getByRole('combobox', { name: 'Start' }))).width,
        102,
        'Start pop-up',
      );
      await panel.getByRole('radio', { name: 'Fly In', exact: true }).click();
      await button('Effect Options').click();
      assert.deepEqual(await menuItems(), [
        { name: 'From Bottom', checked: true, disabled: false },
        { name: 'From Bottom-Left', checked: false, disabled: false },
        { name: 'From Left', checked: false, disabled: false },
        { name: 'From Top-Left', checked: false, disabled: false },
        { name: 'From Top', checked: false, disabled: false },
        { name: 'From Top-Right', checked: false, disabled: false },
        { name: 'From Right', checked: false, disabled: false },
        { name: 'From Bottom-Right', checked: false, disabled: false },
        { name: 'As One Object', checked: true, disabled: false },
        { name: 'All at Once', checked: false, disabled: true },
        { name: 'By Paragraph', checked: false, disabled: false },
      ]);
      await page.getByRole('menuitemradio', { name: 'From Left', exact: true }).click();
      await button('Effect Options').click();
      assert.equal(
        await page
          .getByRole('menuitemradio', { name: 'From Left', exact: true })
          .getAttribute('aria-checked'),
        'true',
      );
      await page.keyboard.press('Escape');

      // At 1200 pt the transition gallery shows six tiles and the Emphasis
      // gallery becomes a 55 pt button; the other tabs keep their layout.
      await page.setViewportSize({ width: 1200, height: 900 });
      await button('Emphasis Effects').waitFor();
      assert.equal(await panel.getByRole('radiogroup', { name: 'Emphasis Effects' }).count(), 0);
      assert.equal((await box(button('Emphasis Effects'))).width, 55);
      assert.equal((await tiles('Entrance Effects')).length, 6);
      await tab('Transitions');
      assert.equal((await tiles('Transition Styles')).length, 6);
      for (const name of ['Slide Show', 'Review', 'View']) {
        await tab(name);
        assert.deepEqual(await order(), ORDER[name], `${name} at 1200`);
      }

      // Japanese uses Mac PowerPoint's wording and still fits the row.
      await page.setViewportSize({ width: 1512, height: 900 });
      await page.locator('.lang select').selectOption('ja');
      await tab('スライド ショー');
      assert.deepEqual((await order()).slice(0, 4), [
        '最初から',
        '現在のスライドから',
        '発表者ビュー',
        'カスタム スライド ショー',
      ]);
      await tab('画面切り替え');
      assert.deepEqual((await tiles('画面切り替え効果')).slice(0, 4), [
        'なし',
        '変形',
        'フェード',
        'プッシュ',
      ]);
      await tab('アニメーション');
      await page.locator('.hit').first().click();
      const row = await box(panel.locator(':scope > div'));
      const last = await box(panel.locator(':scope > div > section').last());
      assert.ok(last.x + last.width <= row.x + row.width, 'the Japanese Animations row fits');
      await button('効果のオプション').click();
      assert.deepEqual((await menuItems()).map((item) => item.name).slice(0, 3), [
        '下から',
        '左下から',
        '左から',
      ]);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

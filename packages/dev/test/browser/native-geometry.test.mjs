import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { expandFormatSections } from '../helpers/format-pane.mjs';
import { startPreview } from '../helpers/server.mjs';

// Metrics measured from Mac PowerPoint 16 through the accessibility API in a
// 1512 × 900 pt window (2026-10-07). CSS px equal Mac points.
const DECK = `import {Presentation,Slide,Shape} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Shape preset="rect" x={4} y={4} width={3} height={1.5} fill="#2E75B6" /></Slide><Slide><Shape preset="rect" x={1} y={1} width={3} height={1} text="Two" /></Slide></Presentation>`;

const box = (locator) =>
  locator.evaluate((node) => {
    const rect = node.getBoundingClientRect();
    return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
  });

test(
  'window, ribbon, Format pane and status bar follow Mac PowerPoint geometry',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-native-geometry-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(file, DECK);
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1512, height: 900 } });
      await page.goto(preview.url + '/editor');
      await page.getByText('Saved to this project', { exact: true }).waitFor();

      // Home: a 72 pt command row; Add-ins and Designer are separate groups;
      // the second Font row ends with Highlight and Font Color after a rule.
      assert.equal((await box(page.locator('.home'))).height, 72);
      const groups = await page
        .locator('.home > section')
        .evaluateAll((nodes) => nodes.map((node) => node.getAttribute('aria-label')));
      assert.deepEqual(groups, [
        'Clipboard',
        'Slides',
        'Font',
        'Paragraph',
        'Insert',
        'Drawing',
        'Add-ins',
        'Designer',
      ]);
      const fontRow = await page.locator('.home .font-buttons').evaluate((row) =>
        [...row.querySelectorAll('button, input, .sep')]
          .filter((node) => !node.closest('.highlight-options'))
          .map((node) =>
            node.classList.contains('sep')
              ? '|'
              : (node.getAttribute('aria-label') ?? node.textContent.trim()),
          )
          .filter((name) => !name.endsWith('More Colors...')),
      );
      assert.deepEqual(fontRow, [
        'Bold',
        'Italic',
        'Underline',
        'Strikethrough',
        'Superscript',
        'Subscript',
        'Character Spacing',
        'Change Case',
        '|',
        'Highlight color',
        'Text color',
      ]);
      for (const name of ['Bold', 'Align Left', 'Increase Font Size']) {
        const button = page.locator('.home').getByRole('button', { name, exact: true });
        assert.deepEqual([(await box(button)).width, (await box(button)).height], [26, 26], name);
      }
      for (const name of ['Line spacing', 'Columns', 'Text Direction', 'Change Case']) {
        const button = page.locator('.home').getByRole('button', { name, exact: true });
        assert.deepEqual([(await box(button)).width, (await box(button)).height], [38, 26], name);
      }
      assert.equal(
        (await box(page.locator('.home').getByRole('button', { name: 'Picture', exact: true })))
          .width,
        50,
      );
      assert.equal(
        await page.getByRole('button', { name: 'Font dialog', exact: true }).count(),
        0,
        'PowerPoint has no Font dialog button on the Home ribbon',
      );

      // Thumbnail pane, one-line notes pane and status bar.
      assert.equal((await box(page.locator('.slide-workspace'))).x, 249);
      assert.equal((await box(page.locator('.notes-pane'))).height, 44);
      const status = await box(page.locator('.statusbar'));
      assert.equal(status.height, 28);
      assert.equal(status.y + status.height, 900);
      assert.equal(
        (await box(page.getByRole('slider', { name: 'Zoom percentage', exact: true }))).width,
        102,
      );
      assert.equal(
        (await box(page.getByRole('button', { name: 'Normal', exact: true }).last())).width,
        37,
      );

      // Zoom percentages are Mac PowerPoint's: 100% shows a 13.33 in slide
      // 960 pt wide, and Fit leaves 22 pt around it (native: 120% here).
      const percent = page.locator('.statusbar .zpct');
      const fit = Number((await percent.textContent()).replace('%', ''));
      assert.ok(fit >= 118 && fit <= 120, `fit zoom ${fit}%`);
      await percent.click();
      const dialog = page.getByRole('dialog', { name: 'Zoom', exact: true });
      await dialog.getByLabel('100%', { exact: true }).check();
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      assert.equal(Math.round((await box(page.locator('.stage'))).width), 960);
      await page.getByRole('button', { name: 'Fit slide to current window', exact: true }).click();

      // Format Shape pane: a 300 pt pane with a 27 pt title, 42 pt category
      // tabs, line-type radios, and 30 pt rows of 26 pt controls whose
      // pop-ups are 112 pt wide and end 17 pt from the pane edge.
      await page.locator('.hit').first().click({ button: 'right' });
      await page.getByRole('menuitem', { name: 'Format Shape...', exact: true }).click();
      const pane = await box(page.locator('.panel'));
      assert.equal(pane.width, 300);
      assert.equal((await box(page.locator('.panel-head'))).height, 27);
      // Shape Options / Text Options: a 278 × 26 pt switch, 14 pt below the
      // title, 11 pt in; every section collapsed, headers on a 25 pt pitch.
      const options = await box(page.getByRole('radiogroup', { name: 'Format Shape options' }));
      assert.deepEqual(
        [options.width, options.height, options.x - pane.x, options.y - (pane.y + 27)],
        [278, 26, 11, 14],
      );
      const headers = await page
        .locator('details.pane-section:visible > summary')
        .evaluateAll((nodes) =>
          nodes.map((node) => [node.parentElement.open, node.getBoundingClientRect().top]),
        );
      assert.deepEqual(
        headers.map(([open]) => open),
        [false, false],
      );
      assert.equal(headers[1][1] - headers[0][1], 25);
      await expandFormatSections(page, 'Fill', 'Line');
      assert.deepEqual(
        await page
          .getByRole('tablist', { name: 'Format Shape' })
          .getByRole('tab')
          .evaluateAll((tabs) => tabs.map((tab) => tab.getBoundingClientRect().height)),
        [42, 42, 42],
      );
      const lineTypes = page.getByRole('group', { name: 'Line type', exact: true });
      assert.deepEqual(
        await lineTypes
          .getByRole('radio')
          .evaluateAll((radios) => radios.map((radio) => radio.labels[0].textContent.trim())),
        ['No line', 'Solid line', 'Gradient line'],
      );
      const selects = await page.locator('.bespoke select').evaluateAll((nodes) =>
        nodes
          .filter((node) => node.getBoundingClientRect().width > 0)
          .map((node) => {
            const rect = node.getBoundingClientRect();
            return [rect.width, rect.height, rect.top];
          }),
      );
      assert.ok(selects.length >= 4);
      for (const [width, height] of selects) assert.deepEqual([width, height], [112, 26]);
      assert.deepEqual(
        selects.slice(1).map(([, , top], index) => top - selects[index][2]),
        selects.slice(1).map(() => 30),
      );
      const compound = await box(page.getByRole('combobox', { name: 'Compound type' }));
      assert.equal(pane.x + pane.width - (compound.x + compound.width), 17);
      const transparency = await box(
        page.getByRole('spinbutton', { name: 'Fill transparency', exact: true }),
      );
      assert.equal(transparency.height, 26);

      // Japanese uses Mac PowerPoint's own wording.
      await page.locator('.lang select').selectOption('ja');
      assert.deepEqual(
        await page
          .getByRole('group', { name: '線の種類', exact: true })
          .getByRole('radio')
          .evaluateAll((radios) => radios.map((radio) => radio.labels[0].textContent.trim())),
        ['線なし', '線 (単色)', '線 (グラデーション)'],
      );
      await page.locator('.thumb-row').first().click({ button: 'right' });
      const names = await page
        .getByRole('menu')
        .first()
        .locator(':scope > .ctx-item')
        .evaluateAll((nodes) => nodes.map((node) => node.firstElementChild.textContent.trim()));
      assert.deepEqual(names.slice(-3), ['ズーム...', 'スライド ショー', '新しいコメント']);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

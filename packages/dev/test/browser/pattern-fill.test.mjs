import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapePatternFill,
  getSlideXmlString,
  getSlideShapes,
  getSlides,
  loadPresentation,
} from '../../../../dist/index.js';
import { startPreview } from '../helpers/server.mjs';

test(
  'native pattern gallery preserves colors, multiple selections, undo and saved reloads',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-pattern-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={3} height={2} fill="#00FF00">First</Text><Text x={6} y={1} width={3} height={2} fill="#FF0000">Second</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1100 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const read = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getSlideShapes(getSlides(pres)[0]).map((shape) => getShapePatternFill(pres, shape));
      };
      const undo = async () => {
        await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
        await saved();
      };
      const redo = async () => {
        await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
        await saved();
      };
      await saved();
      await editor.locator('.hit').first().click();
      await editor.getByRole('radio', { name: 'Pattern fill', exact: true }).check();
      await saved();
      assert.deepEqual((await read())[0], {
        preset: 'pct5',
        foreground: '#4F81BD',
        background: '#FFFFFF',
      });
      const foreground = editor.getByRole('button', { name: 'Foreground', exact: true });
      const palette = editor.getByRole('menu', { name: 'Foreground', exact: true });
      await foreground.click();
      assert.equal(
        await palette
          .getByRole('menuitemradio', { name: 'Accent 1', exact: true })
          .getAttribute('aria-checked'),
        'true',
      );
      await palette.getByRole('menuitemradio', { name: 'Accent 2', exact: true }).click();
      await saved();
      await editor.getByRole('radio', { name: 'Solid fill', exact: true }).check();
      await saved();
      await editor.getByRole('radio', { name: 'Pattern fill', exact: true }).check();
      await saved();
      await page.reload();
      await saved();
      await editor.locator('.hit').first().click();
      await foreground.click();
      assert.equal(
        await palette
          .getByRole('menuitemradio', { name: 'Accent 2', exact: true })
          .getAttribute('aria-checked'),
        'true',
      );
      await palette.getByRole('menuitemradio', { name: 'Accent 1', exact: true }).click();
      await saved();
      const gallery = editor.getByRole('group', { name: 'Pattern', exact: true });
      assert.equal(await gallery.getByRole('button').count(), 48);
      assert.equal(
        await gallery
          .locator('img')
          .evaluateAll((images) =>
            images.every((image) => image.complete && image.naturalWidth > 0),
          ),
        true,
      );
      await gallery.getByRole('button', { name: 'Dotted: 10%', exact: true }).click();
      await saved();
      const themed = await loadPresentation(
        new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
      );
      assert.match(
        getSlideXmlString(getSlides(themed)[0]),
        /<a:pattFill prst="pct10"><a:fgClr><a:schemeClr val="accent1"\/><\/a:fgClr><a:bgClr><a:schemeClr val="bg1"\/><\/a:bgClr><\/a:pattFill>/,
      );
      await undo();
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      await editor.getByRole('radio', { name: 'Pattern fill', exact: true }).check();
      await saved();
      const mixed = await loadPresentation(
        new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
      );
      assert.equal(
        (
          getSlideXmlString(getSlides(mixed)[0]).match(
            /<a:fgClr><a:schemeClr val="accent1"\/><\/a:fgClr>/g,
          ) ?? []
        ).length,
        2,
      );
      await undo();
      await page.keyboard.press('Escape');
      await editor.locator('.hit').first().click();
      const color = async (label, value) => {
        await editor
          .getByLabel(`${label}: More Colors...`, { exact: true })
          .evaluate((input, value) => {
            input.value = value;
            input.dispatchEvent(new Event('change', { bubbles: true }));
          }, value);
        await saved();
      };
      await color('Foreground', '#123456');
      await color('Background', '#abcdef');
      await gallery.getByRole('button', { name: 'Wave', exact: true }).click();
      await saved();
      const wave = (await read())[0];
      assert.deepEqual(wave, { preset: 'wave', foreground: '#123456', background: '#ABCDEF' });
      await undo();
      assert.equal((await read())[0].preset, 'pct5');
      await redo();
      assert.deepEqual((await read())[0], wave);
      await editor.getByRole('radio', { name: 'Solid fill', exact: true }).check();
      await saved();
      await editor.getByRole('radio', { name: 'Pattern fill', exact: true }).check();
      await saved();
      assert.deepEqual((await read())[0], wave);
      await page.screenshot({ path: '/tmp/pptx-pattern-gallery.png' });
      await page.reload();
      await saved();
      await editor.locator('.hit').first().click();
      assert.equal(
        await gallery
          .getByRole('button', { name: 'Wave', exact: true })
          .getAttribute('aria-pressed'),
        'true',
      );
      await gallery.getByRole('button', { name: 'Wave', exact: true }).focus();
      await page.keyboard.press('End');
      await page.keyboard.press('Enter');
      await saved();
      assert.equal((await read())[0].preset, 'solidDmnd');
      await undo();
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      await editor.getByRole('radio', { name: 'Pattern fill', exact: true }).check();
      await saved();
      assert.equal((await read())[1].preset, 'pct5');
      await gallery
        .getByRole('button', { name: 'Diagonal stripes: Dark upward', exact: true })
        .click();
      await saved();
      assert.deepEqual(
        (await read()).map((fill) => fill.preset),
        ['dkUpDiag', 'dkUpDiag'],
      );
      assert.equal((await read())[0].foreground, '#123456');
      assert.equal((await read())[1].foreground, '#4F81BD');
      await undo();
      assert.deepEqual(
        (await read()).map((fill) => fill.preset),
        ['wave', 'pct5'],
      );
      await editor.locator('.lang select').selectOption('ja');
      assert.equal(
        await editor
          .getByRole('radio', { name: '塗りつぶし（パターン）', exact: true })
          .isChecked(),
        true,
      );
      assert.equal(
        await editor
          .getByRole('group', { name: 'パターン', exact: true })
          .getByRole('button')
          .count(),
        48,
      );
      assert.equal(await editor.getByLabel('前景', { exact: true }).isVisible(), true);
      assert.equal(await editor.getByLabel('背景', { exact: true }).isVisible(), true);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

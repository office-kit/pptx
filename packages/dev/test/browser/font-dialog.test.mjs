import { chromium } from 'playwright';
import { build } from 'esbuild';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';

test('mixed font controls apply only chosen properties and resolve indeterminate effects', async () => {
  const bundle = await build({
    entryPoints: [fileURLToPath(new URL('../../src/font-dialog.ts', import.meta.url))],
    bundle: true,
    write: false,
    format: 'iife',
    globalName: 'fontUI',
  });
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage();
    await page.setContent('<body></body>');
    await page.addScriptTag({ content: bundle.outputFiles[0].text });
    const open = () =>
      page.evaluate(() => {
        window.applied = null;
        window.restored = false;
        fontUI.openFontDialog(
          [
            {
              font: 'Arial',
              size: 20,
              bold: true,
              strike: true,
              color: '#FF0000',
              baseline: 0.3,
              spc: 100,
              kern: 1200,
            },
            {
              font: 'Calibri',
              size: 30,
              italic: true,
              strike: false,
              color: '#0000FF',
              baseline: -0.25,
              spc: -100,
              kern: 0,
            },
          ],
          async (settings) => {
            window.applied = settings;
            return true;
          },
          () => {
            window.restored = true;
          },
        );
      });
    const dialog = page.getByRole('dialog', { name: 'Font', exact: true });
    await open();
    assert.equal(await dialog.getByRole('textbox', { name: 'Font', exact: true }).inputValue(), '');
    for (const label of ['Size', 'Font style', 'Position'])
      assert.equal(await dialog.getByLabel(label, { exact: true }).inputValue(), '');
    assert.equal(
      await dialog
        .getByRole('checkbox', { name: 'Strikethrough' })
        .evaluate((el) => el.indeterminate),
      true,
    );
    await dialog.getByLabel('Size', { exact: true }).fill('24');
    await dialog.getByRole('button', { name: 'OK', exact: true }).click();
    assert.deepEqual(await page.evaluate(() => window.applied), { size: 24 });
    assert.equal(await page.evaluate(() => window.restored), true);
    await open();
    await dialog.getByRole('checkbox', { name: 'Strikethrough' }).uncheck();
    await dialog.getByLabel('Position', { exact: true }).selectOption('0');
    await dialog.getByRole('tab', { name: 'Character Spacing', exact: true }).click();
    assert.equal(await dialog.getByLabel('Spacing', { exact: true }).inputValue(), '');
    assert.equal(
      await dialog
        .getByRole('checkbox', { name: 'Kerning for fonts' })
        .evaluate((el) => el.indeterminate),
      true,
    );
    await dialog.getByLabel('Spacing', { exact: true }).selectOption('normal');
    await dialog.getByRole('checkbox', { name: 'Kerning for fonts' }).uncheck();
    await dialog.getByRole('button', { name: 'OK', exact: true }).click();
    assert.deepEqual(await page.evaluate(() => window.applied), {
      strike: false,
      baseline: 0,
      spc: 0,
      kern: 0,
    });
    await open();
    await dialog.getByLabel('Size', { exact: true }).fill('40');
    await page.keyboard.press('Escape');
    assert.equal(await page.evaluate(() => window.applied), null);
    assert.equal(await page.evaluate(() => window.restored), true);
  } finally {
    await browser.close();
  }
});

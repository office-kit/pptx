import { chromium } from 'playwright';
import { build } from 'esbuild';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';

test('mixed paragraph controls preserve untouched values and submit explicit zero', async () => {
  const bundle = await build({
    entryPoints: [fileURLToPath(new URL('../../src/paragraph-dialog.ts', import.meta.url))],
    bundle: true,
    write: false,
    format: 'iife',
    globalName: 'paragraphUI',
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
        paragraphUI.openParagraphDialog(
          [
            {
              align: 'l',
              marL: 91440,
              indent: -91440,
              spcBefPts: 6,
              spcAftPts: 8,
              lineSpacing: { kind: 'pts', value: 20 },
            },
            {
              align: 'left',
              marL: 182880,
              indent: 0,
              spcBefPts: 12,
              spcAftPts: 8,
              lineSpacing: { kind: 'pts', value: 30 },
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
    await open();
    const dialog = page.getByRole('dialog', { name: 'Paragraph', exact: true });
    assert.equal(await dialog.getByLabel('Alignment', { exact: true }).inputValue(), 'left');
    for (const label of ['Before text', 'Before', 'At', 'Special'])
      assert.equal(await dialog.getByLabel(label, { exact: true }).inputValue(), '');
    assert.equal(await dialog.getByLabel('After', { exact: true }).inputValue(), '8');
    await dialog.getByLabel('Before', { exact: true }).fill('0');
    await dialog.getByRole('button', { name: 'OK', exact: true }).click();
    assert.deepEqual(await page.evaluate(() => window.applied), { beforePts: 0 });
    assert.equal(await page.evaluate(() => window.restored), true);
    await open();
    await dialog.getByLabel('Special', { exact: true }).selectOption('hanging');
    await dialog.getByLabel('By', { exact: true }).fill('0.25');
    await dialog.getByLabel('At', { exact: true }).fill('24');
    await dialog.getByRole('button', { name: 'OK', exact: true }).click();
    assert.deepEqual(await page.evaluate(() => window.applied), {
      firstLineEmu: -228600,
      lineSpacing: { kind: 'pts', value: 24 },
    });
    await open();
    await dialog.getByLabel('Before', { exact: true }).fill('10');
    await page.keyboard.press('Escape');
    assert.equal(await page.evaluate(() => window.applied), null);
    assert.equal(await page.evaluate(() => window.restored), true);
  } finally {
    await browser.close();
  }
});

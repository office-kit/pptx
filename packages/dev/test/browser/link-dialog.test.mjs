import { chromium } from 'playwright';
import { build } from 'esbuild';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';

test('hyperlink dialog edits destinations and tips, cancels, removes and recovers failed saves', async () => {
  const bundle = await build({
    entryPoints: [fileURLToPath(new URL('../../src/link-dialog.ts', import.meta.url))],
    bundle: true,
    write: false,
    format: 'iife',
    globalName: 'linkUI',
  });
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage();
    await page.setContent('<body></body>');
    await page.addScriptTag({ content: bundle.outputFiles[0].text });
    const open = async (initial = null, fail = false) =>
      page.evaluate(
        ({ initial, fail }) => {
          window.applied = 'unchanged';
          window.restored = false;
          linkUI.openLinkDialog(
            initial,
            [{ key: '/ppt/slides/slide1.xml', title: 'Title <test>' }],
            async (link) => {
              if (fail) {
                fail = false;
                throw new Error('Please try again');
              }
              window.applied = link;
              return true;
            },
            () => {
              window.restored = true;
            },
          );
        },
        { initial, fail },
      );
    await open();
    await page.getByLabel('Address:', { exact: true }).fill('https://example.com/?a=1&b=2');
    await page.getByRole('button', { name: 'ScreenTip...' }).click();
    await page.getByLabel('ScreenTip text:').fill('資料 <詳細>');
    await page
      .getByRole('dialog', { name: 'Set Hyperlink ScreenTip' })
      .getByRole('button', { name: 'OK', exact: true })
      .click();
    await page.getByRole('button', { name: 'OK', exact: true }).click();
    assert.deepEqual(await page.evaluate(() => window.applied), {
      action: { kind: 'url', url: 'https://example.com/?a=1&b=2' },
      tooltip: '資料 <詳細>',
    });
    assert.equal(await page.evaluate(() => window.restored), true);
    await open();
    await page.getByRole('tab', { name: 'Web Page or File' }).focus();
    await page.keyboard.press('ArrowRight');
    assert.equal(
      await page.getByRole('tab', { name: 'This Document' }).getAttribute('aria-selected'),
      'true',
    );
    await page
      .getByLabel('Select a place in this document:')
      .selectOption('/ppt/slides/slide1.xml');
    await page.getByRole('button', { name: 'OK', exact: true }).click();
    assert.deepEqual(await page.evaluate(() => window.applied), {
      action: { kind: 'slide', slide: '/ppt/slides/slide1.xml' },
      tooltip: null,
    });
    const mail = {
      action: { kind: 'url', url: 'mailto:person@example.com?subject=Hello&cc=other@example.com' },
      tooltip: null,
    };
    await open(mail);
    await page.getByRole('button', { name: 'OK', exact: true }).click();
    assert.deepEqual(await page.evaluate(() => window.applied), mail);
    await open(mail);
    await page.getByLabel('Subject:').fill('資料 & 続き');
    await page.getByRole('button', { name: 'OK', exact: true }).click();
    assert.equal(
      (await page.evaluate(() => window.applied)).action.url,
      'mailto:person@example.com?subject=' +
        encodeURIComponent('資料 & 続き') +
        '&cc=other%40example.com',
    );
    await open(mail);
    await page.getByRole('button', { name: 'Remove Link' }).click();
    assert.equal(await page.evaluate(() => window.applied), null);
    await open();
    await page.keyboard.press('Escape');
    assert.equal(await page.evaluate(() => window.applied), 'unchanged');
    await open(mail, true);
    await page.getByRole('button', { name: 'OK', exact: true }).click();
    await page.getByRole('alert').filter({ hasText: 'Please try again' }).waitFor();
    await page.getByRole('button', { name: 'OK', exact: true }).click();
    assert.deepEqual(await page.evaluate(() => window.applied), mail);
  } finally {
    await browser.close();
  }
});

// Editing the layout behind the slide — Google Slides' theme builder, in the
// properties pane. The point of the feature is that the edit is shared, so the
// assertions are made on the saved deck's layout part, not on one slide.

import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlideLayout,
  getSlideLayoutBackground,
  getSlideLayoutName,
  getSlides,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const WORDS = {
  en: (key) => key,
  ja: (key) =>
    ({
      'Saved to this project': 'このプロジェクトに保存済み',
      Layout: 'レイアウト',
      'Layout name': 'レイアウト名',
      'Layout background color': 'レイアウトの背景色',
      'Reset layout background': 'レイアウトの背景をリセット',
      'Slides using this layout': 'このレイアウトを使うスライド',
    })[key],
};

const deck = async (dir) => {
  const file = join(dir, 'deck.tsx');
  await writeFile(
    file,
    `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation>{['A','B'].map(text => <Slide><Text x={1} y={1} width={7} height={1}>{text}</Text></Slide>)}</Presentation>`,
  );
  return file;
};

// The layout of the first slide, read back out of the saved .pptx.
const savedLayout = async (preview) => {
  const pres = await loadPresentation(
    new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
  );
  return getSlideLayout(getSlides(pres)[0]);
};

for (const language of ['en', 'ja']) {
  test(`renames and repaints the shared layout (${language})`, { timeout: 60000 }, async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-layout-edit-'));
    let preview, browser;
    const word = WORDS[language];
    try {
      preview = await startPreview(await deck(dir));
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () =>
        editor.getByText(word('Saved to this project'), { exact: true }).waitFor();
      // The pane is still English until the switch, so wait in English first.
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      if (language === 'ja') {
        await editor.locator('.lang select').selectOption('ja');
        await saved();
      }

      const pane = editor.getByRole('region', { name: word('Layout'), exact: true });
      // Both slides share the layout, which is what makes the edit deck-wide.
      await pane.getByText(`${word('Slides using this layout')}: 2`, { exact: true }).waitFor();

      await pane.getByLabel(word('Layout name'), { exact: true }).fill('Brand base');
      await pane.getByLabel(word('Layout name'), { exact: true }).press('Tab');
      await saved();
      assert.equal(getSlideLayoutName(await savedLayout(preview)), 'Brand base');

      const color = pane.getByRole('button', {
        name: word('Layout background color'),
        exact: true,
      });
      await color.click();
      const palette = editor.getByRole('menu', {
        name: word('Layout background color'),
        exact: true,
      });
      await palette
        .getByRole('menuitemradio', {
          name: language === 'ja' ? 'アクセント 4' : 'Accent 4',
          exact: true,
        })
        .click();
      await saved();
      assert.deepEqual(getSlideLayoutBackground(await savedLayout(preview)), {
        kind: 'solid',
        color: 'scheme:accent4',
      });

      await page.reload();
      await saved();
      await color.click();
      assert.equal(
        await palette
          .getByRole('menuitemradio', {
            name: language === 'ja' ? 'アクセント 4' : 'Accent 4',
            exact: true,
          })
          .getAttribute('aria-checked'),
        'true',
      );
      if (language === 'en')
        await page.screenshot({ path: '/tmp/pptx-background-palette.png', fullPage: true });
      await palette.press('Escape');

      await pane
        .getByRole('button', { name: word('Reset layout background'), exact: true })
        .click();
      await saved();
      assert.deepEqual(getSlideLayoutBackground(await savedLayout(preview)), { kind: 'inherit' });

      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  });
}

test('a layout edit is one undo step', { timeout: 60000 }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-layout-undo-'));
  let preview, browser;
  try {
    preview = await startPreview(await deck(dir));
    browser = await chromium.launch({ headless: true });
    const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
    await page.goto(preview.url);
    await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
    const editor = page.frameLocator('#editor-frame');
    const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
    await saved();
    const before = getSlideLayoutName(await savedLayout(preview));

    const pane = editor.getByRole('region', { name: 'Layout', exact: true });
    await pane.getByLabel('Layout name', { exact: true }).fill('Temporary');
    await pane.getByLabel('Layout name', { exact: true }).press('Tab');
    await saved();
    assert.equal(getSlideLayoutName(await savedLayout(preview)), 'Temporary');

    await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
    await saved();

    assert.equal(getSlideLayoutName(await savedLayout(preview)), before);
  } finally {
    await browser?.close();
    await preview?.close();
    await rm(dir, { recursive: true, force: true });
  }
});

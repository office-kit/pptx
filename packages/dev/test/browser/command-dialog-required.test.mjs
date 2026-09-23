// The generated argument dialog is the fallback UI for every capability, so a
// required argument left empty is an ordinary thing for an author to do. It
// must say which field is missing instead of dispatching an undefined argument
// into the core, which surfaces as a raw TypeError from library internals.

import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  inches,
  savePresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const WORDS = {
  en: {
    saved: 'Saved to this project',
    missing: 'Fill in first:',
    home: 'Home',
    textFormat: 'Text format',
    apply: 'Apply',
  },
  ja: {
    saved: 'このプロジェクトに保存済み',
    missing: '先に入力してください:',
    home: 'ホーム',
    textFormat: '文字の書式',
    apply: '適用',
  },
};

const deckWith = async (dir) => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(1),
    text: '一枚目 First',
  });
  const file = join(dir, 'deck.tsx');
  const source = join(dir, 'source.pptx');
  await writeFile(source, await savePresentation(pres));
  await writeFile(
    file,
    `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
  );
  return file;
};

for (const language of ['en', 'ja']) {
  test(
    `an empty required argument names the field instead of throwing (${language})`,
    { timeout: 60000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-required-arg-'));
      const words = WORDS[language];
      let preview, browser;
      try {
        preview = await startPreview(await deckWith(dir));
        browser = await chromium.launch({ headless: true });
        const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
        const errors = [];
        page.on('pageerror', (e) => errors.push(e.message));
        await page.goto(preview.url);
        await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
        const editor = page.frameLocator('#editor-frame');
        await editor.getByText(WORDS.en.saved, { exact: true }).waitFor();
        if (language === 'ja') {
          await editor.locator('.lang select').selectOption('ja');
          await editor.getByText(words.saved, { exact: true }).waitFor();
        }

        await editor.locator('.hit').first().click();
        await editor.getByRole('button', { name: words.home, exact: true }).click();
        await editor
          .locator('.ribbon')
          .getByRole('button', { name: words.textFormat, exact: true })
          .click();
        const dialog = editor.getByRole('dialog');
        await dialog.waitFor();
        await dialog.getByRole('button', { name: words.apply, exact: true }).click();

        // The dialog stays open, names the field, and nothing reached the core.
        const alert = dialog.getByRole('alert');
        await alert.waitFor();
        assert.match(await alert.innerText(), new RegExp(`^${words.missing}\\s+\\S`));
        await dialog.waitFor({ state: 'visible' });
        assert.equal(await editor.locator('.toast').count(), 0);
        assert.deepEqual(errors, []);

        // Filling the argument lets the same Apply through.
        await dialog.getByLabel('Bold', { exact: true }).check();
        await dialog.getByRole('button', { name: words.apply, exact: true }).click();
        await dialog.waitFor({ state: 'detached' });
        await editor.getByText(words.saved, { exact: true }).waitFor();
        assert.deepEqual(errors, []);
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );
}

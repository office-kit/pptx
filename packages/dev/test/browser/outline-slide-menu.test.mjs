import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile, copyFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  loadPresentation,
  getSlides,
  getSlideLayout,
  getSlideLayoutPartName,
  getSlideShapes,
  getShapeText,
} from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';

for (const locale of ['en', 'ja']) {
  test(
    `outline text menus add, duplicate and delete whole slides (${locale})`,
    { timeout: 60000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-outline-menu-'));
      let preview, browser;
      try {
        await copyFile(
          new URL('../../../../test/fixtures/minimal/one-text-slide.pptx', import.meta.url),
          join(dir, 'template.pptx'),
        );
        const file = join(dir, 'deck.tsx');
        await writeFile(
          file,
          `import {readFileSync} from 'node:fs'; import {Presentation,Slide,Fill,Text} from '@office-kit/pptx-dsl'; const source = new Uint8Array(readFileSync(new URL('./template.pptx', import.meta.url))); export default <Presentation source={source} mode="compose"><Slide layout={{name:"Title and Content"}}><Fill target={{placeholder:{type:'title'}}}>Menu title</Fill><Fill target={{placeholder:{idx:1}}}>Menu body</Fill><Text x={1} y={5} width={5} height={1}>Extra object</Text></Slide></Presentation>;`,
        );
        preview = await startPreview(file);
        browser = await chromium.launch({ headless: true });
        const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
        const errors = [];
        page.on('pageerror', (error) => errors.push(error.message));
        await page.addInitScript(
          (language) => localStorage.setItem('ok-editor-locale', language),
          locale,
        );
        await page.goto(preview.url);
        await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
        const editor = page.frameLocator('#editor-frame');
        const view = locale === 'en' ? 'View' : '表示';
        await editor.getByRole('tab', { name: view, exact: true }).click();
        await editor
          .getByRole('tabpanel', { name: view, exact: true })
          .getByRole('button', {
            name: locale === 'en' ? 'Outline View' : 'アウトライン表示',
            exact: true,
          })
          .click();
        const body = editor.getByRole('textbox', {
          name: locale === 'en' ? 'Outline text 1' : 'アウトラインのテキスト 1',
          exact: true,
        });
        const read = async () =>
          loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          );
        const change = async (action) => {
          const revision = (await waitForState(preview.url, () => true)).revision;
          await action();
          await waitForState(preview.url, (state) => state.revision !== revision);
        };
        const texts = (pres) =>
          getSlides(pres).map((slide) => getSlideShapes(slide).map(getShapeText).filter(Boolean));
        const original = await read();
        const originalTexts = texts(original);
        for (const [en, ja, count] of [
          ['New slide', '新しいスライド', 2],
          ['Duplicate slide', 'スライドを複製', 2],
          ['Delete slide', 'スライドを削除', 0],
        ]) {
          await body.click({ button: 'right' });
          await change(() =>
            editor
              .getByRole('menuitem', { name: new RegExp(`^${locale === 'en' ? en : ja}(?:$| )`) })
              .click(),
          );
          const changed = await read();
          assert.equal(getSlides(changed).length, count);
          if (en === 'New slide') {
            assert.equal(
              getSlideLayoutPartName(getSlideLayout(getSlides(changed)[1])),
              getSlideLayoutPartName(getSlideLayout(getSlides(original)[0])),
            );
            assert.deepEqual(texts(changed)[1], []);
          } else if (en === 'Duplicate slide')
            assert.deepEqual(texts(changed)[1], originalTexts[0]);
          await change(() => editor.locator('body').press('Control+z'));
          assert.deepEqual(texts(await read()), originalTexts);
        }
        assert.deepEqual(errors, []);
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );
}

import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile, copyFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  loadPresentation,
  getSlides,
  getSlideShapes,
  getShapeText,
  findSlidePlaceholder,
} from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';

for (const locale of ['en', 'ja'])
  for (const titleOnly of [false, true])
    test(
      `outline demotion confirms object deletion and restores it with Undo (${locale}, title only: ${titleOnly})`,
      { timeout: 90000 },
      async () => {
        const dir = await mkdtemp(join(tmpdir(), 'office-outline-demote-'));
        let preview, browser;
        try {
          await copyFile(
            new URL('../../../../test/fixtures/minimal/one-text-slide.pptx', import.meta.url),
            join(dir, 'template.pptx'),
          );
          const file = join(dir, 'deck.tsx');
          await writeFile(
            file,
            `import {readFileSync} from 'node:fs'; import {Presentation,Slide,Fill,Text} from '@office-kit/pptx-dsl'; const source = new Uint8Array(readFileSync(new URL('./template.pptx', import.meta.url))); export default <Presentation source={source} mode="compose"><Slide layout={{name:${JSON.stringify(titleOnly ? 'Title Only' : 'Title and Content')}}}><Fill target={{placeholder:{type:'title'}}}>First</Fill>${titleOnly ? '' : '<Fill target={{placeholder:{idx:1}}}>Existing</Fill>'}</Slide><Slide layout={{name:"Title and Content"}}><Fill target={{placeholder:{type:'title'}}}>Second</Fill><Fill target={{placeholder:{idx:1}}}>Child</Fill><Text x={1} y={5} width={5} height={1}>Additional object</Text></Slide></Presentation>;`,
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
          const title = editor.getByRole('textbox', {
            name: locale === 'en' ? 'Outline title 2' : 'アウトラインのタイトル 2',
            exact: true,
          });
          const read = async () =>
            loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            );
          const revision = (await waitForState(preview.url, () => true)).revision;
          await title.press('Tab');
          const dialog = editor.getByRole('dialog', {
            name: locale === 'en' ? 'Demote' : 'レベル下げ',
            exact: true,
          });
          await dialog.waitFor();
          await dialog
            .getByRole('button', { name: locale === 'en' ? 'No' : 'いいえ', exact: true })
            .click();
          assert.equal(getSlides(await read()).length, 2);
          assert.equal((await waitForState(preview.url, () => true)).revision, revision);
          await title.press('Tab');
          await dialog
            .getByRole('button', { name: locale === 'en' ? 'Yes' : 'はい', exact: true })
            .click();
          await waitForState(preview.url, (state) => state.revision !== revision);
          const merged = await read();
          assert.equal(getSlides(merged).length, 1);
          assert.equal(
            getShapeText(findSlidePlaceholder(getSlides(merged)[0], 'body')),
            titleOnly ? 'Second\nChild' : 'Existing\nSecond\nChild',
          );
          const beforeUndo = (await waitForState(preview.url, () => true)).revision;
          await editor
            .getByRole('textbox', {
              name: locale === 'en' ? 'Outline text 1' : 'アウトラインのテキスト 1',
              exact: true,
            })
            .press('Control+z');
          await waitForState(preview.url, (state) => state.revision !== beforeUndo);
          const restored = await read();
          assert.equal(getSlides(restored).length, 2);
          assert.ok(
            getSlideShapes(getSlides(restored)[1]).some(
              (shape) => getShapeText(shape) === 'Additional object',
            ),
          );
          assert.deepEqual(errors, []);
        } finally {
          await browser?.close();
          await preview?.close();
          await rm(dir, { recursive: true, force: true });
        }
      },
    );

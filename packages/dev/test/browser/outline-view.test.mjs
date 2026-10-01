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
  getShapeParagraphElements,
} from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';

for (const locale of ['en', 'ja'])
  test(
    `outline edits placeholder text, inserts slides, undoes and saves (${locale})`,
    { timeout: 60000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-outline-'));
      let preview, browser;
      try {
        await copyFile(
          new URL('../../../../test/fixtures/minimal/one-text-slide.pptx', import.meta.url),
          join(dir, 'template.pptx'),
        );
        const file = join(dir, 'deck.tsx');
        await writeFile(
          file,
          `import {readFileSync} from 'node:fs'; import {Presentation,Slide,Fill,Text} from '@office-kit/pptx-dsl'; const source = new Uint8Array(readFileSync(new URL('./template.pptx', import.meta.url))); export default <Presentation source={source} mode="compose"><Slide layout={{name:"Title and Content"}}><Fill target={{placeholder:{type:'title'}}} format={{bold:true}}>Outline title</Fill><Fill target={{placeholder:{idx:1}}} format={{italic:true}}>{"First point\\nSecond point"}</Fill><Text x={1} y={5} width={5} height={1}>Ordinary text box</Text></Slide></Presentation>;`,
        );
        preview = await startPreview(file);
        browser = await chromium.launch({ headless: true });
        const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
        const pageErrors = [];
        page.on('pageerror', (error) => pageErrors.push(error.message));
        await page.addInitScript(
          (language) => localStorage.setItem('ok-editor-locale', language),
          locale,
        );
        await page.goto(preview.url);
        await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
        const editor = page.frameLocator('#editor-frame');
        const labels =
          locale === 'en'
            ? { view: 'View', outline: 'Outline View', title: 'Outline title', normal: 'Normal' }
            : {
                view: '表示',
                outline: 'アウトライン表示',
                title: 'アウトラインのタイトル',
                normal: '標準',
              };
        await editor.getByRole('tab', { name: labels.view, exact: true }).click();
        await editor
          .getByRole('tabpanel', { name: labels.view, exact: true })
          .getByRole('button', { name: labels.outline, exact: true })
          .click();
        const outline = editor.getByRole('navigation', { name: labels.outline, exact: true });
        const title = outline.getByRole('textbox', { name: `${labels.title} 1`, exact: true });
        assert.equal(await title.inputValue(), 'Outline title');
        assert.equal(await outline.getByRole('textbox').count(), 2);
        const read = async () =>
          loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          );
        const change = async (action) => {
          const before = (await waitForState(preview.url, () => true)).revision;
          await action();
          await waitForState(preview.url, (state) => state.revision !== before);
        };
        const slideIcon = outline.locator('[data-outline-slide="0"] > button');
        await change(() => slideIcon.dblclick());
        assert.equal(await slideIcon.getAttribute('aria-expanded'), 'false');
        assert.equal(await outline.getByRole('textbox').count(), 1);
        await change(() => slideIcon.press('Control+z'));
        assert.equal(await outline.getByRole('textbox').count(), 2);
        await change(() => title.fill('Outline title edited'));
        let pres = await read();
        assert.equal(getShapeText(getSlideShapes(getSlides(pres)[0])[0]), 'Outline title edited');
        assert.equal(
          getShapeParagraphElements(getSlideShapes(getSlides(pres)[0])[0], 0)[0].format.bold,
          true,
        );
        await title.press('End');
        await change(() => title.press('Enter'));
        pres = await read();
        assert.equal(getShapeText(getSlideShapes(getSlides(pres)[0])[1]), '');
        const movedBody = getSlideShapes(getSlides(pres)[1]).find(
          (shape) => getShapeText(shape) === 'First point\nSecond point',
        );
        assert.ok(movedBody, 'title Enter moves the following body to the new slide');
        assert.equal(getShapeParagraphElements(movedBody, 0)[0].format.italic, true);
        const second = outline.getByRole('textbox', { name: `${labels.title} 2`, exact: true });
        await second.waitFor();
        assert.equal(
          await second.evaluate((node) => node === node.ownerDocument.activeElement),
          true,
        );
        await change(() => second.fill('Second title'));
        pres = await read();
        assert.equal(getSlides(pres).length, 2);
        assert.equal(getShapeText(getSlideShapes(getSlides(pres)[1])[0]), 'Second title');
        assert.equal(
          getSlideShapes(getSlides(pres)[1]).some(
            (shape) => getShapeText(shape) === 'Ordinary text box',
          ),
          false,
        );
        await change(() => second.press('Control+z'));
        assert.equal(
          await outline
            .getByRole('textbox', { name: `${labels.title} 2`, exact: true })
            .inputValue(),
          '',
        );
        await change(() =>
          outline
            .getByRole('textbox', { name: `${labels.title} 2`, exact: true })
            .press('Control+z'),
        );
        assert.equal(getSlides(await read()).length, 1);
        await title.focus();
        await title.evaluate((node) => node.setSelectionRange(8, 8));
        await change(() => title.press('Enter'));
        pres = await read();
        assert.equal(getShapeText(getSlideShapes(getSlides(pres)[0])[0]), 'Outline ');
        assert.equal(getShapeText(getSlideShapes(getSlides(pres)[1])[0]), 'title edited');
        assert.equal(
          getShapeParagraphElements(getSlideShapes(getSlides(pres)[1])[0], 0)[0].format.bold,
          true,
        );
        await change(() => second.press('Control+z'));
        assert.equal(getSlides(await read()).length, 1);
        await change(async () => {
          await title.fill('Saved when leaving outline');
          await editor
            .getByRole('tabpanel', { name: labels.view, exact: true })
            .getByRole('button', { name: labels.normal, exact: true })
            .click();
        });
        assert.equal(await editor.locator('.slide-workspace').count(), 1);
        await page.reload();
        await editor.locator('.slide-workspace').waitFor();
        assert.equal(
          getShapeText(getSlideShapes(getSlides(await read())[0])[0]),
          'Saved when leaving outline',
        );
        await editor.getByRole('tab', { name: labels.view, exact: true }).click();
        const openOutline = () =>
          editor
            .getByRole('tabpanel', { name: labels.view, exact: true })
            .getByRole('button', { name: labels.outline, exact: true })
            .click();
        await openOutline();
        await change(() => slideIcon.dblclick());
        await page.reload();
        await editor.locator('.slide-workspace').waitFor();
        await editor.getByRole('tab', { name: labels.view, exact: true }).click();
        await openOutline();
        assert.equal(await slideIcon.getAttribute('aria-expanded'), 'false');
        assert.equal(await outline.getByRole('textbox').count(), 1);
        await change(() => slideIcon.dblclick());
        assert.equal(await slideIcon.getAttribute('aria-expanded'), 'true');
        assert.equal(await outline.getByRole('textbox').count(), 2);
        assert.deepEqual(pageErrors, []);
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );

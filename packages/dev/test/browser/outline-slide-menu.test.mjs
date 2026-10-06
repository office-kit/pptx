import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile, readFile } from 'node:fs/promises';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
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
  getShapeRunFormatEffective,
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
        const parts = unzipSync(
          await readFile(
            new URL('../../../../test/fixtures/minimal/one-text-slide.pptx', import.meta.url),
          ),
        );
        // A non-default master exposes accidental inheritance from a disposable slide.
        const master = 'ppt/slideMasters/slideMaster1.xml';
        const masterXml = strFromU8(parts[master]);
        assert.ok(masterXml.includes('sz="3200"'));
        parts['ppt/slideMasters/slideMaster2.xml'] = strToU8(
          masterXml.replaceAll('sz="3200"', 'sz="6000"'),
        );
        parts['ppt/slideMasters/_rels/slideMaster2.xml.rels'] =
          parts['ppt/slideMasters/_rels/slideMaster1.xml.rels'];
        const layoutRels = 'ppt/slideLayouts/_rels/slideLayout2.xml.rels';
        parts[layoutRels] = strToU8(
          strFromU8(parts[layoutRels]).replace('slideMaster1.xml', 'slideMaster2.xml'),
        );
        parts['[Content_Types].xml'] = strToU8(
          strFromU8(parts['[Content_Types].xml']).replace(
            '</Types>',
            '<Override PartName="/ppt/slideMasters/slideMaster2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/></Types>',
          ),
        );
        parts['ppt/presentation.xml'] = strToU8(
          strFromU8(parts['ppt/presentation.xml']).replace(
            '</p:sldMasterIdLst>',
            '<p:sldMasterId id="2147483649" r:id="rIdOutlineMaster"/></p:sldMasterIdLst>',
          ),
        );
        parts['ppt/_rels/presentation.xml.rels'] = strToU8(
          strFromU8(parts['ppt/_rels/presentation.xml.rels']).replace(
            '</Relationships>',
            '<Relationship Id="rIdOutlineMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster2.xml"/></Relationships>',
          ),
        );
        await writeFile(join(dir, 'template.pptx'), zipSync(parts));
        const file = join(dir, 'deck.tsx');
        await writeFile(
          file,
          `import {readFileSync} from 'node:fs'; import {Presentation,Slide,Fill,Text} from '@office-kit/pptx-dsl'; const source = new Uint8Array(readFileSync(new URL('./template.pptx', import.meta.url))); export default <Presentation source={source} mode="compose"><Slide layout={{name:"Title and Content"}}><Fill target={{placeholder:{type:'title'}}}>Menu title</Fill><Fill target={{placeholder:{idx:1}}} format={{bold:true,italic:true}}>Menu body</Fill><Text x={1} y={5} width={5} height={1}>Extra object</Text></Slide></Presentation>;`,
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
        for (const enabled of [true, false]) {
          const initialRevision = (await waitForState(preview.url, () => true)).revision;
          await body.click({ button: 'right' });
          const toggle = editor.getByRole('menuitemcheckbox', {
            name: locale === 'en' ? 'Show Formatting' : '書式の表示',
            exact: true,
          });
          assert.equal(await toggle.getAttribute('aria-checked'), String(!enabled));
          await toggle.click();
          if (enabled) {
            await body.locator('span').first().waitFor();
            assert.equal(
              await body
                .locator('span')
                .first()
                .evaluate((node) => getComputedStyle(node).fontWeight),
              '700',
            );
            assert.equal(
              await body
                .locator('span')
                .first()
                .evaluate((node) => getComputedStyle(node).fontStyle),
              'italic',
            );
          } else assert.equal(await body.locator('span').count(), 0);
          assert.equal(await body.textContent(), 'Menu body');
          assert.equal((await waitForState(preview.url, () => true)).revision, initialRevision);
          const displayedSize = enabled
            ? await body
                .locator('span')
                .first()
                .evaluate((node) => getComputedStyle(node).fontSize)
            : null;
          await change(async () => {
            await body.evaluate((node) => {
              node.focus();
              const range = document.createRange();
              range.selectNodeContents(node);
              range.collapse(false);
              const selection = window.getSelection();
              selection.removeAllRanges();
              selection.addRange(range);
            });
            if (enabled) {
              const pendingSize = await body.evaluate(async (node) => {
                node.dispatchEvent(
                  new InputEvent('beforeinput', {
                    bubbles: true,
                    inputType: 'insertText',
                    data: '!',
                  }),
                );
                node.querySelector('span').append('!');
                node.dispatchEvent(
                  new InputEvent('input', { bubbles: true, inputType: 'insertText', data: '!' }),
                );
                await new Promise((resolve) => setTimeout(resolve, 0));
                return getComputedStyle(node.querySelector('span')).fontSize;
              });
              assert.equal(
                pendingSize,
                displayedSize,
                'pending outline edits retain inherited font size',
              );
            } else await body.press('!');
            await editor.getByRole('tab', { name: view, exact: true }).click();
          });
          const edited = await read();
          const editedBody = getSlideShapes(getSlides(edited)[0]).find(
            (shape) => getShapeText(shape) === 'Menu body!',
          );
          assert.ok(editedBody, JSON.stringify(texts(edited)));
          const format = getShapeRunFormatEffective(edited, editedBody, 0, 0);
          assert.equal(format.bold, true);
          assert.equal(format.italic, true);
          const originalBody = getSlideShapes(getSlides(original)[0]).find(
            (shape) => getShapeText(shape) === 'Menu body',
          );
          assert.equal(format.size, getShapeRunFormatEffective(original, originalBody, 0, 0).size);
          await change(() => editor.locator('body').press('Control+z'));
          assert.deepEqual(texts(await read()), originalTexts);
        }

        for (const [en, ja, count] of [
          ['New Slide', '新しいスライド', 2],
          ['Duplicate Slide', 'スライドを複製', 2],
          ['Delete Slide', 'スライドを削除', 0],
        ]) {
          await body.click({ button: 'right' });
          await change(() =>
            editor
              .getByRole('menuitem', { name: new RegExp(`^${locale === 'en' ? en : ja}(?:$| )`) })
              .click(),
          );
          const changed = await read();
          assert.equal(getSlides(changed).length, count);
          if (en === 'New Slide') {
            assert.equal(
              getSlideLayoutPartName(getSlideLayout(getSlides(changed)[1])),
              getSlideLayoutPartName(getSlideLayout(getSlides(original)[0])),
            );
            assert.deepEqual(texts(changed)[1], []);
          } else if (en === 'Duplicate Slide')
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

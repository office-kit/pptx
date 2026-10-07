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
  getParagraphLevel,
  getShapeRunHyperlink,
} from '@office-kit/pptx';
import { installRichTextSelection } from '../helpers/rich-text.mjs';
import { startPreview, waitForState } from '../helpers/server.mjs';

for (const locale of ['en', 'ja'])
  test(
    `outline edits placeholder text, inserts slides, undoes and saves (${locale})`,
    { timeout: 90000 },
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
          `import {readFileSync} from 'node:fs'; import {getSlideShapes,setShapeRunHyperlink,setParagraphBullet} from '@office-kit/pptx'; import {Presentation,Slide,Fill,Text,Raw} from '@office-kit/pptx-dsl'; const source = new Uint8Array(readFileSync(new URL('./template.pptx', import.meta.url))); export default <Presentation source={source} mode="compose"><Slide layout={{name:"Title and Content"}}><Fill target={{placeholder:{type:'title'}}} format={{bold:true}}>Outline title</Fill><Fill target={{placeholder:{idx:1}}} format={{italic:true}}>{"First point\\nSecond point"}</Fill><Text x={1} y={5} width={5} height={1}>Ordinary text box</Text><Raw scope="slide" apply={({slide}) => {setShapeRunHyperlink(getSlideShapes(slide)[0],0,0,"https://example.com/outline");setParagraphBullet(getSlideShapes(slide)[1],0,"bullet");setParagraphBullet(getSlideShapes(slide)[1],1,"bullet");}} /></Slide></Presentation>;`,
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
        await installRichTextSelection(page);
        await page.goto(preview.url);
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
        assert.equal(await title.textContent(), 'Outline title');
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
        const body = outline.getByRole('textbox', {
          name: `${locale === 'en' ? 'Outline text' : 'アウトラインのテキスト'} 1`,
          exact: true,
        });
        const paragraphLayout = () =>
          body.evaluate((node) =>
            [...node.querySelectorAll('[data-outline-paragraph]')].map((paragraph) => ({
              top: paragraph.getBoundingClientRect().top,
              padding: parseFloat(getComputedStyle(paragraph).paddingLeft),
              marker: paragraph.dataset.outlineMarker,
            })),
          );
        const initialLayout = await paragraphLayout();
        assert.equal(initialLayout.length, 2);
        assert.equal(initialLayout[1].top - initialLayout[0].top, 24);
        assert.deepEqual(
          initialLayout.map((paragraph) => paragraph.marker),
          ['•', '•'],
        );
        // Mac PowerPoint outline body Shift+Enter creates a:p, unlike canvas a:br.
        for (const key of ['Enter', 'Shift+Enter']) {
          await body.focus();
          await body.evaluate((node) => window.selectEditorText(node, 5, 5));
          await change(() => body.press(key));
          const savedSlides = getSlides(await read());
          assert.equal(savedSlides.length, 1);
          const savedBody = getSlideShapes(savedSlides[0])[1];
          assert.equal(getShapeText(savedBody), 'First\n point\nSecond point');
          for (const [index, text] of ['First', ' point', 'Second point'].entries()) {
            const elements = getShapeParagraphElements(savedBody, index);
            assert.equal(elements.map((element) => element.text ?? '').join(''), text);
            assert.ok(elements.every((element) => element.kind === 'r'));
            assert.equal(elements[0].format.italic, true);
          }
          await change(() => body.press('Control+z'));
          assert.equal(await body.textContent(), 'First point\nSecond point');
        }
        const moveParagraph = async (up) => {
          await body.click({ button: 'right', position: { x: 30, y: up ? 30 : 10 } });
          await change(() =>
            editor
              .getByRole('menuitem', {
                name:
                  locale === 'en' ? (up ? 'Move Up' : 'Move Down') : up ? '上へ移動' : '下へ移動',
                exact: true,
              })
              .click(),
          );
        };
        await body.focus();
        await body.evaluate((node) => window.selectEditorText(node, 15, 15));
        await moveParagraph(true);
        assert.equal(await body.textContent(), 'Second point\nFirst point');
        assert.equal(
          getShapeText(getSlideShapes(getSlides(await read())[0])[1]),
          'Second point\nFirst point',
        );
        assert.equal(
          getShapeParagraphElements(getSlideShapes(getSlides(await read())[0])[1], 0)[0].format
            .italic,
          true,
        );
        assert.deepEqual(
          await body.evaluate((node) =>
            (() => {
              const selected = document.getSelection().getRangeAt(0);
              const before = document.createRange();
              before.selectNodeContents(node);
              before.setEnd(selected.startContainer, selected.startOffset);
              const start = before.toString().length;
              return [start, start + selected.toString().length];
            })(),
          ),
          [0, 12],
        );
        await moveParagraph(false);
        assert.equal(await body.textContent(), 'First point\nSecond point');
        await change(() => body.press('Control+z'));
        assert.equal(await body.textContent(), 'Second point\nFirst point');
        await change(() => body.press('Control+z'));
        assert.equal(await body.textContent(), 'First point\nSecond point');
        await body.focus();
        await body.evaluate((node) => window.selectEditorText(node, 0, 0));
        await body.click({ button: 'right', position: { x: 30, y: 10 } });
        assert.equal(
          await editor
            .getByRole('menuitem', { name: locale === 'en' ? 'Move Up' : '上へ移動', exact: true })
            .isDisabled(),
          true,
        );
        await page.keyboard.press('Escape');
        await body.focus();
        await body.evaluate((node) => window.selectEditorText(node, 15, 15));
        await change(() => body.press('Tab'));
        let bodyShape = getSlideShapes(getSlides(await read())[0])[1];
        assert.equal(getParagraphLevel(bodyShape, 0), 0);
        assert.equal(getParagraphLevel(bodyShape, 1), 1);
        const indentedLayout = await paragraphLayout();
        assert.equal(indentedLayout[1].padding - indentedLayout[0].padding, 10);
        assert.equal(indentedLayout[1].top - indentedLayout[0].top, 24);
        assert.equal(getShapeText(bodyShape), 'First point\nSecond point');
        assert.equal(getShapeParagraphElements(bodyShape, 1)[0].format.italic, true);
        assert.equal(
          await body.evaluate((node) => node === node.ownerDocument.activeElement),
          true,
        );
        await change(() => body.press('Shift+Tab'));
        bodyShape = getSlideShapes(getSlides(await read())[0])[1];
        assert.equal(getParagraphLevel(bodyShape, 1), 0);
        await change(() => body.press('Control+z'));
        await change(() => body.press('Control+z'));
        bodyShape = getSlideShapes(getSlides(await read())[0])[1];
        assert.equal(getParagraphLevel(bodyShape, 1), 0);
        await body.focus();
        await body.evaluate((node) => window.selectEditorText(node, 0, 12));
        await change(() => body.press('Tab'));
        bodyShape = getSlideShapes(getSlides(await read())[0])[1];
        assert.equal(getParagraphLevel(bodyShape, 0), 1);
        assert.equal(
          getParagraphLevel(bodyShape, 1),
          0,
          'exclusive selection end excludes the next paragraph',
        );
        await change(() => body.press('Control+z'));
        await body.focus();
        await body.evaluate((node) => window.selectEditorText(node, 12, 24));
        await change(() => body.press('Shift+Tab'));
        const promotedTitle = outline.getByRole('textbox', {
          name: `${labels.title} 2`,
          exact: true,
        });
        assert.equal(await promotedTitle.textContent(), 'Second point');
        assert.equal(getShapeText(getSlideShapes(getSlides(await read())[0])[1]), 'First point');
        await change(() => promotedTitle.press('Control+z'));
        assert.equal(await body.textContent(), 'First point\nSecond point');
        await change(() => body.press('Control+Shift+z'));
        assert.equal(await promotedTitle.textContent(), 'Second point');
        await change(() => promotedTitle.press('Tab'));
        assert.equal(getSlides(await read()).length, 1);
        assert.equal(await body.textContent(), 'First point\nSecond point');
        assert.equal(
          await body.evaluate((node) => node === node.ownerDocument.activeElement),
          true,
        );
        await change(() => body.press('Control+z'));
        assert.equal(await promotedTitle.textContent(), 'Second point');
        await change(() => promotedTitle.press('Control+z'));
        await body.focus();
        await body.evaluate((node) => window.selectEditorText(node, 0, 24));
        await body.click({ button: 'right' });
        await change(() =>
          editor
            .getByRole('menuitem', {
              name: locale === 'en' ? 'Promote' : 'レベル上げ',
              exact: true,
            })
            .click(),
        );
        assert.equal(getSlides(await read()).length, 3);
        assert.equal(await promotedTitle.textContent(), 'First point');
        assert.equal(
          await outline
            .getByRole('textbox', {
              name: `${labels.title} 3`,
              exact: true,
            })
            .textContent(),
          'Second point',
        );
        await change(() => promotedTitle.press('Control+z'));
        assert.equal(await body.textContent(), 'First point\nSecond point');
        await change(async () => {
          await title.fill('Outline title edited');
          await title.press('Meta+1');
          await outline.waitFor({ state: 'detached' });
        });
        await editor.getByRole('menuitem', { name: labels.view, exact: true }).focus();
        await page.keyboard.press('Meta+4');
        await outline.waitFor();
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
        // Arrow navigation at an outline textbox boundary continues into the
        // adjacent title/body, matching PowerPoint's keyboard outline flow.
        const secondBody = outline.getByRole('textbox', {
          name: `${locale === 'en' ? 'Outline text' : 'アウトラインのテキスト'} 2`,
          exact: true,
        });
        const selectionOffset = (node) =>
          node.evaluate((root) => {
            const selection = root.ownerDocument.getSelection();
            if (!selection?.rangeCount) return null;
            const range = selection.getRangeAt(0);
            const before = root.ownerDocument.createRange();
            before.selectNodeContents(root);
            before.setEnd(range.startContainer, range.startOffset);
            return [before.toString().length, selection.toString().length];
          });
        await second.focus();
        await second.press('End');
        await second.press('ArrowDown');
        assert.equal(
          await secondBody.evaluate((node) => node === node.ownerDocument.activeElement),
          true,
        );
        assert.deepEqual(await selectionOffset(secondBody), [0, 0]);
        await secondBody.press('ArrowUp');
        assert.equal(
          await second.evaluate((node) => node === node.ownerDocument.activeElement),
          true,
        );
        assert.deepEqual(await selectionOffset(second), ['Second title'.length, 0]);
        await change(async () => {
          await second.fill('Pending title');
          await second.press('End');
          await second.press('ArrowDown');
        });
        assert.equal(
          await secondBody.evaluate((node) => node === node.ownerDocument.activeElement),
          true,
          'pending title edits commit before ArrowDown moves to the body',
        );
        await change(() => secondBody.press('Control+z'));
        assert.equal(await second.textContent(), 'Second title');
        await second.focus();
        await second.press('End');
        await second.press('Control+ArrowDown');
        assert.equal(
          await second.evaluate((node) => node === node.ownerDocument.activeElement),
          true,
          'Ctrl+ArrowDown remains native text navigation at the outline boundary',
        );
        await second.press('Meta+ArrowDown');
        assert.equal(
          await second.evaluate((node) => node === node.ownerDocument.activeElement),
          true,
          'Meta+ArrowDown remains native text navigation at the outline boundary',
        );
        for (const up of [false, true]) {
          await second.click({ button: 'right' });
          await change(() =>
            editor
              .getByRole('menuitem', {
                name:
                  locale === 'en' ? (up ? 'Move Up' : 'Move Down') : up ? '上へ移動' : '下へ移動',
                exact: true,
              })
              .click(),
          );
          const saved = getSlides(await read());
          assert.equal(getShapeText(getSlideShapes(saved[0])[1]), up ? '' : 'First point');
          const savedBody = getSlideShapes(saved[1]).find(
            (shape) => getShapeText(shape) === (up ? 'First point\nSecond point' : 'Second point'),
          );
          assert.ok(savedBody);
          assert.equal(getShapeParagraphElements(savedBody, 0)[0].format.italic, true);
          assert.deepEqual(
            await second.evaluate((node) =>
              (() => {
                const selected = document.getSelection().getRangeAt(0);
                const before = document.createRange();
                before.selectNodeContents(node);
                before.setEnd(selected.startContainer, selected.startOffset);
                const start = before.toString().length;
                return [start, start + selected.toString().length];
              })(),
            ),
            [0, 12],
          );
        }
        await change(() => second.press('Control+z'));
        assert.equal(await body.textContent(), 'First point');
        await change(() => second.press('Control+z'));
        assert.equal(await body.textContent(), '');
        const secondIcon = outline.locator('[data-outline-slide="1"] > button');
        const collapseLabel = locale === 'en' ? 'Collapse' : '折りたたむ';
        const collapseAllLabel = locale === 'en' ? 'Collapse All' : 'すべて折りたたむ';
        const expandLabel = locale === 'en' ? 'Expand' : '展開';
        const expandAllLabel = locale === 'en' ? 'Expand All' : 'すべて展開';
        await second.click({ button: 'right' });
        const collapseMenu = editor.getByRole('menuitem', { name: collapseLabel, exact: true });
        await collapseMenu.focus();
        await collapseMenu.press('ArrowRight');
        const collapseSubmenu = editor.getByRole('menu', { name: collapseLabel, exact: true });
        await change(() =>
          collapseSubmenu.getByRole('menuitem', { name: collapseAllLabel, exact: true }).click(),
        );
        assert.equal(await outline.getByRole('textbox').count(), 2);
        await change(() => secondIcon.press('Control+z'));
        assert.equal(await outline.getByRole('textbox').count(), 4);
        await second.click({ button: 'right' });
        await editor.getByRole('menuitem', { name: collapseLabel, exact: true }).click();
        await change(() =>
          editor
            .getByRole('menu', { name: collapseLabel, exact: true })
            .getByRole('menuitem', { name: collapseLabel, exact: true })
            .click(),
        );
        assert.equal(await outline.getByRole('textbox').count(), 3);
        assert.equal(await slideIcon.getAttribute('aria-expanded'), 'true');
        await outline
          .getByRole('textbox', { name: `${labels.title} 2`, exact: true })
          .click({ button: 'right' });
        await editor.getByRole('menuitem', { name: expandLabel, exact: true }).click();
        await change(() =>
          editor
            .getByRole('menu', { name: expandLabel, exact: true })
            .getByRole('menuitem', { name: expandAllLabel, exact: true })
            .click(),
        );
        assert.equal(await outline.getByRole('textbox').count(), 4);
        await change(() => secondIcon.press('Control+z'));
        await change(() => secondIcon.press('Control+z'));
        await change(() => second.press('Control+z'));
        assert.equal(
          await outline
            .getByRole('textbox', { name: `${labels.title} 2`, exact: true })
            .textContent(),
          '',
        );
        await change(() =>
          outline
            .getByRole('textbox', { name: `${labels.title} 2`, exact: true })
            .press('Control+z'),
        );
        assert.equal(getSlides(await read()).length, 1);
        await title.focus();
        await title.evaluate((node) => window.selectEditorText(node, 8, 8));
        await change(() => title.press('Shift+Enter'));
        pres = await read();
        assert.equal(getShapeText(getSlideShapes(getSlides(pres)[0])[0]), 'Outline ');
        assert.equal(getShapeText(getSlideShapes(getSlides(pres)[1])[0]), 'title edited');
        assert.equal(
          getShapeParagraphElements(getSlideShapes(getSlides(pres)[1])[0], 0)[0].format.bold,
          true,
        );
        assert.equal(
          getShapeRunHyperlink(getSlideShapes(getSlides(pres)[1])[0], 0, 0),
          'https://example.com/outline',
          'splitting an outline title retains its hyperlink after saving',
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

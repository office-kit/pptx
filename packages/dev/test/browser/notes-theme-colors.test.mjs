import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  _internalPackageOf,
  createPresentation,
  addBlankSlide,
  setSlideNotes,
  setSlideNotesFormat,
  savePresentation,
  loadPresentation,
  getSlideNotesTextFormats,
  getSlideNotes,
  getSlides,
} from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';

test(
  'notes use their own theme and retain scheme colors after typing and Undo',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-notes-theme-'));
    let preview, browser;
    try {
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      setSlideNotes(slide, 'Theme note');
      setSlideNotesFormat(slide, { color: 'tx1' }, { range: { start: 0, end: 5 } });
      setSlideNotesFormat(slide, { color: 'accent1' }, { range: { start: 5, end: 10 } });
      // The paragraph-end format is the source of truth for a newly-created
      // empty paragraph.  Keep it as a scheme token so typing can inherit it
      // without flattening the notes theme color.
      setSlideNotesFormat(slide, { color: 'tx1' }, { paragraphEnd: 0 });
      const pkg = _internalPackageOf(pres);
      const encoder = new TextEncoder();
      const decoder = new TextDecoder();
      const notesName = '/ppt/notesSlides/notesSlide1.xml';
      const masterName = '/ppt/notesMasters/notesMaster1.xml';
      const themeName = '/ppt/theme/notesTheme.xml';
      const notesXml = decoder.decode(pkg.getPart(notesName).data);
      const map =
        'bg1="lt1" tx1="accent2" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"';
      pkg.addPart(
        masterName,
        'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml',
        encoder.encode(
          notesXml
            .replace('<p:notes ', '<p:notesMaster ')
            .replace('</p:notes>', '</p:notesMaster>')
            .replace(/<p:clrMapOvr>.*?<\/p:clrMapOvr>/s, `<p:clrMap ${map}/>`),
        ),
      );
      const theme = pkg.parts.find(
        (part) => part.contentType === 'application/vnd.openxmlformats-officedocument.theme+xml',
      );
      const themeXml = decoder.decode(theme.data);
      assert.match(themeXml, /<a:accent2>/);
      pkg.addPart(
        themeName,
        theme.contentType,
        encoder.encode(
          themeXml.replace(
            /<a:accent2>.*?<\/a:accent2>/s,
            '<a:accent2><a:srgbClr val="12AB34"/></a:accent2>',
          ),
        ),
      );
      const relType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
      pkg.setRels(masterName, {
        items: [
          {
            id: 'rId1',
            type: relType + 'theme',
            target: '../theme/notesTheme.xml',
            targetMode: 'Internal',
          },
        ],
      });
      const rels = pkg.getRels(notesName);
      rels.items.push({
        id: 'rId2',
        type: relType + 'notesMaster',
        target: '../notesMasters/notesMaster1.xml',
        targetMode: 'Internal',
      });
      pkg.setRels(notesName, rels);
      await writeFile(join(dir, 'source.pptx'), await savePresentation(pres));
      await writeFile(
        join(dir, 'deck.tsx'),
        `import {readFileSync} from 'node:fs'; import {Presentation} from '@office-kit/pptx-dsl'; export default <Presentation source={new Uint8Array(readFileSync(new URL('./source.pptx', import.meta.url)))}/>;`,
      );
      preview = await startPreview(join(dir, 'deck.tsx'));
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.getByRole('button', { name: 'Notes', exact: true }).click();
      const input = editor.getByRole('textbox', { name: 'Notes content', exact: true });
      const color = () =>
        input.evaluate((el) => {
          const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
          const node = walker.nextNode();
          return getComputedStyle(node.parentElement).color;
        });
      const lastColor = () =>
        input.evaluate((el) => {
          const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
          let node;
          let last;
          while ((node = walker.nextNode())) last = node;
          return getComputedStyle(last.parentElement).color;
        });
      assert.equal(await color(), 'rgb(18, 171, 52)');
      await input.click();
      await input.press('Home');
      const toolbarColor = editor.getByRole('button', { name: 'Text color', exact: true });
      await toolbarColor.waitFor();
      assert.equal(
        await toolbarColor
          .locator('.swatch')
          .evaluate((el) => getComputedStyle(el).backgroundColor),
        'rgb(18, 171, 52)',
        'the toolbar should show the resolved theme color while retaining the stored token',
      );
      await input.screenshot({ path: join(tmpdir(), 'notes-theme-colors.png') });
      const before = (await waitForState(preview.url, () => true)).revision;
      await input.click();
      await input.press('End');
      await input.press('!');
      assert.equal(
        await color(),
        'rgb(18, 171, 52)',
        'a pending insertion must inherit the resolved theme color before commit',
      );
      assert.equal(await lastColor(), 'rgb(79, 129, 189)');
      await input.press('Tab');
      const edited = await waitForState(preview.url, (state) => state.revision !== before);
      const read = async () =>
        loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
      const saved = await read();
      assert.equal(getSlideNotes(getSlides(saved)[0]), 'Theme note!');
      assert.ok(
        getSlideNotesTextFormats(getSlides(saved)[0]).some((item) => item.format.color === 'tx1') &&
          getSlideNotesTextFormats(getSlides(saved)[0]).some(
            (item) => item.format.color === 'accent1',
          ),
        'typing must retain the literal colors of both adjacent runs',
      );
      assert.equal(await color(), 'rgb(18, 171, 52)');
      await input.press('Meta+z');
      await waitForState(preview.url, (state) => state.revision !== edited.revision);
      assert.equal(await input.textContent(), 'Theme note');
      assert.equal(await color(), 'rgb(18, 171, 52)');
      assert.equal(await lastColor(), 'rgb(79, 129, 189)');
      assert.equal(
        await toolbarColor
          .locator('.swatch')
          .evaluate((el) => getComputedStyle(el).backgroundColor),
        'rgb(79, 129, 189)',
        'Undo must refresh the resolved notes toolbar color without moving the caret',
      );
      await input.press('Meta+a');
      await toolbarColor.click();
      await editor.getByRole('menuitemradio', { name: 'Red', exact: true }).click();
      const red = await waitForState(preview.url, (state) => state.revision !== edited.revision);
      await input.press('Meta+z');
      await waitForState(preview.url, (state) => state.revision !== red.revision);
      assert.equal(
        await toolbarColor
          .locator('.swatch')
          .evaluate((el) => getComputedStyle(el).backgroundColor),
        'rgb(79, 129, 189)',
        'Undo after a different color must restore the resolved theme color',
      );

      // Replace text across the two differently formatted runs.  The pending
      // replacement must inherit the literal format of its carrier run, and
      // the unaffected right-hand run must retain its own scheme token after
      // the edit is committed and undone.
      await input.click();
      await input.press('Home');
      for (let i = 0; i < 3; i += 1) await input.press('ArrowRight');
      for (let i = 0; i < 4; i += 1) await input.press('Shift+ArrowRight');
      await input.press('X');
      assert.match(await input.textContent(), /TheXote/);
      assert.equal(
        await color(),
        'rgb(18, 171, 52)',
        'a cross-run replacement must display the inherited notes theme color before commit',
      );
      assert.equal(await lastColor(), 'rgb(79, 129, 189)');
      const beforeReplacement = (await waitForState(preview.url, () => true)).revision;
      await input.press('Tab');
      const replacement = await waitForState(
        preview.url,
        (state) => state.revision !== beforeReplacement,
      );
      const replaced = await read();
      const replacedSlide = getSlides(replaced)[0];
      assert.equal(getSlideNotes(replacedSlide), 'TheXote');
      const replacedFormats = getSlideNotesTextFormats(replacedSlide);
      assert.ok(replacedFormats.some((item) => item.format.color === 'tx1'));
      assert.ok(replacedFormats.some((item) => item.format.color === 'accent1'));

      await input.press('Meta+z');
      await waitForState(preview.url, (state) => state.revision !== replacement.revision);
      assert.equal(await input.textContent(), 'Theme note');
      const afterReplacementUndo = await read();
      const afterReplacementUndoSlide = getSlides(afterReplacementUndo)[0];
      assert.equal(getSlideNotes(afterReplacementUndoSlide), 'Theme note');
      const undoneFormats = getSlideNotesTextFormats(afterReplacementUndoSlide);
      assert.ok(undoneFormats.some((item) => item.format.color === 'tx1'));
      assert.ok(undoneFormats.some((item) => item.format.color === 'accent1'));

      // Enter creates a genuinely empty paragraph. Its first typed character
      // must inherit the paragraph-end tx1 token, including while the edit is
      // still pending, and save/undo must preserve the token.
      await input.click();
      await input.press('End');
      await input.press('Enter');
      await input.press('ArrowDown');
      await input.press('Z');
      assert.match(await input.textContent(), /Theme note[\n ]+Z/);
      assert.equal(
        await lastColor(),
        'rgb(18, 171, 52)',
        'typing into an empty paragraph must resolve its paragraph-end theme color',
      );
      const beforeEmptyParagraph = (await waitForState(preview.url, () => true)).revision;
      await input.press('Tab');
      const emptyParagraphEdit = await waitForState(
        preview.url,
        (state) => state.revision !== beforeEmptyParagraph,
      );
      const emptySaved = await read();
      const emptySavedSlide = getSlides(emptySaved)[0];
      const emptySavedText = getSlideNotes(emptySavedSlide);
      assert.match(emptySavedText, /Theme note[\n ]+Z/);
      const zIndex = emptySavedText.indexOf('Z');
      const emptyFormats = getSlideNotesTextFormats(emptySavedSlide);
      const savedZFormat = emptyFormats.find(
        (item) => item.start <= zIndex && item.end >= zIndex + 1,
      );
      assert.ok(
        savedZFormat,
        'empty-paragraph typing must save a format range covering the new character',
      );
      assert.equal(
        savedZFormat?.format.color,
        'tx1',
        'the new empty-paragraph character must retain the literal paragraph-end token',
      );
      const reopenedResolvedZFormat = getSlideNotesTextFormats(emptySavedSlide, {
        resolveColors: true,
      }).find((item) => item.start <= zIndex && item.end >= zIndex + 1);
      assert.equal(
        reopenedResolvedZFormat?.format.color,
        '#12AB34',
        'reopening the saved deck must resolve the new character through the notes theme',
      );
      await input.press('Meta+z');
      await waitForState(preview.url, (state) => state.revision !== emptyParagraphEdit.revision);
      assert.equal(await input.textContent(), 'Theme note');
      const emptyUndo = await read();
      assert.equal(getSlideNotes(getSlides(emptyUndo)[0]), 'Theme note');
      assert.ok(
        getSlideNotesTextFormats(getSlides(emptyUndo)[0]).some(
          (item) => item.format.color === 'tx1',
        ),
        'Undo must restore the original literal notes token',
      );

      // Repeated pending newlines must continue carrying the same paragraph
      // end format even though the committed slide has not gained those
      // paragraphs yet.
      const repeatedBeforeSave = (await waitForState(preview.url, () => true)).revision;
      await input.click();
      await input.press('End');
      await input.press('Enter');
      assert.equal(await input.textContent(), 'Theme note\n');
      await input.press('Enter');
      await input.press('Z');
      assert.equal(await input.textContent(), 'Theme note\n\nZ');
      assert.equal(
        await lastColor(),
        'rgb(18, 171, 52)',
        'repeated pending empty paragraphs must retain the inherited notes theme color',
      );
      await input.press('Tab');
      await waitForState(preview.url, (state) => state.revision !== repeatedBeforeSave);
      const repeatedSavedSlide = getSlides(await read())[0];
      const repeatedText = getSlideNotes(repeatedSavedSlide);
      const repeatedZ = repeatedText.indexOf('Z');
      const repeatedZFormat = getSlideNotesTextFormats(repeatedSavedSlide).find(
        (item) => item.start <= repeatedZ && item.end >= repeatedZ + 1,
      );
      assert.equal(
        repeatedZFormat?.format.color,
        'tx1',
        'repeated pending paragraph typing must save the inherited literal token',
      );
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

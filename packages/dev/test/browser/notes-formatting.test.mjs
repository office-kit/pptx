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
  getSlideNotes,
  getSlides,
} from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';
import { openNotes } from '../helpers/notes.mjs';

test(
  'editing notes preserves untouched formatted runs across disjoint edits and undo',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-notes-format-'));
    let preview, browser;
    try {
      const pres = createPresentation();
      setSlideNotes(addBlankSlide(pres), 'First');
      const name = '/ppt/notesSlides/notesSlide1.xml';
      const part = _internalPackageOf(pres).getPart(name);
      const suffix = '<a:r><a:rPr i="1"/><a:t> suffix</a:t></a:r>';
      part.data = new TextEncoder().encode(
        new TextDecoder()
          .decode(part.data)
          .replace('<a:rPr lang="en-US"/>', '<a:rPr b="1"/>')
          .replace('</a:r>', '</a:r>' + suffix),
      );
      await writeFile(join(dir, 'source.pptx'), await savePresentation(pres));
      await writeFile(
        join(dir, 'deck.tsx'),
        `import {readFileSync} from 'node:fs'; import {Presentation} from '@office-kit/pptx-dsl'; export default <Presentation source={new Uint8Array(readFileSync(new URL('./source.pptx', import.meta.url)))}/>;`,
      );
      preview = await startPreview(join(dir, 'deck.tsx'));
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await openNotes(editor);
      const input = editor.getByRole('textbox', { name: 'Notes content', exact: true });
      assert.match(await input.innerHTML(), /font-weight:\s*bold/i);
      const revision = (await waitForState(preview.url, () => true)).revision;
      await input.evaluate((el) => {
        const point = (offset) => {
          const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
          let remaining = offset;
          let node;
          while ((node = walker.nextNode())) {
            if (remaining <= node.textContent.length) return [node, remaining];
            remaining -= node.textContent.length;
          }
          return [el, el.childNodes.length];
        };
        for (const [start, end, text] of [
          [0, 5, 'New'],
          [10, 10, '!'],
        ]) {
          const selection = document.getSelection();
          const domRange = document.createRange();
          const [startNode, startOffset] = point(start);
          const [endNode, endOffset] = point(end);
          domRange.setStart(startNode, startOffset);
          domRange.setEnd(endNode, endOffset);
          selection.removeAllRanges();
          selection.addRange(domRange);
          document.dispatchEvent(new Event('selectionchange'));
          el.dispatchEvent(
            new InputEvent('beforeinput', { bubbles: true, inputType: 'insertText', data: text }),
          );
          document.execCommand('insertText', false, text);
        }
      });
      await input.press('Tab');
      const after = await waitForState(preview.url, (s) => s.revision !== revision);
      const read = async () =>
        loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
      const edited = await read();
      assert.equal(getSlideNotes(getSlides(edited)[0]), 'New suffix!');
      const xml = new TextDecoder().decode(_internalPackageOf(edited).getPart(name).data);
      // The '!' typed at the end of the italic run extends it, as in the reference desktop app.
      assert.ok(
        xml.includes('<a:r><a:rPr i="1"/><a:t> suffix!</a:t></a:r>'),
        'the italic run survives both edits',
      );
      assert.ok(xml.includes('<a:rPr b="1"/>'), 'replacement inherits original bold formatting');
      await input.press('Meta+z');
      await waitForState(preview.url, (s) => s.revision !== after.revision);
      const undone = await read();
      assert.equal(getSlideNotes(getSlides(undone)[0]), 'First suffix');
      assert.ok(
        new TextDecoder().decode(_internalPackageOf(undone).getPart(name).data).includes(suffix),
      );
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test('formatted notes paste is one undoable transaction', { timeout: 60000 }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-notes-paste-'));
  let preview, browser;
  try {
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide} from '@office-kit/pptx-dsl';export default <Presentation><Slide notes="Original"/></Presentation>`,
    );
    preview = await startPreview(file);
    browser = await chromium.launch({ headless: true });
    const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
    await page.goto(preview.url);
    const editor = page.frameLocator('#editor-frame');
    await editor.getByText('Saved to this project', { exact: true }).waitFor();
    await openNotes(editor);
    const input = editor.getByRole('textbox', { name: 'Notes content', exact: true });
    await input.selectText();
    const before = (await waitForState(preview.url, () => true)).revision;
    await input.evaluate((element) => {
      const data = new DataTransfer();
      data.setData('text/plain', 'Pasted');
      data.setData('text/html', '<span style="font-weight: bold">Pasted</span>');
      element.dispatchEvent(new ClipboardEvent('paste', { bubbles: true, clipboardData: data }));
    });
    const pasted = await waitForState(preview.url, (state) => state.revision !== before);
    const read = async () =>
      loadPresentation(
        new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
      );
    const formatted = await read();
    assert.equal(getSlideNotes(getSlides(formatted)[0]), 'Pasted');
    const name = '/ppt/notesSlides/notesSlide1.xml';
    assert.match(
      new TextDecoder().decode(_internalPackageOf(formatted).getPart(name).data),
      /<a:rPr[^>]*\bb="1"/,
      'HTML paste preserves the bold run',
    );

    await input.press('Meta+z');
    await waitForState(preview.url, (state) => state.revision !== pasted.revision);
    const undone = await read();
    assert.equal(
      getSlideNotes(getSlides(undone)[0]),
      'Original',
      'one Undo restores both pasted text and its formatting',
    );
  } finally {
    await browser?.close();
    await preview?.close();
    await rm(dir, { recursive: true, force: true });
  }
});

test('setSlideNotesFormat applies run formatting to a selected notes range', () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  setSlideNotes(slide, 'Bold note');
  setSlideNotesFormat(slide, { bold: true }, { range: { start: 0, end: 4 } });
  const name = '/ppt/notesSlides/notesSlide1.xml';
  const xml = new TextDecoder().decode(_internalPackageOf(pres).getPart(name).data);
  assert.match(xml, /<a:rPr[^>]*\bb="1"/);
  assert.match(xml, /<a:t> note<\/a:t>/);
});

test(
  'notes script buttons use proportional baselines and toggle off',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-notes-script-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide} from '@office-kit/pptx-dsl';export default <Presentation><Slide notes="Script"/></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await openNotes(editor);
      const input = editor.getByRole('textbox', { name: 'Notes content', exact: true });
      await input.selectText();
      for (const [button, baseline] of [
        ['Superscript', 30000],
        ['Superscript', 0],
        ['Subscript', -25000],
        ['Subscript', 0],
      ]) {
        const before = (await waitForState(preview.url, () => true)).revision;
        await editor.getByRole('button', { name: button, exact: true }).click();
        await waitForState(preview.url, (state) => state.revision !== before);
        const saved = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        const xml = new TextDecoder().decode(
          _internalPackageOf(saved).getPart('/ppt/notesSlides/notesSlide1.xml').data,
        );
        assert.match(
          xml,
          new RegExp(`baseline="${baseline}"`),
          `${button} writes the OOXML percentage for its actual offset`,
        );
        assert.equal(getSlideNotes(getSlides(saved)[0]), 'Script');
      }
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

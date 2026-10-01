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
  savePresentation,
  loadPresentation,
  getSlideNotes,
  getSlides,
} from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';

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
      await editor.getByRole('button', { name: 'Notes', exact: true }).click();
      const input = editor.getByRole('textbox', { name: 'Notes content', exact: true });
      const revision = (await waitForState(preview.url, () => true)).revision;
      await input.evaluate((el) => {
        for (const [start, end, text] of [
          [0, 5, 'New'],
          [10, 10, '!'],
        ]) {
          el.setSelectionRange(start, end);
          el.dispatchEvent(
            new InputEvent('beforeinput', { bubbles: true, inputType: 'insertText', data: text }),
          );
          el.setRangeText(text, start, end, 'end');
          el.dispatchEvent(
            new InputEvent('input', { bubbles: true, inputType: 'insertText', data: text }),
          );
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
      assert.ok(xml.includes(suffix), 'unchanged italic run survives both edits');
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

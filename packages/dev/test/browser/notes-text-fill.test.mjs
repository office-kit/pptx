import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  addBlankSlide,
  createPresentation,
  getSlideNotes,
  getSlideNotesTextFormats,
  getSlides,
  loadPresentation,
  savePresentation,
  setSlideNotes,
  setSlideNotesFormat,
} from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';
import { openNotes } from '../helpers/notes.mjs';

test(
  'notes font color replaces inherited pattern fill for subsequent typing',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-notes-text-fill-'));
    let preview, browser;
    try {
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      setSlideNotes(slide, 'Pattern');
      setSlideNotesFormat(
        slide,
        {
          textFill: {
            kind: 'pattern',
            preset: 'dkUpDiag',
            foreground: 'accent1',
            background: '#FFFFFF',
          },
        },
        { range: { start: 0, end: 7 } },
      );
      await writeFile(join(dir, 'source.pptx'), await savePresentation(pres));
      await writeFile(
        join(dir, 'deck.tsx'),
        `import {readFileSync} from 'node:fs'; import {Presentation} from '@office-kit/pptx-dsl'; export default <Presentation source={new Uint8Array(readFileSync(new URL('./source.pptx', import.meta.url)))}/>;`,
      );
      preview = await startPreview(join(dir, 'deck.tsx'));
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await openNotes(editor);
      const input = editor.getByRole('textbox', { name: 'Notes content', exact: true });
      await input.press('End');
      // An unrelated command first captures the inherited pattern in the pending
      // typing format. The next color command must replace that fill choice.
      await editor.getByRole('button', { name: 'Bold', exact: true }).click();
      await editor.getByRole('button', { name: 'Text color', exact: true }).click();
      await editor
        .getByRole('menu', { name: 'Text color', exact: true })
        .getByRole('menuitemradio', { name: 'Red', exact: true })
        .click();
      const before = (await waitForState(preview.url, () => true)).revision;
      await input.pressSequentially('X');
      await input.press('Tab');
      await waitForState(preview.url, (state) => state.revision !== before);
      const read = async () =>
        getSlides(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
        )[0];
      const check = (saved) => {
        assert.equal(getSlideNotes(saved), 'PatternX');
        const formats = getSlideNotesTextFormats(saved);
        assert.equal(formats.find((span) => span.start === 0)?.format.textFill?.kind, 'pattern');
        const inserted = formats.find((span) => span.start <= 7 && span.end > 7)?.format;
        assert.equal(inserted?.color, '#FF0000');
        assert.equal(inserted?.textFill, undefined);
        assert.equal(inserted?.bold, true);
      };
      check(await read());
      await page.reload();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      check(await read());
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

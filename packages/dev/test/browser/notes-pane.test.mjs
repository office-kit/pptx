import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  _internalPackageOf,
  getShapeParagraphElements,
  getSlideNotes,
  getSlideNotesParagraphEndFormat,
  getSlideNotesTextFormats,
  getSlideShapes,
  getSlides,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';
import { openNotes } from '../helpers/notes.mjs';

test('caret formatting is inherited by subsequent notes typing', { timeout: 60000 }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-notes-caret-format-'));
  let preview, browser;
  try {
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide} from '@office-kit/pptx-dsl';export default <Presentation><Slide notes="First"/></Presentation>`,
    );
    preview = await startPreview(file);
    browser = await chromium.launch({ headless: true });
    const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
    await page.goto(preview.url);
    await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
    const editor = page.frameLocator('#editor-frame');
    await editor.getByText('Saved to this project', { exact: true }).waitFor();
    await openNotes(editor);
    const input = editor.getByRole('textbox', { name: 'Notes content', exact: true });
    const beforeTyping = (await waitForState(preview.url, () => true)).revision;
    await input.press('End');
    await editor.getByRole('button', { name: 'Bold', exact: true }).click();
    await input.pressSequentially('Ab');
    await input.press('Tab');
    await waitForState(preview.url, (state) => state.revision !== beforeTyping);
    const saved = await loadPresentation(
      new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
    );
    const notesPart = _internalPackageOf(saved).parts.find(
      (part) => part.name === '/ppt/notesSlides/notesSlide1.xml',
    );
    const xml = new TextDecoder().decode(notesPart.data);
    assert.match(xml, /<a:rPr[^>]*\bb="1"/);
    assert.equal(getSlideNotes(getSlides(saved)[0]), 'FirstAb');
  } finally {
    await browser?.close();
    await preview?.close();
    await rm(dir, { recursive: true, force: true });
  }
});

test(
  'empty notes caret formatting creates four native-style undo steps',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-notes-empty-undo-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide} from '@office-kit/pptx-dsl';export default <Presentation><Slide/></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await openNotes(editor);
      const input = editor.getByRole('textbox', { name: 'Notes content', exact: true });
      const state = () => waitForState(preview.url, () => true);
      const waitRevision = async (before) =>
        waitForState(preview.url, (s) => s.revision !== before);
      const notes = async () =>
        getSlides(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
        )[0];

      await input.press('End');
      assert.equal(await input.evaluate((element) => element === document.activeElement), true);
      let before = (await state()).revision;
      await editor.getByRole('button', { name: 'Bold', exact: true }).click();
      await page.waitForTimeout(1000);
      await waitRevision(before);
      before = (await state()).revision;
      await editor.getByRole('button', { name: 'Italic', exact: true }).click();
      await waitRevision(before);
      before = (await state()).revision;
      await input.pressSequentially('Ab');
      await waitRevision(before);
      before = (await state()).revision;
      await input.pressSequentially('Cd');
      await waitRevision(before);
      assert.equal(getSlideNotes(await notes()), 'AbCd');
      assert.deepEqual(
        getSlideNotesTextFormats(await notes()).map(({ start, end, format }) => ({
          start,
          end,
          bold: format.bold,
          italic: format.italic,
        })),
        [
          { start: 0, end: 1, bold: true, italic: true },
          { start: 1, end: 2, bold: true, italic: true },
          { start: 2, end: 3, bold: true, italic: true },
          { start: 3, end: 4, bold: true, italic: true },
        ],
      );

      await input.press('Meta+z');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal(getSlideNotes(await notes()), 'Ab');
      assert.equal(
        getSlideNotesTextFormats(await notes()).every(({ format }) => format.bold && format.italic),
        true,
      );

      await input.press('Meta+z');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal(getSlideNotes(await notes()), '');
      assert.equal(getSlideNotesParagraphEndFormat(await notes(), 0).bold, true);
      assert.equal(getSlideNotesParagraphEndFormat(await notes(), 0).italic, true);

      await input.press('Meta+z');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal(getSlideNotesParagraphEndFormat(await notes(), 0).bold, true);
      assert.equal(getSlideNotesParagraphEndFormat(await notes(), 0).italic, undefined);

      await input.press('Meta+z');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal(getSlideNotes(await notes()), null);
      assert.deepEqual(getSlideNotesTextFormats(await notes()), []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'undoing nonempty caret formatting clears the pending style before retyping',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-notes-caret-undo-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide} from '@office-kit/pptx-dsl';export default <Presentation><Slide notes="Ab"/></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await openNotes(editor);
      const input = editor.getByRole('textbox', { name: 'Notes content', exact: true });
      const state = () => waitForState(preview.url, () => true);
      const notes = async () =>
        getSlides(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
        )[0];
      await input.press('End');
      await editor.getByRole('button', { name: 'Bold', exact: true }).click();
      let before = (await state()).revision;
      await input.pressSequentially('X');
      await waitForState(preview.url, (s) => s.revision !== before);
      assert.equal(getSlideNotes(await notes()), 'AbX');
      before = (await state()).revision;
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await waitForState(preview.url, (s) => s.revision !== before);
      assert.equal(getSlideNotes(await notes()), 'Ab');
      before = (await state()).revision;
      await input.pressSequentially('Y');
      await waitForState(preview.url, (s) => s.revision !== before);
      assert.equal(getSlideNotes(await notes()), 'AbY');
      const y = getSlideNotesTextFormats(await notes()).find(
        ({ start, end }) => start <= 2 && end >= 3,
      );
      assert.equal(y?.format.bold, undefined);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'an empty paragraph in nonempty notes keeps its own caret formatting history',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-notes-empty-paragraph-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide} from '@office-kit/pptx-dsl';export default <Presentation><Slide notes=""/></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await openNotes(editor);
      const input = editor.getByRole('textbox', { name: 'Notes content', exact: true });
      await input.pressSequentially('A');
      await input.press('Enter');
      await input.press('ArrowDown');
      await editor.getByRole('button', { name: 'Bold', exact: true }).click();
      await input.pressSequentially('X');
      await input.press('Tab');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const readSlide = async () =>
        getSlides(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
        )[0];
      let saved = await readSlide();
      assert.equal(getSlideNotes(saved), 'A\nX');
      assert.equal(getSlideNotesParagraphEndFormat(saved, 1).bold, true);
      let before = (await waitForState(preview.url, () => true)).revision;
      await input.press('Meta+z');
      await waitForState(preview.url, (state) => state.revision !== before);
      saved = await readSlide();
      assert.equal(getSlideNotes(saved), 'A\n');
      assert.equal(getSlideNotesParagraphEndFormat(saved, 1).bold, true);
      before = (await waitForState(preview.url, () => true)).revision;
      await input.press('Meta+z');
      await waitForState(preview.url, (state) => state.revision !== before);
      saved = await readSlide();
      assert.equal(getSlideNotes(saved), 'A\n');
      assert.equal(getSlideNotesParagraphEndFormat(saved, 1).bold, undefined);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'composition end schedules notes autosave without a follow-up input',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-notes-composition-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide} from '@office-kit/pptx-dsl';export default <Presentation><Slide notes="A"/></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await openNotes(editor);
      const input = editor.getByRole('textbox', { name: 'Notes content', exact: true });
      await input.press('End');
      const before = (await waitForState(preview.url, () => true)).revision;
      await input.evaluate((node) => {
        node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
        node.textContent = 'A候';
        const selection = window.getSelection();
        const range = document.createRange();
        range.selectNodeContents(node);
        range.collapse(false);
        selection?.removeAllRanges();
        selection?.addRange(range);
        node.dispatchEvent(
          new InputEvent('input', {
            bubbles: true,
            inputType: 'insertCompositionText',
            isComposing: true,
          }),
        );
        node.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '候' }));
      });
      await waitForState(preview.url, (state) => state.revision !== before);
      const saved = await loadPresentation(
        new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
      );
      assert.equal(getSlideNotes(getSlides(saved)[0]), 'A候');
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'leaving focused notes routes Home formatting to the selected slide',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-notes-routing-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide notes="Speaker note"><Text x={1} y={1} width={6} height={1}>Slide text</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      // Open the Format pane first; it follows the selection to the shape later.
      await editor.locator('.stage').click({ button: 'right', position: { x: 8, y: 8 } });
      await editor.getByRole('menuitem', { name: 'Format Background...', exact: true }).click();
      await openNotes(editor);
      const notes = editor.getByRole('textbox', { name: 'Notes content', exact: true });
      assert.equal(await notes.evaluate((element) => element === document.activeElement), true);
      // Selecting the slide while Notes still exists must release the notes
      // formatting target before Home receives the next command.
      await editor.locator('.hit').first().click();
      assert.equal(await editor.locator('.hit.selected').count(), 1);
      await editor.getByRole('tab', { name: 'Size & Properties', exact: true }).click();
      const bar = editor.locator('.ribbon .font-ribbon');
      const beforeBold = await waitForState(preview.url, () => true);
      await bar.getByRole('button', { name: 'Bold', exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const afterBold = await waitForState(preview.url, () => true);
      assert.notEqual(
        afterBold.revision,
        beforeBold.revision,
        'Home Bold creates a slide revision',
      );
      const pres = await loadPresentation(
        new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
      );
      const runs = getShapeParagraphElements(getSlideShapes(getSlides(pres)[0])[0], 0).filter(
        (run) => run.kind === 'r',
      );
      assert.equal(runs[0]?.format?.bold, true, 'Home Bold formats the selected slide shape');
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'notes pane commits drafts to their slide across selection and view changes',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-notes-pane-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide} from '@office-kit/pptx-dsl';export default <Presentation><Slide notes="First"/><Slide notes="Second"/></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const state = () => waitForState(preview.url, () => true);
      const notes = async () =>
        getSlides(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
        ).map(getSlideNotes);
      const revision = (await state()).revision;
      await openNotes(editor);
      const input = editor.getByRole('textbox', { name: 'Notes content', exact: true });
      const inputValue = () => input.textContent();
      assert.equal(await inputValue(), 'First');
      assert.equal(await input.evaluate((el) => el === document.activeElement), true);
      const resize = editor.getByRole('separator', { name: 'Notes pane height' });
      const height = Number(await resize.getAttribute('aria-valuenow'));
      await resize.focus();
      await resize.press('ArrowUp');
      assert.equal(Number(await resize.getAttribute('aria-valuenow')), height + 10);
      await resize.press('End');
      assert.equal(
        Math.round((await editor.locator('.notes-pane').boundingBox()).height),
        Number(await resize.getAttribute('aria-valuenow')),
        'notes splitter reports the visible height at its upper limit',
      );
      await page.setViewportSize({ width: 1500, height: 700 });
      await editor
        .locator('.notes-pane')
        .evaluate(
          () =>
            new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))),
        );
      assert.equal(
        Math.round((await editor.locator('.notes-pane').boundingBox()).height),
        Number(await resize.getAttribute('aria-valuenow')),
        'shrinking the window also updates the splitter value',
      );
      await page.setViewportSize({ width: 1500, height: 1000 });
      await resize.press('Home');
      assert.equal((await state()).revision, revision);
      await input.fill('First edited\n日本語');
      await editor.locator('.nav [data-slide-index="1"]').click();
      assert.equal(await inputValue(), 'Second');
      assert.equal(await input.evaluate((el) => el === document.activeElement), false);
      await waitForState(preview.url, (s) => s.revision !== revision);
      assert.deepEqual(await notes(), ['First edited\n日本語', 'Second']);
      const secondRevision = (await state()).revision;
      await input.fill('Second edited');
      await input.press('Meta+2');
      await editor.locator('.nav.sorter').waitFor();
      await waitForState(preview.url, (s) => s.revision !== secondRevision);
      assert.deepEqual(await notes(), ['First edited\n日本語', 'Second edited']);
      await editor.getByRole('button', { name: 'Normal', exact: true }).click();
      assert.equal(await inputValue(), 'Second edited');
      await input.fill('Pending draft');
      await input.press('Meta+z');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal(await inputValue(), 'Second edited');
      await input.press('Meta+Shift+z');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal(await inputValue(), 'Pending draft');
      const hideRevision = (await state()).revision;
      await input.fill('Saved on hide');
      await editor.getByRole('button', { name: 'Notes', exact: true }).click();
      await input.waitFor({ state: 'detached' });
      await waitForState(preview.url, (s) => s.revision !== hideRevision);
      assert.deepEqual(await notes(), ['First edited\n日本語', 'Saved on hide']);
      await editor.getByRole('button', { name: 'Notes', exact: true }).click();
      const idleRevision = (await state()).revision;
      await input.fill('Saved while focused');
      await waitForState(preview.url, (s) => s.revision !== idleRevision);
      assert.deepEqual(await notes(), ['First edited\n日本語', 'Saved while focused']);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'Home ribbon formats selected speaker notes and undo restores the run XML',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-notes-format-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide} from '@office-kit/pptx-dsl';export default <Presentation><Slide notes="First"/></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const state = () => waitForState(preview.url, () => true);
      await openNotes(editor);
      const input = editor.getByRole('textbox', { name: 'Notes content', exact: true });
      await input.selectText();
      const before = (await state()).revision;
      await editor.getByRole('button', { name: 'Bold', exact: true }).click();
      await waitForState(preview.url, (s) => s.revision !== before);
      const formatted = await loadPresentation(
        new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
      );
      const formattedXml = new TextDecoder().decode(
        _internalPackageOf(formatted).getPart('/ppt/notesSlides/notesSlide1.xml').data,
      );
      assert.match(formattedXml, /<a:rPr[^>]*\bb="1"/);
      const afterFormat = (await state()).revision;
      await input.press('Meta+z');
      await waitForState(preview.url, (s) => s.revision !== afterFormat);
      const undone = await loadPresentation(
        new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
      );
      const undoneXml = new TextDecoder().decode(
        _internalPackageOf(undone).getPart('/ppt/notesSlides/notesSlide1.xml').data,
      );
      assert.doesNotMatch(undoneXml, /<a:rPr[^>]*\bb="1"/);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

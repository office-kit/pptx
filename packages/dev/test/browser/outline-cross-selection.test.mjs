import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile, copyFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { loadPresentation, getSlides, getSlideShapes, getShapeText } from '@office-kit/pptx';
import { installRichTextSelection } from '../helpers/rich-text.mjs';
import { startPreview, waitForState } from '../helpers/server.mjs';

async function writeOutlineDeck(dir) {
  await copyFile(
    new URL('../../../../test/fixtures/minimal/one-text-slide.pptx', import.meta.url),
    join(dir, 'template.pptx'),
  );
  await writeFile(
    join(dir, 'deck.tsx'),
    `import {readFileSync} from 'node:fs'; import {Presentation,Slide,Fill} from '@office-kit/pptx-dsl'; const source = new Uint8Array(readFileSync(new URL('./template.pptx', import.meta.url))); export default <Presentation source={source} mode="compose"><Slide layout={{name:"Title and Content"}}><Fill target={{placeholder:{type:'title'}}}>Heading</Fill><Fill target={{placeholder:{idx:1}}}>Body</Fill></Slide><Slide layout={{name:"Title and Content"}}><Fill target={{placeholder:{type:'title'}}}>Next</Fill><Fill target={{placeholder:{idx:1}}}>Following</Fill></Slide></Presentation>;`,
  );
}

async function openOutline(dir) {
  const preview = await startPreview(join(dir, 'deck.tsx'));
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
  await installRichTextSelection(page);
  await page.goto(preview.url);
  await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
  const editor = page.frameLocator('#editor-frame');
  await editor.getByRole('tab', { name: 'View', exact: true }).click();
  await editor
    .getByRole('tabpanel', { name: 'View', exact: true })
    .getByRole('button', { name: 'Outline View', exact: true })
    .click();
  const outline = editor.getByRole('navigation', { name: 'Outline View', exact: true });
  return { preview, browser, page, outline };
}

async function extendForward(page, title) {
  await title.focus();
  await title.evaluate((input) => window.selectEditorText(input, 2, input.textContent.length));
  await title.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+ArrowDown');
}

async function readDeck(preview) {
  return loadPresentation(
    new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
  );
}

test(
  'outline selection copies and cuts a range across title and body fields',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-outline-range-'));
    let preview, browser;
    try {
      await copyFile(
        new URL('../../../../test/fixtures/minimal/one-text-slide.pptx', import.meta.url),
        join(dir, 'template.pptx'),
      );
      await writeFile(
        join(dir, 'deck.tsx'),
        `import {readFileSync} from 'node:fs'; import {Presentation,Slide,Fill} from '@office-kit/pptx-dsl'; const source = new Uint8Array(readFileSync(new URL('./template.pptx', import.meta.url))); export default <Presentation source={source} mode="compose"><Slide layout={{name:"Title and Content"}}><Fill target={{placeholder:{type:'title'}}}>Heading</Fill><Fill target={{placeholder:{idx:1}}}>Body</Fill></Slide><Slide layout={{name:"Title and Content"}}><Fill target={{placeholder:{type:'title'}}}>Next</Fill><Fill target={{placeholder:{idx:1}}}>Following</Fill></Slide></Presentation>;`,
      );
      preview = await startPreview(join(dir, 'deck.tsx'));
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
      await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByRole('tab', { name: 'View', exact: true }).click();
      await editor
        .getByRole('tabpanel', { name: 'View', exact: true })
        .getByRole('button', { name: 'Outline View', exact: true })
        .click();
      const outline = editor.getByRole('navigation', { name: 'Outline View', exact: true });
      const title = outline.getByRole('textbox').nth(0);
      const body = outline.getByRole('textbox').nth(1);
      await title.focus();
      await title.evaluate((input) => {
        window.selectEditorText(input, input.textContent.length, input.textContent.length);
        input.dispatchEvent(
          new InputEvent('beforeinput', { bubbles: true, inputType: 'insertText', data: 'X' }),
        );
        input.textContent += 'X';
        window.selectEditorText(input, input.textContent.length, input.textContent.length);
        input.dispatchEvent(
          new InputEvent('input', { bubbles: true, inputType: 'insertText', data: 'X' }),
        );
        window.selectEditorText(input, 2, input.textContent.length);
      });
      await title.press('Shift+ArrowDown');
      // Continue from the active body field.  Calling locator.press on another
      // textbox would focus it first and reset the shared anchor.
      await page.keyboard.press('Shift+ArrowDown');
      await page.keyboard.press('Shift+ArrowDown');
      await page.keyboard.press('Shift+ArrowDown');
      const copied = await body.evaluate((input) => {
        const data = new DataTransfer();
        input.dispatchEvent(
          new ClipboardEvent('copy', { clipboardData: data, bubbles: true, cancelable: true }),
        );
        return data.getData('text/plain');
      });
      assert.equal(copied, 'adingX\nBody\nNext');
      const before = (await waitForState(preview.url, () => true)).revision;
      await body.evaluate((input) => {
        const data = new DataTransfer();
        input.dispatchEvent(
          new ClipboardEvent('cut', { clipboardData: data, bubbles: true, cancelable: true }),
        );
      });
      const cutState = await waitForState(preview.url, (state) => state.revision !== before);
      const read = async () =>
        loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
      let slides = getSlides(await read());
      let shapes = getSlideShapes(slides[0]);
      assert.equal(getShapeText(shapes.find((shape) => getShapeText(shape) === 'He')), 'He');
      assert.ok(!shapes.some((shape) => getShapeText(shape) === 'Body'));
      assert.ok(
        !getSlideShapes(getSlides(await read())[1]).some((shape) => getShapeText(shape) === 'Next'),
      );
      assert.ok(
        getSlideShapes(getSlides(await read())[1]).some(
          (shape) => getShapeText(shape) === 'Following',
        ),
      );
      await body.press('Control+z');
      await waitForState(preview.url, (state) => state.revision !== cutState.revision);
      slides = getSlides(await read());
      shapes = getSlideShapes(slides[0]);
      assert.ok(shapes.some((shape) => getShapeText(shape) === 'Heading'));
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'outline cross-field selection replaces the range with normal typing',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-outline-typing-'));
    let preview, browser, page, outline;
    try {
      await writeOutlineDeck(dir);
      ({ preview, browser, page, outline } = await openOutline(dir));
      const title = outline.getByRole('textbox').nth(0);
      await extendForward(page, title);
      const before = (await waitForState(preview.url, () => true)).revision;
      await page.keyboard.press('Z');
      await waitForState(preview.url, (state) => state.revision !== before);
      await new Promise((resolve) => setTimeout(resolve, 1000));
      const afterZ = (await waitForState(preview.url, () => true)).revision;
      await page.keyboard.type('XY');
      await waitForState(preview.url, (state) => state.revision !== afterZ);
      await new Promise((resolve) => setTimeout(resolve, 700));
      let slides = getSlides(await readDeck(preview));
      let texts = getSlideShapes(slides[0]).map((shape) => getShapeText(shape));
      const busyAfterTyping = await title.getAttribute('aria-busy');
      assert.ok(
        texts.includes('HeZXY'),
        `expected subsequent typing at the replacement caret (busy=${busyAfterTyping}): ${texts.join('|')}`,
      );
      const body = outline.getByRole('textbox').nth(1);
      await title.focus();
      await extendForward(page, title);
      await body.click();
      const beforeQ = (await waitForState(preview.url, () => true)).revision;
      await page.keyboard.press('Q');
      await waitForState(preview.url, (state) => state.revision !== beforeQ);
      await new Promise((resolve) => setTimeout(resolve, 700));
      slides = getSlides(await readDeck(preview));
      texts = getSlideShapes(slides[0]).map((shape) => getShapeText(shape));
      assert.ok(
        texts.includes('HeZXY'),
        `pointer click must clear the cross-field range: ${texts.join('|')}`,
      );
      assert.ok(
        texts.some((text) => text.includes('Q')),
        `pointer click should leave a local insertion: ${texts.join('|')}`,
      );
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'outline replacement does not reclaim focus from the ribbon before caret restoration',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-outline-ribbon-focus-'));
    let preview, browser, page, outline;
    try {
      await writeOutlineDeck(dir);
      ({ preview, browser, page, outline } = await openOutline(dir));
      const editor = page.frameLocator('#editor-frame');
      await editor.getByRole('tab', { name: 'Home', exact: true }).click();
      await page.evaluate(() => {
        const frame = document.querySelector('#editor-frame');
        if (!(frame instanceof HTMLIFrameElement) || !frame.contentWindow)
          throw new Error('editor frame is unavailable');
        const win = frame.contentWindow;
        const native = win.requestAnimationFrame.bind(win);
        const queue = [];
        win.__outlineTestRaf = { native, queue, holding: true };
        win.requestAnimationFrame = (callback) => {
          if (!win.__outlineTestRaf.holding) return native(callback);
          queue.push(callback);
          return queue.length;
        };
      });
      const title = outline.getByRole('textbox').nth(0);
      await extendForward(page, title);
      const before = (await waitForState(preview.url, () => true)).revision;
      await page.keyboard.press('Z');
      await waitForState(preview.url, (state) => state.revision !== before);
      assert.ok(
        await page.evaluate(
          () => document.querySelector('#editor-frame').contentWindow.__outlineTestRaf.queue.length,
        ),
      );

      const ribbonControl = editor.getByRole('button', { name: 'Font options', exact: true });
      await ribbonControl.focus();
      assert.equal(
        await ribbonControl.evaluate((node) => node === node.ownerDocument.activeElement),
        true,
      );

      await page.evaluate(() => {
        const win = document.querySelector('#editor-frame').contentWindow;
        const state = win.__outlineTestRaf;
        state.holding = false;
        for (const callback of state.queue.splice(0)) callback(performance.now());
        win.requestAnimationFrame = state.native;
      });
      assert.equal(
        await ribbonControl.evaluate((node) => node === node.ownerDocument.activeElement),
        true,
      );
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'outline cross-field selection replaces the range from a paste event',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-outline-paste-'));
    let preview, browser, page, outline;
    try {
      await writeOutlineDeck(dir);
      ({ preview, browser, page, outline } = await openOutline(dir));
      const title = outline.getByRole('textbox').nth(0);
      await extendForward(page, title);
      const before = (await waitForState(preview.url, () => true)).revision;
      await title.evaluate((input) => {
        const data = new DataTransfer();
        data.setData('text/plain', 'P');
        input.dispatchEvent(
          new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }),
        );
      });
      await waitForState(preview.url, (state) => state.revision !== before);
      const slides = getSlides(await readDeck(preview));
      const texts = getSlideShapes(slides[0]).map((shape) => getShapeText(shape));
      assert.ok(
        texts.includes('HeP'),
        `expected paste to replace the cross-field range at its anchor: ${texts.join('|')}`,
      );
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'outline selection preserves a backward cross-field range for copy',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-outline-backward-'));
    let preview, browser, page, outline;
    try {
      await writeOutlineDeck(dir);
      ({ preview, browser, page, outline } = await openOutline(dir));
      const nextTitle = outline.getByRole('textbox').nth(2);
      await nextTitle.focus();
      await nextTitle.evaluate((input) => window.selectEditorText(input, 0, 0));
      await nextTitle.press('Shift+ArrowUp');
      await page.keyboard.press('Shift+ArrowUp');
      await page.keyboard.press('Shift+ArrowUp');
      await page.keyboard.press('Shift+ArrowUp');
      await page.keyboard.press('Shift+ArrowUp');
      const copied = await nextTitle.evaluate((input) => {
        const data = new DataTransfer();
        input.dispatchEvent(
          new ClipboardEvent('copy', { clipboardData: data, bubbles: true, cancelable: true }),
        );
        return data.getData('text/plain');
      });
      // The reversed range now includes the preceding title and body fields.
      assert.equal(copied, 'Heading\nBody\n');
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'outline cross-field selection handles Shift+Enter as one replacement',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-outline-shift-enter-'));
    let preview, browser, page, outline;
    try {
      await writeOutlineDeck(dir);
      ({ preview, browser, page, outline } = await openOutline(dir));
      const title = outline.getByRole('textbox').nth(0);
      await extendForward(page, title);
      const before = (await waitForState(preview.url, () => true)).revision;
      await page.keyboard.press('Shift+Enter');
      await waitForState(preview.url, (state) => state.revision !== before);
      const texts = getSlideShapes(getSlides(await readDeck(preview))[0]).map((shape) =>
        getShapeText(shape),
      );
      assert.ok(
        texts.includes('He\n'),
        `expected one cross-field line-break replacement: ${texts.join('|')}`,
      );
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test('outline cross-field selection handles committed IME text', { timeout: 60000 }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-outline-ime-'));
  let preview, browser, page, outline;
  try {
    await writeOutlineDeck(dir);
    ({ preview, browser, page, outline } = await openOutline(dir));
    const title = outline.getByRole('textbox').nth(0);
    await extendForward(page, title);
    const cdp = await page.context().newCDPSession(page);
    const before = (await waitForState(preview.url, () => true)).revision;
    await cdp.send('Input.imeSetComposition', { text: 'か', selectionStart: 1, selectionEnd: 1 });
    await cdp.send('Input.imeSetComposition', { text: '漢', selectionStart: 1, selectionEnd: 1 });
    await cdp.send('Input.insertText', { text: '漢' });
    await waitForState(preview.url, (state) => state.revision !== before);
    const texts = getSlides(await readDeck(preview)).flatMap((slide) =>
      getSlideShapes(slide).map((shape) => getShapeText(shape)),
    );
    assert.deepEqual(
      texts.filter(Boolean),
      ['He漢', 'Following'],
      `expected committed IME replacement across all four fields: ${texts.join('|')}`,
    );
    const undoBefore = (await waitForState(preview.url, () => true)).revision;
    await title.press('Control+z');
    await waitForState(preview.url, (state) => state.revision !== undoBefore);
    const restored = getSlides(await readDeck(preview)).flatMap((slide) =>
      getSlideShapes(slide).map((shape) => getShapeText(shape)),
    );
    assert.deepEqual(
      restored.filter(Boolean),
      ['Heading', 'Body', 'Next', 'Following'],
      `IME replacement should undo all four fields as one edit: ${restored.join('|')}`,
    );
  } finally {
    await browser?.close();
    await preview?.close();
    await rm(dir, { recursive: true, force: true });
  }
});

test(
  'outline cross-field selection collapses to native edges with horizontal arrows',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-outline-collapse-'));
    let preview, browser, page, outline;
    try {
      await writeOutlineDeck(dir);
      ({ preview, browser, page, outline } = await openOutline(dir));
      const title = outline.getByRole('textbox').nth(0);
      await extendForward(page, title);
      const leftBefore = (await waitForState(preview.url, () => true)).revision;
      await page.keyboard.press('ArrowLeft');
      await page.keyboard.type('L');
      await waitForState(preview.url, (state) => state.revision !== leftBefore);
      await new Promise((resolve) => setTimeout(resolve, 700));
      let texts = getSlideShapes(getSlides(await readDeck(preview))[0]).map((shape) =>
        getShapeText(shape),
      );
      assert.ok(
        texts.includes('HeLading'),
        `left collapse should insert at range start: ${texts.join('|')}`,
      );
      await browser.close();
      await preview.close();
      browser = undefined;
      preview = undefined;

      ({ preview, browser, page, outline } = await openOutline(dir));
      const rightTitle = outline.getByRole('textbox').nth(0);
      await extendForward(page, rightTitle);
      const rightBefore = (await waitForState(preview.url, () => true)).revision;
      await page.keyboard.press('ArrowRight');
      await page.keyboard.type('R');
      await waitForState(preview.url, (state) => state.revision !== rightBefore);
      await new Promise((resolve) => setTimeout(resolve, 700));
      texts = getSlides(await readDeck(preview)).flatMap((slide) =>
        getSlideShapes(slide).map((shape) => getShapeText(shape)),
      );
      assert.ok(
        texts.includes('NextR'),
        `right collapse should insert at logical range end: ${texts.join('|')}`,
      );
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test('outline right-click preserves a cross-field range for copy', { timeout: 60000 }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-outline-context-copy-'));
  let preview, browser, page, outline;
  try {
    await writeOutlineDeck(dir);
    ({ preview, browser, page, outline } = await openOutline(dir));
    const title = outline.getByRole('textbox').nth(0);
    const body = outline.getByRole('textbox').nth(1);
    await extendForward(page, title);
    await body.click({ button: 'right' });
    const editor = page.frameLocator('#editor-frame');
    await editor.getByRole('menuitem', { name: /^Copy/ }).click();
    await page.waitForFunction(
      async () => (await navigator.clipboard.readText()) === 'ading\nBody\nNext',
      undefined,
      { timeout: 5000 },
    );
    assert.equal(await page.evaluate(() => navigator.clipboard.readText()), 'ading\nBody\nNext');
  } finally {
    await browser?.close();
    await preview?.close();
    await rm(dir, { recursive: true, force: true });
  }
});

import assert from 'node:assert/strict';
import { copyFile, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getParagraphLevel,
  getShapeText,
  getSlides,
  getSlideShapes,
  getSlideText,
  loadPresentation,
} from '@office-kit/pptx';
import { installRichTextSelection } from '../helpers/rich-text.mjs';
import { startPreview, waitForState } from '../helpers/server.mjs';

const LABELS = {
  en: { view: 'View', outline: 'Outline View' },
  ja: { view: '表示', outline: 'アウトライン表示' },
};

/** Three Title and Content slides; `levels` sets body paragraph levels per slide. */
async function openDeck(locale, slides) {
  const dir = await mkdtemp(join(tmpdir(), 'office-outline-cross-slide-'));
  await copyFile(
    new URL('../../../../test/fixtures/minimal/one-text-slide.pptx', import.meta.url),
    join(dir, 'template.pptx'),
  );
  const markup = slides
    .map(
      ([title, body, levels = []]) =>
        `<Slide layout={{name:'Title and Content'}}><Fill target={{placeholder:{type:'title'}}}>{${JSON.stringify(title)}}</Fill><Fill target={{placeholder:{idx:1}}}>{${JSON.stringify(body)}}</Fill><Raw scope="slide" apply={({slide}) => { ${JSON.stringify(levels)}.forEach((level, index) => setParagraphLevel(getSlideShapes(slide)[1], index, level)); }} /></Slide>`,
    )
    .join('');
  await writeFile(
    join(dir, 'deck.tsx'),
    `import {readFileSync} from 'node:fs'; import {getSlideShapes,setParagraphLevel} from '@office-kit/pptx'; import {Presentation,Slide,Fill,Raw} from '@office-kit/pptx-dsl'; const source = new Uint8Array(readFileSync(new URL('./template.pptx', import.meta.url))); export default <Presentation source={source} mode="compose">${markup}</Presentation>;`,
  );
  const preview = await startPreview(join(dir, 'deck.tsx'));
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  const errors = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.addInitScript(
    (language) => localStorage.setItem('ok-editor-locale', language),
    locale,
  );
  await installRichTextSelection(page);
  await page.goto(preview.url);
  const editor = page.frameLocator('#editor-frame');
  const labels = LABELS[locale];
  await editor.getByRole('tab', { name: labels.view, exact: true }).click();
  await editor
    .getByRole('tabpanel', { name: labels.view, exact: true })
    .getByRole('button', { name: labels.outline, exact: true })
    .click();
  const outline = editor.getByRole('navigation', { name: labels.outline, exact: true });
  const frame = await page.locator('#editor-frame').boundingBox();

  /** Page coordinates of a UTF-16 caret offset inside an outline textbox. */
  const pointOf = async (box, offset) => {
    const point = await box.evaluate((root, target) => {
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      let seen = 0;
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        const length = node.textContent.length;
        // Newlines between paragraph sections are literal text nodes.
        if (node.parentElement === root) {
          seen += length;
          continue;
        }
        if (seen + length >= target) {
          const local = target - seen;
          const range = document.createRange();
          const after = local < length;
          range.setStart(node, after ? local : local - 1);
          range.setEnd(node, after ? local + 1 : local);
          const rect = range.getBoundingClientRect();
          return { x: after ? rect.left + 0.5 : rect.right - 0.5, y: rect.top + rect.height / 2 };
        }
        seen += length;
      }
      throw new Error(`offset ${target} is outside the textbox`);
    }, offset);
    return { x: frame.x + point.x, y: frame.y + point.y, local: point };
  };
  // Headless Chromium does not start native drags of selected text, so text
  // drags are replayed as the drag events the outline handles.
  const dragText = (source, target, copy = false) =>
    source.evaluate(
      (root, { point, copy }) => {
        const data = new DataTransfer();
        const init = { bubbles: true, cancelable: true, composed: true, dataTransfer: data };
        root.dispatchEvent(new DragEvent('dragstart', init));
        const over = root.getRootNode().elementFromPoint(point.x, point.y);
        const at = { ...init, clientX: point.x, clientY: point.y, altKey: copy };
        over.dispatchEvent(new DragEvent('dragover', at));
        over.dispatchEvent(new DragEvent('drop', at));
        root.dispatchEvent(new DragEvent('dragend', init));
      },
      { point: target.local, copy },
    );
  const read = async () =>
    loadPresentation(new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()));
  const texts = async () => getSlides(await read()).map((slide) => getSlideText(slide));
  const change = async (action) => {
    const before = (await waitForState(preview.url, () => true)).revision;
    await action();
    return waitForState(preview.url, (state) => state.revision !== before);
  };
  const textbox = (index) => outline.getByRole('textbox').nth(index);
  const close = async () => {
    await browser.close();
    await preview.close();
    await rm(dir, { recursive: true, force: true });
  };
  return { page, outline, editor, pointOf, dragText, read, texts, change, textbox, errors, close };
}

const DECK = [
  ['One', 'Alpha\nBeta'],
  ['Two', 'Gamma'],
  ['Three', 'Delta'],
];

for (const locale of ['en', 'ja']) {
  test(
    `outline pointer drag selects across slides and typing merges them in one step (${locale})`,
    { timeout: 90000 },
    async () => {
      const deck = await openDeck(locale, DECK);
      try {
        const { page, pointOf, textbox, texts, change } = deck;
        // Slide 1 body "Alpha\nB|eta" to slide 3 title "Th|ree".
        const from = await pointOf(textbox(1), 7);
        const to = await pointOf(textbox(4), 2);
        await page.mouse.move(from.x, from.y);
        await page.mouse.down();
        await page.mouse.move(from.x + 10, from.y, { steps: 2 });
        await page.mouse.move(to.x, to.y, { steps: 8 });
        await page.mouse.up();
        const copied = await textbox(1).evaluate((input) => {
          const data = new DataTransfer();
          input.dispatchEvent(
            new ClipboardEvent('copy', { clipboardData: data, bubbles: true, cancelable: true }),
          );
          return data.getData('text/plain');
        });
        assert.equal(copied, 'eta\nTwo\nGamma\nTh');
        const typed = await change(() => page.keyboard.press('X'));
        assert.deepEqual(await texts(), ['One\nAlpha\nBXree\nDelta']);
        await change(() => page.keyboard.press('Control+z'));
        assert.ok(typed.revision);
        assert.deepEqual(await texts(), ['One\nAlpha\nBeta', 'Two\nGamma', 'Three\nDelta']);
        assert.deepEqual(deck.errors, []);
      } finally {
        await deck.close();
      }
    },
  );
}

test(
  'outline Shift-click, boundary Backspace/Delete and Enter edit across slides',
  { timeout: 90000 },
  async () => {
    const deck = await openDeck('en', DECK);
    try {
      const { page, pointOf, textbox, texts, change } = deck;
      // Caret in slide 1 body after "Al", then Shift-click slide 2 body after "Gam".
      await textbox(1).focus();
      await textbox(1).evaluate((input) => window.selectEditorText(input, 2, 2));
      const target = await pointOf(textbox(3), 3);
      await page.keyboard.down('Shift');
      await page.mouse.click(target.x, target.y);
      await page.keyboard.up('Shift');
      await change(() => page.keyboard.press('Delete'));
      assert.deepEqual(await texts(), ['One\nAlma', 'Three\nDelta']);
      await change(() => page.keyboard.press('Control+z'));
      assert.deepEqual(await texts(), ['One\nAlpha\nBeta', 'Two\nGamma', 'Three\nDelta']);

      // Backspace at the start of slide 2's title joins it to slide 1's last paragraph.
      await textbox(2).focus();
      await textbox(2).evaluate((input) => window.selectEditorText(input, 0, 0));
      await change(() => page.keyboard.press('Backspace'));
      assert.deepEqual(await texts(), ['One\nAlpha\nBetaTwo\nGamma', 'Three\nDelta']);
      // Delete at the end of a body pulls the next slide's title in.
      const body = textbox(1);
      await body.focus();
      await body.evaluate((input) => window.selectEditorText(input, input.textContent.length));
      await change(() => page.keyboard.press('Delete'));
      assert.deepEqual(await texts(), ['One\nAlpha\nBetaTwo\nGammaThree\nDelta']);
      await change(() => page.keyboard.press('Control+z'));
      await change(() => page.keyboard.press('Control+z'));
      assert.deepEqual(await texts(), ['One\nAlpha\nBeta', 'Two\nGamma', 'Three\nDelta']);

      // Enter over a body-to-body range across slides removes the range and
      // starts a new paragraph at the caret.
      await textbox(1).focus();
      await textbox(1).evaluate((input) => window.selectEditorText(input, 7, 7));
      const end = await pointOf(textbox(3), 2);
      await page.keyboard.down('Shift');
      await page.mouse.click(end.x, end.y);
      await page.keyboard.up('Shift');
      await change(() => page.keyboard.press('Enter'));
      assert.deepEqual(await texts(), ['One\nAlpha\nB\nmma', 'Three\nDelta']);
      assert.deepEqual(deck.errors, []);
    } finally {
      await deck.close();
    }
  },
);

for (const locale of ['en', 'ja']) {
  test(
    `outline bullet drags move paragraphs across slides and change levels (${locale})`,
    { timeout: 90000 },
    async () => {
      const deck = await openDeck(locale, [
        ['One', 'Alpha\nChild\nBeta', [0, 1, 0]],
        ['Two', 'Gamma'],
      ]);
      try {
        const { page, textbox, texts, change, read } = deck;
        const bullet = async (box, paragraph) => {
          const section = box.locator('[data-outline-paragraph]').nth(paragraph);
          const bounds = await section.boundingBox();
          const gutter = await section.evaluate((node) =>
            parseFloat(getComputedStyle(node).paddingLeft),
          );
          return { x: bounds.x + gutter - 2, y: bounds.y + bounds.height / 2 };
        };
        const levels = (slide, index = 1) => {
          const shape = getSlideShapes(slide)[index];
          return getParagraphLevel(shape, { start: 0, end: getShapeText(shape).length });
        };
        // Drag "Alpha" (with its child) below "Gamma" on slide 2.
        const start = await bullet(textbox(1), 0);
        const gamma = await textbox(3).locator('[data-outline-paragraph]').first().boundingBox();
        await change(async () => {
          await page.mouse.move(start.x, start.y);
          await page.mouse.down();
          await page.mouse.move(start.x, start.y + 10, { steps: 3 });
          await page.mouse.move(start.x, gamma.y + gamma.height + 1, { steps: 10 });
          await page.mouse.up();
        });
        assert.deepEqual(await texts(), ['One\nBeta', 'Two\nGamma\nAlpha\nChild']);
        assert.deepEqual(levels(getSlides(await read())[1]), [0, 0, 1]);
        await change(() => page.keyboard.press('Control+z'));
        assert.deepEqual(await texts(), ['One\nAlpha\nChild\nBeta', 'Two\nGamma']);

        // Drag "Beta" right by two levels in place: one Demote step.
        const beta = await bullet(textbox(1), 2);
        await change(async () => {
          await page.mouse.move(beta.x, beta.y);
          await page.mouse.down();
          await page.mouse.move(beta.x + 24, beta.y, { steps: 6 });
          await page.mouse.up();
        });
        assert.deepEqual(levels(getSlides(await read())[0]), [0, 1, 2]);
        await change(() => page.keyboard.press('Control+z'));

        // Drag a top-level bullet left: it becomes a slide title, as Promote does.
        const alpha = await bullet(textbox(1), 0);
        await change(async () => {
          await page.mouse.move(alpha.x, alpha.y);
          await page.mouse.down();
          await page.mouse.move(alpha.x - 14, alpha.y, { steps: 6 });
          await page.mouse.up();
        });
        const promoted = await texts();
        assert.equal(promoted.length, 3);
        assert.equal(promoted[1].split('\n')[0], 'Alpha');
        await change(() => page.keyboard.press('Control+z'));
        assert.deepEqual(await texts(), ['One\nAlpha\nChild\nBeta', 'Two\nGamma']);
        assert.deepEqual(deck.errors, []);
      } finally {
        await deck.close();
      }
    },
  );
}

test(
  'outline slide icons demote by dragging right and selected text moves or copies by dragging',
  { timeout: 90000 },
  async () => {
    const deck = await openDeck('en', DECK);
    try {
      const { page, outline, pointOf, dragText, textbox, texts, change } = deck;
      const icon = outline.locator('[data-outline-slide="1"] > button');
      const bounds = await icon.boundingBox();
      await change(async () => {
        await page.mouse.move(bounds.x + 10, bounds.y + bounds.height / 2);
        await page.mouse.down();
        await page.mouse.move(bounds.x + 20, bounds.y + bounds.height / 2, { steps: 3 });
        await page.mouse.move(bounds.x + 50, bounds.y + bounds.height / 2 + 2, { steps: 6 });
        await page.mouse.up();
      });
      assert.deepEqual(await texts(), ['One\nAlpha\nBeta\nTwo\nGamma', 'Three\nDelta']);
      await change(() => textbox(1).press('Control+z'));
      assert.deepEqual(await texts(), ['One\nAlpha\nBeta', 'Two\nGamma', 'Three\nDelta']);

      // Select "Gamma" and drag it after "Del" on slide 3.
      await textbox(3).focus();
      await textbox(3).evaluate((input) => window.selectEditorText(input, 0, 5));
      await change(async () => dragText(textbox(3), await pointOf(textbox(5), 3)));
      assert.deepEqual(await texts(), ['One\nAlpha\nBeta', 'Two', 'Three\nDelGammata']);
      await change(() => page.keyboard.press('Control+z'));
      assert.deepEqual(await texts(), ['One\nAlpha\nBeta', 'Two\nGamma', 'Three\nDelta']);

      // The copy modifier leaves the source in place.
      await textbox(1).focus();
      await textbox(1).evaluate((input) => window.selectEditorText(input, 0, 5));
      await change(async () => dragText(textbox(1), await pointOf(textbox(2), 3), true));
      assert.deepEqual(await texts(), ['One\nAlpha\nBeta', 'TwoAlpha\nGamma', 'Three\nDelta']);
      await change(() => page.keyboard.press('Control+z'));

      // A cross-slide range moves as text: its slide boundary is merged away.
      const from = await pointOf(textbox(1), 7);
      const to = await pointOf(textbox(3), 2);
      await page.mouse.move(from.x, from.y);
      await page.mouse.down();
      await page.mouse.move(to.x, to.y, { steps: 8 });
      await page.mouse.up();
      await change(async () => dragText(textbox(1), await pointOf(textbox(5), 5)));
      assert.deepEqual(await texts(), ['One\nAlpha\nBmma', 'Three\nDeltaeta\nTwo\nGa']);
      await change(() => page.keyboard.press('Control+z'));
      assert.deepEqual(await texts(), ['One\nAlpha\nBeta', 'Two\nGamma', 'Three\nDelta']);
      assert.deepEqual(deck.errors, []);
    } finally {
      await deck.close();
    }
  },
);

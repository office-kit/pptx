import assert from 'node:assert/strict';
import { copyFile, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeParagraphElements,
  getShapeText,
  getSlideShapes,
  getSlides,
  loadPresentation,
} from '@office-kit/pptx';
import { installRichTextSelection } from '../helpers/rich-text.mjs';
import { startPreview, waitForState } from '../helpers/server.mjs';

async function openOutline(dir) {
  await copyFile(
    new URL('../../../../test/fixtures/minimal/one-text-slide.pptx', import.meta.url),
    join(dir, 'template.pptx'),
  );
  await writeFile(
    join(dir, 'deck.tsx'),
    `import {readFileSync} from 'node:fs'; import {Presentation,Slide,Fill} from '@office-kit/pptx-dsl'; const source = new Uint8Array(readFileSync(new URL('./template.pptx', import.meta.url))); export default <Presentation source={source} mode="compose"><Slide layout={{name:"Title and Content"}}><Fill target={{placeholder:{type:'title'}}}>Heading</Fill><Fill target={{placeholder:{idx:1}}}>Body</Fill></Slide><Slide layout={{name:"Title and Content"}}><Fill target={{placeholder:{type:'title'}}}>Next</Fill><Fill target={{placeholder:{idx:1}}}>Following</Fill></Slide></Presentation>;`,
  );
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
  await editor.getByRole('tab', { name: 'Home', exact: true }).click();
  return {
    preview,
    browser,
    page,
    editor,
    outline: editor.getByRole('navigation', { name: 'Outline View', exact: true }),
  };
}

async function readDeck(preview) {
  return loadPresentation(
    new Uint8Array(await (await fetch(`${preview.url}/deck.pptx`)).arrayBuffer()),
  );
}

test(
  'outline Ribbon formats only selected spans across fields and undoes as one edit',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-outline-ribbon-'));
    let preview;
    let browser;
    let editor;
    let outline;
    let page;
    try {
      ({ preview, browser, page, editor, outline } = await openOutline(dir));
      const title = outline.getByRole('textbox').nth(0);
      await title.focus();
      await title.evaluate((input) => window.selectEditorText(input, 2, input.textContent.length));
      await title.press('Shift+ArrowDown');
      await page.keyboard.press('Shift+ArrowDown');
      await page.keyboard.press('Shift+ArrowDown');
      await page.keyboard.press('Shift+ArrowDown');
      const before = (await waitForState(preview.url, () => true)).revision;
      const bar = editor.locator('.ribbon .font-ribbon');
      await bar.getByRole('button', { name: 'Bold', exact: true }).click();
      const formatted = await waitForState(preview.url, (state) => state.revision !== before);
      const readRuns = async () => {
        const pres = await readDeck(preview);
        return getSlides(pres).flatMap((slide) =>
          getSlideShapes(slide).map((shape) => ({
            text: getShapeText(shape),
            runs: getShapeParagraphElements(shape, 0),
          })),
        );
      };
      let shapes = await readRuns();
      const heading = shapes.find(({ text }) => text === 'Heading');
      const body = shapes.find(({ text }) => text === 'Body');
      const next = shapes.find(({ text }) => text === 'Next');
      const following = shapes.find(({ text }) => text === 'Following');
      assert.equal(heading.runs[0].text, 'He');
      assert.notEqual(heading.runs[0].format?.bold, true);
      assert.equal(heading.runs[1].text, 'ading');
      assert.equal(heading.runs[1].format?.bold, true);
      assert.ok(body.runs.every((run) => run.format?.bold === true));
      assert.ok(next.runs.every((run) => run.format?.bold === true));
      assert.ok(following.runs.every((run) => run.format?.bold !== true));
      await page.keyboard.press('Control+z');
      await waitForState(preview.url, (state) => state.revision !== formatted.revision);
      shapes = await readRuns();
      assert.ok(
        shapes
          .find(({ text }) => text === 'Heading')
          .runs.every((run) => run.format?.bold !== true),
      );
      assert.ok(
        shapes.find(({ text }) => text === 'Body').runs.every((run) => run.format?.bold !== true),
      );
      assert.ok(
        shapes.find(({ text }) => text === 'Next').runs.every((run) => run.format?.bold !== true),
      );
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'outline Ribbon keeps pending caret formatting for newly typed text',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-outline-caret-format-'));
    let preview;
    let browser;
    let editor;
    let outline;
    try {
      ({ preview, browser, outline, editor } = await openOutline(dir));
      const title = outline.getByRole('textbox').nth(0);
      await title.focus();
      await title.evaluate((input) =>
        window.selectEditorText(input, input.textContent.length, input.textContent.length),
      );
      await editor
        .locator('.ribbon .font-ribbon')
        .getByRole('button', { name: 'Bold', exact: true })
        .click();
      await editor
        .locator('.ribbon .font-ribbon')
        .getByRole('button', { name: 'Italic', exact: true })
        .click();
      const before = (await waitForState(preview.url, () => true)).revision;
      await title.press('Z');
      await waitForState(preview.url, (state) => state.revision !== before);
      const pres = await readDeck(preview);
      const heading = getSlides(pres)
        .flatMap((slide) => getSlideShapes(slide))
        .find((shape) => getShapeText(shape) === 'HeadingZ');
      assert.ok(heading);
      const runs = getShapeParagraphElements(heading, 0);
      assert.equal(runs.at(-1)?.text, 'Z');
      assert.equal(runs.at(-1)?.format?.bold, true);
      assert.equal(runs.at(-1)?.format?.italic, true);
      assert.ok(runs.slice(0, -1).every((run) => run.format?.bold !== true));
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'outline Ribbon increases selected spans and preserves the caret size',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-outline-font-size-'));
    let preview;
    let browser;
    let editor;
    let outline;
    let page;
    try {
      ({ preview, browser, page, outline, editor } = await openOutline(dir));
      const readRuns = async () => {
        const pres = await readDeck(preview);
        return getSlides(pres).flatMap((slide) =>
          getSlideShapes(slide).map((shape) => ({
            text: getShapeText(shape),
            runs: getShapeParagraphElements(shape, 0),
          })),
        );
      };
      const title = outline.getByRole('textbox').nth(0);
      await title.focus();
      await title.evaluate((input) => window.selectEditorText(input, 2, input.textContent.length));
      await title.press('Shift+ArrowDown');
      await page.keyboard.press('Shift+ArrowDown');
      await page.keyboard.press('Shift+ArrowDown');
      await page.keyboard.press('Shift+ArrowDown');
      const before = await readRuns();
      const increase = editor
        .locator('.ribbon .font-ribbon')
        .getByRole('button', { name: 'Increase Font Size', exact: true });
      let revision = (await waitForState(preview.url, () => true)).revision;
      await increase.click();
      await waitForState(preview.url, (state) => state.revision !== revision);
      revision = (await waitForState(preview.url, () => true)).revision;
      const after = await readRuns();
      const beforeHeading = before.find(({ text }) => text === 'Heading');
      const afterHeading = after.find(({ text }) => text === 'Heading');
      const beforeFollowing = before.find(({ text }) => text === 'Following');
      const afterFollowing = after.find(({ text }) => text === 'Following');
      assert.deepEqual(afterFollowing.runs, beforeFollowing.runs);
      assert.equal(afterHeading.runs[0].text, 'He');
      assert.equal(afterHeading.runs[1].text, 'ading');
      assert.equal(afterHeading.runs[0].format?.size, beforeHeading.runs[0].format?.size);
      assert.equal(afterHeading.runs[1].format?.size, 48);
      assert.equal(after.find((entry) => entry.text === 'Body').runs[0].format?.size, 36);
      assert.equal(after.find((entry) => entry.text === 'Next').runs[0].format?.size, 48);

      await page.keyboard.press('Control+z');
      await waitForState(preview.url, (state) => state.revision !== revision);
      assert.deepEqual(await readRuns(), before);

      await title.focus();
      await title.evaluate((input) =>
        window.selectEditorText(input, input.textContent.length, input.textContent.length),
      );
      revision = (await waitForState(preview.url, () => true)).revision;
      await increase.click();
      await waitForState(preview.url, (state) => state.revision !== revision);
      revision = (await waitForState(preview.url, () => true)).revision;
      await increase.click();
      await waitForState(preview.url, (state) => state.revision !== revision);
      revision = (await waitForState(preview.url, () => true)).revision;
      await title.press('Z');
      await waitForState(preview.url, (state) => state.revision !== revision);
      revision = (await waitForState(preview.url, () => true)).revision;
      let shapes = await readRuns();
      const typed = shapes.find(({ text }) => text === 'HeadingZ');
      assert.ok(typed);
      assert.equal(typed.runs.at(-1)?.format?.size, 54);
      await page.keyboard.press('Control+z');
      await waitForState(preview.url, (state) => state.revision !== revision);
      shapes = await readRuns();
      assert.equal(
        shapes.find(({ text }) => text === 'HeadingZ'),
        undefined,
      );
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

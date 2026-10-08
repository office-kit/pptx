import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeImageBytes,
  getShapeKind,
  getShapeParagraphElements,
  getShapeText,
  getSlides,
  getSlideShapes,
  getSlideSize,
  getSlideText,
  isTableShape,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

// Copying slides and objects writes one system clipboard item: the private
// payload .pptx other editor instances paste, plus HTML, SVG, PNG and plain
// text renditions for other presentation apps.

const DECK_TYPE = 'web application/x-pptx-editor-clipboard+zip';
const PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const SOURCE = `import { Presentation, Slide, Image, Shape, Table, Text } from '@office-kit/pptx-dsl';
const png = Uint8Array.from(atob('${PNG}'), (c) => c.charCodeAt(0));
export default (
  <Presentation>
    <Slide>
      <Text x={0.5} y={0.3} width={5} height={1}>Title 日本語</Text>
      <Shape preset="rect" x={6} y={0.3} width={2} height={1} />
      <Image data={png} x={0.5} y={2} width={2} height={2} />
      <Table x={4} y={2} width={4} height={2} rows={[["A","B"],["C","D"]]} cellStyle={{ fill: '#FFFF00' }} />
    </Slide>
    <Slide><Text x={1} y={1} width={6} height={1}>Second</Text></Slide>
    <Slide><Text x={1} y={1} width={6} height={1}>Third</Text></Slide>
  </Presentation>
);
`;
const TARGET = `import { Presentation, Slide, Text } from '@office-kit/pptx-dsl';
export default <Presentation><Slide><Text x={1} y={1} width={6} height={1}>Target</Text></Slide></Presentation>;
`;

async function deckOf(preview) {
  return loadPresentation(
    new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
  );
}

/** The data: PNGs in clipboard HTML, decoded. */
function htmlImages(html) {
  return [...html.matchAll(/<img[^>]+src="data:image\/png;base64,([^"]+)"/g)].map(([, data]) =>
    Buffer.from(data, 'base64'),
  );
}

/** The pHYs density of a PNG in dots per inch, or null without one. */
function pngDensityDpi(png) {
  const at = png.indexOf('pHYs');
  return at < 0 ? null : Math.round(png.readUInt32BE(at + 4) * 0.0254);
}

/** Reads the system clipboard in the editor frame, waiting for a copy newer than `previous`. */
async function readClipboard(frame, previous) {
  const deadline = Date.now() + 15000;
  while (Date.now() < deadline) {
    const result = await frame
      .evaluate(async (deckType) => {
        const [item] = await navigator.clipboard.read();
        if (!item) return null;
        const out = { types: [...item.types] };
        for (const type of item.types) {
          const blob = await item.getType(type);
          if (type.startsWith('text/') || type === 'image/svg+xml') out[type] = await blob.text();
          else if (type === 'image/png') {
            const bitmap = await createImageBitmap(blob);
            out[type] = { width: bitmap.width, height: bitmap.height };
          } else if (type === deckType) {
            const bytes = new Uint8Array(await blob.arrayBuffer());
            let binary = '';
            for (const byte of bytes) binary += String.fromCharCode(byte);
            out[type] = btoa(binary);
          }
        }
        return out;
      }, DECK_TYPE)
      .catch((error) => {
        // Reading while the editor's write is still pending fails with "Clipboard
        // data has changed", from read() or from a later getType(); poll again.
        if (/Clipboard data has changed/.test(error.message)) return null;
        throw error;
      });
    const marker = /data-pptx-editor-clipboard="([^"]+)"/.exec(result?.['text/html'] ?? '')?.[1];
    if (marker && marker !== previous && result.types.includes(DECK_TYPE))
      return {
        ...result,
        marker,
        deck: await loadPresentation(Buffer.from(result[DECK_TYPE], 'base64')),
      };
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
  throw new Error('The copy did not reach the system clipboard');
}

test(
  'slides and objects reach the system clipboard and paste across editor instances',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-system-clipboard-'));
    let source, target, browser;
    const shot = (page, name) =>
      page.screenshot({ path: `/tmp/pptx-system-clipboard-${name}.png`, fullPage: true });
    let pageA, pageB;
    try {
      await writeFile(join(dir, 'source.tsx'), SOURCE);
      await writeFile(join(dir, 'target.tsx'), TARGET);
      source = await startPreview(join(dir, 'source.tsx'));
      target = await startPreview(join(dir, 'target.tsx'));
      browser = await chromium.launch({ headless: true });
      const context = await browser.newContext({ viewport: { width: 1500, height: 1000 } });
      await context.grantPermissions(['clipboard-read', 'clipboard-write']);
      const errors = [];
      pageA = await context.newPage();
      pageA.on('pageerror', (e) => errors.push(e.message));
      await pageA.goto(source.url);
      const a = pageA.frameLocator('#editor-frame');
      const frameA = await (await pageA.$('#editor-frame')).contentFrame();
      const savedA = () => a.getByText('Saved to this project', { exact: true }).waitFor();
      await savedA();
      const slideSize = getSlideSize(await deckOf(source));
      const slidePx = { width: slideSize.width / 9525, height: slideSize.height / 9525 };

      // One slide, from the keyboard.
      await a.locator('.thumb-row').first().click();
      await a.locator('.thumb-row').first().press('Control+c');
      let clip = await readClipboard(frameA);
      assert.ok(clip.types.includes('text/html'));
      assert.ok(clip.types.includes('text/plain'));
      assert.ok(clip.types.includes('image/png'));
      assert.ok(clip.types.includes('image/svg+xml'));
      assert.match(clip.marker, /^slides:/);
      assert.equal(getSlides(clip.deck).length, 1);
      assert.equal(getSlideText(getSlides(clip.deck)[0]).includes('Title 日本語'), true);
      // Desktop presentation apps paste HTML ahead of pictures, so a slide's HTML is a
      // picture of it. Its pHYs density makes those apps size it as on the slide.
      assert.equal(htmlImages(clip['text/html']).length, 1);
      assert.doesNotMatch(clip['text/html'], /<table[\s>]/);
      assert.equal(pngDensityDpi(htmlImages(clip['text/html'])[0]), 192);
      assert.match(clip['text/plain'], /Title 日本語/);
      assert.match(clip['text/plain'], /A\tB/);
      assert.match(clip['image/svg+xml'], /<text[\s>]/);
      assert.doesNotMatch(clip['image/svg+xml'], /foreignObject/);
      assert.equal(clip['image/png'].width, Math.round(slidePx.width * 2));
      assert.equal(clip['image/png'].height, Math.round(slidePx.height * 2));

      // Several slides, from the thumbnail menu; the HTML has a picture of each.
      await a
        .locator('.thumb-row')
        .nth(2)
        .click({ modifiers: ['Shift'] });
      await a.locator('.thumb-row').nth(2).click({ button: 'right' });
      await a.getByRole('menuitem', { name: 'Copy', exact: true }).click();
      clip = await readClipboard(frameA, clip.marker);
      assert.equal(getSlides(clip.deck).length, 3);
      assert.deepEqual(getSlides(clip.deck).slice(1).map(getSlideText), ['Second', 'Third']);
      assert.equal(htmlImages(clip['text/html']).length, 3);
      assert.match(clip['text/plain'], /Second\n\nThird/);
      // The drawing is the first slide.
      assert.equal(clip['image/png'].width, Math.round(slidePx.width * 2));

      // Download Selected Slides... saves exactly the selected slides.
      await a.locator('.thumb-row').nth(1).click();
      await a
        .locator('.thumb-row')
        .nth(2)
        .click({ modifiers: ['Shift'] });
      await a.locator('.thumb-row').nth(2).click({ button: 'right' });
      const [download] = await Promise.all([
        pageA.waitForEvent('download'),
        a.getByRole('menuitem', { name: 'Download Selected Slides...', exact: true }).click(),
      ]);
      assert.match(download.suggestedFilename(), /-slides-2-3\.pptx$/);
      const downloaded = await loadPresentation(await readFile(await download.path()));
      assert.deepEqual(getSlides(downloaded).map(getSlideText), ['Second', 'Third']);

      // One object: the picture.
      await a.locator('.thumb-row').first().click();
      const hits = a.locator('.hit');
      await hits.nth(2).click();
      await hits.nth(2).press('Control+c');
      clip = await readClipboard(frameA, clip.marker);
      assert.match(clip.marker, /^shapes:/);
      let shapes = getSlideShapes(getSlides(clip.deck)[0]);
      assert.equal(shapes.length, 1);
      assert.equal(getShapeKind(shapes[0]), 'picture');
      assert.ok(getShapeImageBytes(shapes[0]));
      assert.match(clip['text/html'], /<img[^>]+src="data:image\/png;base64,/);
      // The drawing is cropped to the object: 2 in plus the margin, at 2x.
      assert.equal(clip['image/png'].width, Math.round((192 + 8) * 2));
      assert.match(clip['image/svg+xml'], /viewBox="44 188 200 200"/);
      // Objects are drawn without the slide behind them.
      assert.doesNotMatch(
        clip['image/svg+xml'],
        new RegExp(
          `<rect width="${slidePx.width.toFixed(2)}" height="${slidePx.height.toFixed(2)}"`,
        ),
      );

      // The table on its own: an HTML table with its cell fill.
      await hits.nth(3).click();
      await hits.nth(3).press('Control+c');
      clip = await readClipboard(frameA, clip.marker);
      shapes = getSlideShapes(getSlides(clip.deck)[0]);
      assert.equal(shapes.filter(isTableShape).length, 1);
      assert.match(
        clip['text/html'],
        /<table[\s\S]*background-color: rgb\(255, 255, 0\)[\s\S]*>A</,
      );
      assert.equal(clip['text/plain'], 'A\tB\nC\tD');

      // Text on its own stays styled text.
      await hits.nth(0).click();
      await hits.nth(0).press('Control+c');
      clip = await readClipboard(frameA, clip.marker);
      assert.match(clip['text/html'], /Title 日本語/);
      assert.equal(htmlImages(clip['text/html']).length, 0);

      // Several objects, cut: the payload keeps them all, and with a drawing among
      // them the HTML is one picture of the whole selection.
      await hits.nth(0).click();
      await hits.nth(1).click({ modifiers: ['Shift'] });
      await hits.nth(2).click({ modifiers: ['Shift'] });
      await hits.nth(2).press('Control+x');
      clip = await readClipboard(frameA, clip.marker);
      shapes = getSlideShapes(getSlides(clip.deck)[0]);
      assert.equal(shapes.length, 3);
      assert.equal(htmlImages(clip['text/html']).length, 1);
      await savedA();
      assert.equal(getSlideShapes(getSlides(await deckOf(source))[0]).length, 1);

      // Another editor (another tab and project) pastes the objects with their picture.
      pageB = await context.newPage();
      pageB.on('pageerror', (e) => errors.push(e.message));
      await pageB.goto(target.url);
      const b = pageB.frameLocator('#editor-frame');
      const savedB = () => b.getByText('Saved to this project', { exact: true }).waitFor();
      await savedB();
      await b.locator('.thumb-row').first().click();
      await b.locator('.thumb-row').first().press('ControlOrMeta+v');
      await b.locator('.hit').nth(3).waitFor();
      await savedB();
      let result = getSlideShapes(getSlides(await deckOf(target))[0]);
      assert.equal(result.length, 4);
      assert.ok(getShapeImageBytes(result.find((shape) => getShapeKind(shape) === 'picture')));

      // ... and slides, through the context menu's Paste.
      await a.locator('.thumb-row').nth(1).click();
      await a
        .locator('.thumb-row')
        .nth(2)
        .click({ modifiers: ['Shift'] });
      await a.locator('.thumb-row').nth(2).press('Control+c');
      clip = await readClipboard(frameA, clip.marker);
      await pageB.bringToFront();
      await b.locator('.thumb-row').first().click({ button: 'right' });
      await b.getByRole('menuitem', { name: 'Paste', exact: true }).click();
      await b.locator('.thumb-row').nth(2).waitFor();
      await savedB();
      const pasted = getSlides(await deckOf(target));
      assert.equal(pasted.length, 3);
      assert.deepEqual(pasted.slice(1).map(getSlideText), ['Second', 'Third']);

      // A picture copied in another app pastes as a picture.
      const frameB = await (await pageB.$('#editor-frame')).contentFrame();
      await frameB.evaluate(async (png) => {
        const bytes = Uint8Array.from(atob(png), (c) => c.charCodeAt(0));
        await navigator.clipboard.write([
          new ClipboardItem({ 'image/png': new Blob([bytes], { type: 'image/png' }) }),
        ]);
      }, PNG);
      await b.locator('.thumb-row').first().click();
      await b.locator('.thumb-row').first().press('ControlOrMeta+v');
      await b.locator('.hit').nth(4).waitFor();
      await savedB();
      result = getSlideShapes(getSlides(await deckOf(target))[0]);
      assert.equal(result.filter((shape) => getShapeKind(shape) === 'picture').length, 2);

      // Styled text from another app pastes as a text box.
      await frameB.evaluate(async () => {
        await navigator.clipboard.write([
          new ClipboardItem({
            'text/html': new Blob(['<p><b>Bold</b> outside</p>'], { type: 'text/html' }),
            'text/plain': new Blob(['Bold outside'], { type: 'text/plain' }),
          }),
        ]);
      });
      await b.locator('.thumb-row').first().click({ button: 'right' });
      await b.getByRole('menuitem', { name: 'Paste', exact: true }).click();
      await b.locator('.hit').nth(5).waitFor();
      await savedB();
      result = getSlideShapes(getSlides(await deckOf(target))[0]);
      const box = result.find((shape) => getShapeText(shape) === 'Bold outside');
      assert.equal(getShapeParagraphElements(box)[0][0].format?.bold, true);
      assert.deepEqual(errors, []);
    } catch (error) {
      if (pageA) await shot(pageA, 'a-failure');
      if (pageB) await shot(pageB, 'b-failure');
      throw error;
    } finally {
      await browser?.close();
      await source?.close();
      await target?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'without clipboard access, objects still copy and paste within the editor',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-system-clipboard-denied-'));
    let preview, browser, page;
    try {
      await writeFile(join(dir, 'deck.tsx'), SOURCE);
      preview = await startPreview(join(dir, 'deck.tsx'));
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (e) => errors.push(e.message));
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      await saved();
      const count = async () => getSlideShapes(getSlides(await deckOf(preview))[0]).length;
      const hits = editor.locator('.hit');
      await hits.nth(1).click();
      await hits.nth(1).press('Control+c');
      await hits.nth(1).press('Control+v');
      await hits.nth(4).waitFor();
      await saved();
      assert.equal(await count(), 5);
      await editor.locator('.stage').click({ button: 'right', position: { x: 8, y: 8 } });
      await editor.getByRole('menuitem', { name: 'Paste', exact: true }).click();
      await hits.nth(5).waitFor();
      await saved();
      assert.equal(await count(), 6);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

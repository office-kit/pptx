import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  getShapeBounds,
  getShapeRunFormatEffective,
  getShapeText,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setPresentationTheme,
  setShapeTextFormat,
} from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';

test(
  'inline editing keeps character effects, theme colors, bounds, save/load and undo',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-inline-effects-entry-'));
    let preview;
    let browser;
    try {
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      const shape = addSlideTextBox(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(5),
        h: inches(1.2),
        text: 'WordArt editing',
      });
      setPresentationTheme(pres, {
        accent1: '#112233',
        accent2: '#223344',
        accent3: '#334455',
        accent4: '#445566',
      });
      setShapeTextFormat(shape, {
        size: 40,
        color: 'accent1',
        outline: { color: 'accent2', widthEmu: 19050 },
        shadow: { color: 'accent3', offsetEmu: 38100, blurEmu: 50800, angleDeg: 45 },
        glow: { color: 'accent4', radiusEmu: 63500 },
      });
      const source = join(dir, 'source.pptx');
      const file = join(dir, 'deck.tsx');
      await writeFile(source, await savePresentation(pres));
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.getByTitle('Zoom...', { exact: true }).click();
      const zoomDialog = editor.getByRole('dialog', { name: 'Zoom', exact: true });
      await zoomDialog.locator('input[type="number"]').fill('150');
      await zoomDialog.getByRole('button', { name: 'OK', exact: true }).click();
      const hit = editor.locator('.hit').first();
      const before = await hit.boundingBox();
      assert.ok(before);
      const rangeRect = async (locator) =>
        locator.evaluate((node) => {
          const walker = node.ownerDocument.createTreeWalker(node, NodeFilter.SHOW_TEXT);
          const text = walker.nextNode();
          if (!text) throw new Error('text node missing');
          const range = node.ownerDocument.createRange();
          range.selectNodeContents(text);
          return range.getBoundingClientRect().toJSON();
        });
      const staticText = editor
        .locator('.paint foreignObject > div')
        .filter({ hasText: 'WordArt editing' })
        .first();
      const staticRange = await rangeRect(staticText);
      await page.screenshot({ path: '/tmp/inline-effects-before.png' });
      await hit.dblclick();
      const input = editor.locator('.canvas-shell .inline-edit').first();
      await input.waitFor();
      const effects = await input.evaluate((node) => {
        const span = node.querySelector('span');
        if (!span) throw new Error('inline editor has no formatted span');
        const css = getComputedStyle(span);
        return {
          stroke: css.getPropertyValue('-webkit-text-stroke'),
          shadow: css.textShadow,
          color: css.color,
          zoom: getComputedStyle(node).getPropertyValue('--text-zoom'),
          box: node.getBoundingClientRect().toJSON(),
        };
      });
      const inlineRange = await rangeRect(input.locator('span').first());
      await page.screenshot({ path: '/tmp/inline-effects-active.png' });
      const rangeDiff = Object.fromEntries(
        ['x', 'y', 'width', 'height'].map((key) => [key, inlineRange[key] - staticRange[key]]),
      );
      console.log(
        JSON.stringify({
          staticRange,
          inlineRange,
          rangeDiff,
          screenshots: ['/tmp/inline-effects-before.png', '/tmp/inline-effects-active.png'],
        }),
      );
      const zoom = Number(effects.zoom);
      assert.equal(await editor.getByTitle('Zoom...', { exact: true }).innerText(), '150%');
      assert.ok(zoom > 0, `invalid editor text zoom: ${effects.zoom}`);
      assert.ok(Math.abs(parseFloat(effects.stroke) - 2 * zoom) < 0.1, effects.stroke);
      assert.match(effects.stroke, /rgb\(34, 51, 68\)|#223344/i);
      assert.match(effects.shadow, /rgb\(51, 68, 85\)/i);
      assert.match(effects.shadow, /rgb\(68, 85, 102\)/i);
      assert.match(effects.color, /rgb\(17, 34, 51\)/i);
      for (const key of ['x', 'y', 'width', 'height'])
        assert.ok(
          Math.abs(inlineRange[key] - staticRange[key]) < 2,
          `${key} shifted in text range`,
        );
      const afterEntry = await hit.boundingBox();
      assert.ok(afterEntry);
      for (const key of ['x', 'y', 'width', 'height'])
        assert.ok(Math.abs(afterEntry[key] - before[key]) < 0.5, `${key} shifted on edit entry`);

      // Exercise the actual RichTextInput -> SlideCanvas copy handler. The
      // browser clipboard path must receive resolved theme colors rather than
      // leaking scheme tokens into CSS.
      await input.focus();
      await input.evaluate((node) => {
        const selection = node.ownerDocument.getSelection();
        const range = node.ownerDocument.createRange();
        range.selectNodeContents(node);
        selection?.removeAllRanges();
        selection?.addRange(range);
      });
      const clipboard = await input.evaluate((node) => {
        const clipboardData = new DataTransfer();
        const event = new ClipboardEvent('copy', {
          bubbles: true,
          cancelable: true,
          clipboardData,
        });
        node.dispatchEvent(event);
        return {
          html: clipboardData.getData('text/html'),
          plain: clipboardData.getData('text/plain'),
          defaultPrevented: event.defaultPrevented,
        };
      });
      assert.equal(clipboard.defaultPrevented, true);
      assert.equal(clipboard.plain, 'WordArt editing');
      const cssColor = (hex, rgb) =>
        new RegExp(`(?:${hex}|rgb\\(${rgb.replaceAll(', ', ',\\s*')}\\))`, 'i');
      assert.match(clipboard.html, cssColor('#112233', '17, 34, 51'));
      assert.match(clipboard.html, cssColor('#223344', '34, 51, 68'));
      assert.match(clipboard.html, cssColor('#334455', '51, 68, 85'));
      assert.match(clipboard.html, cssColor('#445566', '68, 85, 102'));
      assert.doesNotMatch(clipboard.html, /accent[1-4]/i);

      const beforeEditRevision = (await waitForState(preview.url, () => true)).revision;
      await input.fill('WordArt edited');
      await page.keyboard.press('ControlOrMeta+Enter');
      await waitForState(preview.url, (state) => state.revision > beforeEditRevision);
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const saved = await loadPresentation(
        new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
      );
      const savedShape = getSlideShapes(getSlides(saved)[0])[0];
      assert.equal(getShapeText(savedShape), 'WordArt edited');
      const savedFormat = getShapeRunFormatEffective(saved, savedShape, 0, 0);
      assert.equal(savedFormat.outline?.color, '#223344');
      assert.equal(savedFormat.shadow?.color, '#334455');
      assert.equal(savedFormat.glow?.color, '#445566');

      const beforeUndoRevision = (await waitForState(preview.url, () => true)).revision;
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await waitForState(preview.url, (state) => state.revision > beforeUndoRevision);
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const undone = await loadPresentation(
        new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
      );
      const undoneShape = getSlideShapes(getSlides(undone)[0])[0];
      assert.equal(getShapeRunFormatEffective(undone, undoneShape, 0, 0).color, '#112233');
      assert.equal(getShapeText(undoneShape), 'WordArt editing');
      assert.equal(getShapeRunFormatEffective(undone, undoneShape, 0, 0).outline?.color, '#223344');
      assert.deepEqual(getShapeBounds(undoneShape), getShapeBounds(shape));
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

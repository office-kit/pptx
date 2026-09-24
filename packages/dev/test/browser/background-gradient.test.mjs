import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { chromium } from 'playwright';
import {
  getSlideBackgroundGradientFill,
  getSlideBackground,
  getSlideLayout,
  getSlideLayoutBackgroundGradientFill,
  getSlides,
  loadPresentation,
} from '../../../../dist/index.js';
import { startPreview } from '../helpers/server.mjs';

test(
  'background gradients edit selected slides, undo and persist after reload',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-background-gradient-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation>{['A','B','C'].map(text => <Slide><Text x={1} y={1} width={7} height={1}>{text}</Text></Slide>)}</Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1100 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const gradients = async () =>
        getSlides(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
        ).map(getSlideBackgroundGradientFill);
      const thumbs = editor.locator('.thumb-row');
      await saved();
      await thumbs.nth(0).click();
      await thumbs.nth(1).click({ modifiers: ['Shift'] });
      await editor.getByRole('tab', { name: 'Design', exact: true }).click();
      await editor
        .getByRole('tabpanel', { name: 'Design', exact: true })
        .getByRole('button', { name: 'Format Background', exact: true })
        .click();
      const pane = editor.getByRole('region', { name: 'Format Background', exact: true });
      const backgroundValues = async () =>
        getSlides(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
        ).map(getSlideBackground);
      const transparency = pane.getByRole('spinbutton', {
        name: 'Background transparency',
        exact: true,
      });
      await transparency.fill('40');
      await transparency.press('Tab');
      await saved();
      let backgrounds = await backgroundValues();
      assert.equal(backgrounds[0].opacity, 0.6);
      assert.deepEqual(backgrounds[0], backgrounds[1]);
      assert.equal(backgrounds[2].kind, 'inherit');
      await page.reload();
      await saved();
      await editor.getByRole('tab', { name: 'Design', exact: true }).click();
      await editor
        .getByRole('tabpanel', { name: 'Design', exact: true })
        .getByRole('button', { name: 'Format Background', exact: true })
        .click();

      assert.deepEqual(await backgroundValues(), backgrounds);
      await thumbs.nth(0).click();
      await thumbs.nth(1).click({ modifiers: ['Shift'] });
      await transparency.fill('100');
      await transparency.press('Tab');
      await saved();
      assert.equal((await backgroundValues())[0].opacity, 0);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await backgroundValues(), backgrounds);
      await pane.getByRole('radio', { name: 'Gradient fill', exact: true }).check();
      await saved();
      let values = await gradients();
      assert.equal(values[0].angleDeg, 0);
      assert.deepEqual(values[0], values[1]);
      assert.equal(values[2], null);
      assert.equal(
        await pane.getByRole('checkbox', { name: 'Rotate with shape', exact: true }).isEnabled(),
        false,
      );
      assert.equal(
        await pane
          .getByRole('checkbox', { name: 'Rotate with shape', exact: true })
          .evaluate((node) => node.indeterminate),
        true,
      );
      const brightness = pane.getByRole('spinbutton', {
        name: 'Gradient stop brightness',
        exact: true,
      });
      await brightness.fill('25');
      await brightness.press('Tab');
      await saved();
      values = await gradients();
      assert.equal(values[0].stops[0].brightness, 0.25);
      assert.deepEqual(values[0], values[1]);
      const opacity = pane.getByRole('spinbutton', {
        name: 'Gradient stop transparency',
        exact: true,
      });
      await opacity.fill('40');
      await opacity.press('Tab');
      await saved();
      assert.equal((await gradients())[0].stops[0].opacity, 0.6);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await gradients(), values);
      await pane.getByRole('button', { name: 'Add gradient stop', exact: true }).click();
      await saved();
      assert.equal((await gradients())[0].stops.length, 3);
      await pane.getByLabel('Gradient type', { exact: true }).selectOption('rect');
      await saved();
      values = await gradients();
      assert.equal(values[0].path, 'rect');
      assert.deepEqual(values[0], values[1]);
      assert.equal(values[2], null);
      await page.screenshot({ path: '/tmp/pptx-background-gradient-panel.png' });
      await page.reload();
      await saved();
      await editor.getByRole('tab', { name: 'Design', exact: true }).click();
      await editor
        .getByRole('tabpanel', { name: 'Design', exact: true })
        .getByRole('button', { name: 'Format Background', exact: true })
        .click();

      assert.deepEqual(await gradients(), values);
      await thumbs.nth(0).click();
      await pane.getByRole('button', { name: 'Reset background', exact: true }).click();
      await saved();
      const reset = await gradients();
      assert.equal(reset[0], null);
      assert.deepEqual(reset[1], values[1]);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await gradients(), values);
      await editor.locator('.lang select').selectOption('ja');
      assert.equal(
        await editor
          .getByRole('radio', { name: '塗りつぶし（グラデーション）', exact: true })
          .isChecked(),
        true,
      );
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'editing an inherited gradient creates a slide override and reset restores inheritance',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-background-inherit-'));
    let preview, browser;
    try {
      const zip = unzipSync(
        await readFile(
          new URL('../../../../test/fixtures/minimal/two-slides.pptx', import.meta.url),
        ),
      );
      for (const [name, bytes] of Object.entries(zip)) {
        if (!/^ppt\/(slides|slideLayouts)\/[^/]+\.xml$/.test(name)) continue;
        let xml = strFromU8(bytes).replace(/<p:bg\b[^>]*>[\s\S]*?<\/p:bg>/g, '');
        if (name.startsWith('ppt/slideLayouts/'))
          xml = xml.replace(
            /(<p:cSld\b[^>]*>)/,
            '$1<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="123456"><a:tint val="80000"/><a:satMod val="300000"/></a:srgbClr></a:gs><a:gs pos="100000"><a:srgbClr val="ABCDEF"/></a:gs></a:gsLst><a:lin ang="2700000"/></a:gradFill></p:bgPr></p:bg>',
          );
        zip[name] = strToU8(xml);
      }
      const source = join(dir, 'source.pptx');
      await writeFile(source, zipSync(zip));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1100 } });
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const read = async () =>
        getSlides(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
        );
      await saved();
      await editor.locator('.thumb-row').nth(0).click();
      await editor.getByRole('tab', { name: 'Design', exact: true }).click();
      await editor
        .getByRole('tabpanel', { name: 'Design', exact: true })
        .getByRole('button', { name: 'Format Background', exact: true })
        .click();
      const pane = editor.getByRole('region', { name: 'Format Background', exact: true });
      assert.equal(
        await pane.getByRole('radio', { name: 'Gradient fill', exact: true }).isChecked(),
        true,
      );
      const reset = pane.getByRole('button', { name: 'Reset background', exact: true });
      assert.equal(await reset.isEnabled(), false);
      const initial = getSlideLayoutBackgroundGradientFill(getSlideLayout((await read())[0]));
      const position = pane.getByRole('spinbutton', {
        name: 'Gradient stop position',
        exact: true,
      });
      await position.fill('10');
      await position.press('Tab');
      await saved();
      const moved = getSlideBackgroundGradientFill((await read())[0]);
      assert.equal(moved.stops[0].offset, 0.1);
      assert.deepEqual(moved.stops[0].colorTransforms, initial.stops[0].colorTransforms);
      assert.equal(moved.stops[0].resolvedColor, initial.stops[0].resolvedColor);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(getSlideBackgroundGradientFill((await read())[0]), null);

      const brightness = pane.getByRole('spinbutton', {
        name: 'Gradient stop brightness',
        exact: true,
      });
      await brightness.fill('25');
      await brightness.press('Tab');
      await saved();
      let slides = await read();
      assert.equal(getSlideBackgroundGradientFill(slides[0]).stops[0].brightness, 0.25);
      assert.equal(getSlideBackgroundGradientFill(slides[0]).angleDeg, 45);
      assert.deepEqual(
        getSlideBackgroundGradientFill(slides[0]).stops[0].colorTransforms.slice(0, 2),
        initial.stops[0].colorTransforms,
      );
      assert.deepEqual(getSlideLayoutBackgroundGradientFill(getSlideLayout(slides[0])), initial);
      assert.equal(getSlideBackgroundGradientFill(slides[1]), null);
      await reset.click();
      await saved();
      assert.equal(getSlideBackgroundGradientFill((await read())[0]), null);
      assert.equal(await brightness.inputValue(), '0');
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      await editor.getByRole('tab', { name: 'Design', exact: true }).click();
      await editor
        .getByRole('tabpanel', { name: 'Design', exact: true })
        .getByRole('button', { name: 'Format Background', exact: true })
        .click();

      slides = await read();
      assert.equal(getSlideBackgroundGradientFill(slides[0]).stops[0].brightness, 0.25);
      assert.deepEqual(
        getSlideBackgroundGradientFill(slides[0]).stops[0].colorTransforms.slice(0, 2),
        initial.stops[0].colorTransforms,
      );
      assert.deepEqual(getSlideLayoutBackgroundGradientFill(getSlideLayout(slides[0])), initial);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

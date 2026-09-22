import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { chromium } from 'playwright';
import {
  addSlide,
  createPresentation,
  findSlideLayout,
  getShapeBoundsResolved,
  savePresentation,
  getSlides,
  getSlideShapes,
  getShapeBounds,
  getShapeRotation,
  inches,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'geometry inputs preserve untouched precision and reject invalid values in both languages',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-geometry-input-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1.23456} y={1.34567} width={3.45678} height={1.56789}>日本語 English</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      let ja = false;
      const saved = () =>
        editor
          .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
          .waitFor();
      const read = async () => {
        const deck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        const shape = getSlideShapes(getSlides(deck)[0])[0];
        return { bounds: getShapeBounds(shape), rotation: getShapeRotation(shape) };
      };
      await saved();
      await editor.locator('.hit').first().click();
      const original = await read();
      const field = (name) =>
        editor.locator('.bespoke').getByRole('spinbutton', { name, exact: true });
      const change = async (name, value) => {
        await field(name).fill(value);
        await field(name).press('Tab');
      };
      await change('X', '2.125');
      await saved();
      const moved = { ...original, bounds: { ...original.bounds, x: inches(2.125) } };
      assert.deepEqual(await read(), moved);
      for (const [name, value] of [
        ['W', '-1'],
        ['H', ''],
        ['X', '1e20'],
        ['Y', ''],
      ]) {
        const before = await field(name).inputValue();
        await change(name, value);
        assert.equal(await field(name).inputValue(), before);
        assert.deepEqual(await read(), moved);
      }
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), original);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), moved);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await change('Y', '-0.125');
      await saved();
      await change('回転', '-30.5');
      await saved();
      const rotated = await read();
      assert.deepEqual(rotated.bounds, { ...moved.bounds, y: inches(-0.125) });
      assert.equal(rotated.rotation, 329.5);
      await change('回転', '');
      assert.equal(await field('回転').inputValue(), '329.5');
      assert.deepEqual(await read(), rotated);
      const lock = editor.getByRole('checkbox', { name: '縦横比を固定', exact: true });
      await lock.check();
      await change('W', '6');
      await saved();
      const wider = await read();
      assert.equal(wider.bounds.w, inches(6));
      assert.equal(wider.bounds.h, Math.round((rotated.bounds.h * inches(6)) / rotated.bounds.w));
      assert.equal(wider.bounds.x, rotated.bounds.x);
      assert.equal(wider.bounds.y, rotated.bounds.y);
      assert.equal(wider.rotation, rotated.rotation);
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), rotated);
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), wider);
      await editor.locator('.lang select').selectOption('en');
      ja = false;
      assert.equal(
        await editor.getByRole('checkbox', { name: 'Lock aspect ratio', exact: true }).isChecked(),
        true,
      );
      const beforeOverflow = await field('H').inputValue();
      await change('H', String(27273042316900 / inches(1)));
      assert.equal(await field('H').inputValue(), beforeOverflow);
      assert.deepEqual(await read(), wider);
      await change('H', '2');
      await saved();
      const taller = await read();
      assert.equal(taller.bounds.h, inches(2));
      assert.equal(taller.bounds.w, Math.round((wider.bounds.w * inches(2)) / wider.bounds.h));
      await editor.getByRole('checkbox', { name: 'Lock aspect ratio', exact: true }).uncheck();
      await change('W', '5');
      await saved();
      const resized = await read();
      assert.deepEqual(resized, { ...taller, bounds: { ...taller.bounds, w: inches(5) } });
      await page.reload();
      await saved();
      assert.deepEqual(await read(), resized);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'inherited placeholder geometry remains editable in the numeric panel',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-inherited-geometry-'));
    let preview, browser;
    try {
      const pres = createPresentation();
      addSlide(pres, { layout: findSlideLayout(pres, 'Title and Content') });
      const parts = unzipSync(await savePresentation(pres));
      const path = 'ppt/slides/slide1.xml';
      parts[path] = strToU8(
        strFromU8(parts[path]).replace(/<a:xfrm\b[^>]*>[\s\S]*?<\/a:xfrm>/g, ''),
      );
      const source = join(dir, 'source.pptx');
      await writeFile(source, zipSync(parts));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      let ja = false;
      const saved = () =>
        editor
          .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
          .waitFor();
      const read = async () => {
        const deck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        const shape = getSlideShapes(getSlides(deck)[0])[0];
        return { raw: getShapeBounds(shape), resolved: getShapeBoundsResolved(deck, shape) };
      };
      await saved();
      const original = await read();
      assert.equal(original.raw, null);
      assert.ok(original.resolved);
      await editor.locator('.hit').first().click();
      const field = (name) =>
        editor.locator('.bespoke').getByRole('spinbutton', { name, exact: true });
      assert.equal(await field('X').count(), 1);
      await field('X').fill('2.125');
      await field('X').press('Tab');
      await saved();
      const moved = { ...original.resolved, x: inches(2.125) };
      assert.deepEqual((await read()).resolved, moved);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), original);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await editor.getByRole('checkbox', { name: '縦横比を固定', exact: true }).check();
      await field('W').fill('6');
      await field('W').press('Tab');
      await saved();
      const resized = {
        ...original.resolved,
        w: inches(6),
        h: Math.round((original.resolved.h * inches(6)) / original.resolved.w),
      };
      assert.deepEqual((await read()).resolved, resized);
      await page.reload();
      await saved();
      assert.deepEqual((await read()).resolved, resized);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'multiple selected geometry values show mixed values and apply together in both languages',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-mixed-rotation-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={2} height={1} rotation={20}>日本語</Text><Text x={4} y={1} width={2} height={2}>English</Text><Text x={7} y={1} width={2} height={1}>Untouched</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      let ja = false;
      const saved = () =>
        editor
          .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
          .waitFor();
      const read = async () => {
        const deck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getSlideShapes(getSlides(deck)[0]).map((shape) => ({
          rotation: getShapeRotation(shape),
          bounds: getShapeBounds(shape),
        }));
      };
      await saved();
      const before = await read();
      await editor.locator('.hit').nth(0).click();
      const rotation = () =>
        editor
          .locator('.bespoke')
          .getByRole('spinbutton', { name: ja ? '回転' : 'Rotation', exact: true });
      assert.equal(await rotation().inputValue(), '20');
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      const sizeField = (name) =>
        editor.locator('.bespoke').getByRole('spinbutton', { name, exact: true });
      assert.equal(await sizeField('X').inputValue(), '');
      assert.equal(await sizeField('X').getAttribute('placeholder'), 'Mixed');
      await sizeField('X').fill('2.125');
      await sizeField('X').press('Tab');
      await saved();
      assert.deepEqual(
        await read(),
        before.map((value, index) =>
          index < 2 ? { ...value, bounds: { ...value.bounds, x: inches(2.125) } } : value,
        ),
      );
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), before);
      await editor.getByRole('checkbox', { name: 'Lock aspect ratio', exact: true }).check();
      await sizeField('W').fill('3');
      await sizeField('W').press('Tab');
      await saved();
      assert.deepEqual(
        await read(),
        before.map((value, index) =>
          index < 2
            ? {
                ...value,
                bounds: { ...value.bounds, w: inches(3), h: inches(index === 0 ? 1.5 : 3) },
              }
            : value,
        ),
      );
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), before);
      assert.equal(await rotation().inputValue(), '');
      assert.equal(await rotation().getAttribute('placeholder'), 'Mixed');
      await rotation().fill('');
      await rotation().press('Tab');
      assert.deepEqual(await read(), before);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      assert.equal(await sizeField('H').inputValue(), '');
      await sizeField('H').fill('1');
      await sizeField('H').press('Tab');
      await saved();
      assert.deepEqual(
        await read(),
        before.map((value, index) =>
          index < 2
            ? {
                ...value,
                bounds: { ...value.bounds, h: inches(1), w: inches(index === 0 ? 2 : 1) },
              }
            : value,
        ),
      );
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), before);
      await sizeField('H').fill(String(27273042316900 / inches(1)));
      await sizeField('H').press('Tab');
      assert.deepEqual(await read(), before);
      assert.equal(await sizeField('H').inputValue(), '');
      assert.equal(await rotation().inputValue(), '');
      await rotation().fill('-30.5');
      await rotation().press('Tab');
      await saved();
      const expected = before.map((value, index) =>
        index < 2 ? { ...value, rotation: 329.5 } : value,
      );
      assert.deepEqual(await read(), expected);
      assert.equal(await rotation().inputValue(), '329.5');
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), before);
      assert.equal(await rotation().inputValue(), '');
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      await editor.getByRole('checkbox', { name: '縦横比を固定', exact: true }).uncheck();
      await sizeField('W').fill('4');
      await sizeField('W').press('Tab');
      await saved();
      const resized = expected.map((value, index) =>
        index < 2 ? { ...value, bounds: { ...value.bounds, w: inches(4) } } : value,
      );
      assert.deepEqual(await read(), resized);
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), expected);
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      assert.deepEqual(await read(), resized);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

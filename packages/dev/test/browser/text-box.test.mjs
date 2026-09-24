import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  cm,
  getShapeBodyPrEffective,
  getShapeTextAutoFit,
  getShapeTextColumns,
  getSlides,
  getSlideShapes,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'text-box properties preserve other margins, apply to a mixed selection and undo',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-text-box-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={2} height={1}>日本語 English</Text><Text x={4} y={1} width={2} height={1}>Second</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const box = editor.locator('.text-box');
      let ja = false;
      const saved = () =>
        editor
          .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
          .waitFor();
      const read = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getSlideShapes(getSlides(pres)[0]).map((shape) => ({
          ...getShapeBodyPrEffective(pres, shape),
          fit: getShapeTextAutoFit(shape),
          columns: getShapeTextColumns(shape),
        }));
      };
      await saved();
      const initial = await read();
      await editor.locator('.hit').first().click();
      await box.locator('summary').click();
      const left = () =>
        box.getByRole('spinbutton', { name: ja ? '左余白' : 'Left margin', exact: true });
      await left().fill('0.375');
      await left().press('Tab');
      await saved();
      const changed = await read();
      assert.deepEqual(changed[0].margins, { ...initial[0].margins, left: cm(0.375) });
      assert.deepEqual(changed[1], initial[1]);
      for (const invalid of ['-1', '55.89', '']) {
        const before = await left().inputValue();
        await left().fill(invalid);
        await left().press('Tab');
        assert.equal(await left().inputValue(), before);
        assert.deepEqual(await read(), changed);
      }
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      assert.equal(await left().inputValue(), '');
      await left().fill('0.5');
      await left().press('Tab');
      await saved();
      assert.deepEqual(
        (await read()).map((item) => item.margins.left),
        [cm(0.5), cm(0.5)],
      );
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), changed);
      await box
        .getByRole('combobox', { name: 'Text direction', exact: true })
        .selectOption('vert270');
      await saved();
      assert.deepEqual(
        (await read()).map((item) => item.vert),
        ['vert270', 'vert270'],
      );
      await box.getByRole('radio', { name: 'Shrink text on overflow', exact: true }).check();
      await saved();
      assert.deepEqual(
        (await read()).map((item) => item.fit),
        ['normal', 'normal'],
      );
      await box.getByRole('checkbox', { name: 'Wrap text in shape', exact: true }).uncheck();
      await saved();
      assert.deepEqual(
        (await read()).map((item) => item.wrap),
        ['none', 'none'],
      );
      const beforeColumns = await read();
      await box.getByRole('button', { name: 'Columns...', exact: true }).click();
      const columns = editor.getByRole('dialog', { name: 'Columns', exact: true });
      await columns.getByRole('spinbutton', { name: 'Number of columns:' }).fill('3');
      await columns.getByRole('spinbutton', { name: 'Spacing between columns:' }).fill('0.25');
      assert.deepEqual(await read(), beforeColumns);
      await columns.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.deepEqual(await read(), beforeColumns);
      await box.getByRole('button', { name: 'Columns...', exact: true }).click();
      await columns.getByRole('spinbutton', { name: 'Number of columns:' }).fill('17');
      await columns.getByRole('button', { name: 'OK', exact: true }).click();
      assert.equal(await columns.isVisible(), true);
      await columns.getByRole('spinbutton', { name: 'Number of columns:' }).fill('3');
      await columns.getByRole('spinbutton', { name: 'Spacing between columns:' }).fill('0.25');
      await columns.getByRole('spinbutton', { name: 'Spacing between columns:' }).press('Enter');
      await saved();
      assert.deepEqual(
        (await read()).map((item) => item.columns),
        [
          { count: 3, gapEmu: cm(0.25) },
          { count: 3, gapEmu: cm(0.25) },
        ],
      );
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), beforeColumns);
      await box.getByRole('button', { name: 'Columns...', exact: true }).click();
      await columns.getByRole('spinbutton', { name: 'Number of columns:' }).press('Escape');
      await columns.waitFor({ state: 'hidden' });
      assert.deepEqual(await read(), beforeColumns);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await box
        .getByRole('combobox', { name: '垂直方向の配置', exact: true })
        .selectOption('bottom');
      await saved();
      assert.deepEqual(
        (await read()).map((item) => item.anchor),
        ['bottom', 'bottom'],
      );
      const top = box.getByRole('spinbutton', { name: '上余白', exact: true });
      await top.fill('0.2');
      await top.press('Tab');
      await saved();
      assert.deepEqual(
        (await read()).map((item) => item.margins.top),
        [cm(0.2), cm(0.2)],
      );
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(
        (await read()).map((item) => item.margins.top),
        initial.map((item) => item.margins.top),
      );
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

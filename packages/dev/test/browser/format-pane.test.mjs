import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { startPreview } from '../helpers/server.mjs';

test(
  'format tabs isolate controls, retain editing state and support keyboard navigation',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-format-pane-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={3} height={1}>Format pane</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      await saved();
      assert.equal(await editor.getByRole('tablist', { name: 'Format Shape' }).count(), 0);
      await editor.locator('.hit').click();
      const tabs = editor.getByRole('tablist', { name: 'Format Shape' });
      const paint = tabs.getByRole('tab', { name: 'Fill & Line' });
      const effects = tabs.getByRole('tab', { name: 'Effects', exact: true });
      const size = tabs.getByRole('tab', { name: 'Size & Properties' });
      assert.equal(await paint.getAttribute('aria-selected'), 'true');
      assert.equal(await editor.getByLabel('Fill', { exact: true }).isVisible(), true);
      assert.equal(await editor.getByRole('spinbutton', { name: 'Width', exact: true }).count(), 0);
      await paint.press('ArrowRight');
      assert.equal(await effects.getAttribute('aria-selected'), 'true');
      assert.equal(await effects.evaluate((el) => el === document.activeElement), true);
      assert.equal(await editor.getByLabel('Fill', { exact: true }).isVisible(), false);
      assert.equal(await editor.locator('.cat-head').count(), 1);
      await effects.press('End');
      assert.equal(await size.getAttribute('aria-selected'), 'true');
      const scale = editor.getByRole('spinbutton', { name: 'Scale Width', exact: true });
      await scale.fill('150');
      await scale.press('Tab');
      await saved();
      const width = editor.getByRole('spinbutton', { name: 'Width', exact: true });
      const changedWidth = await width.inputValue();
      await paint.click();
      await size.click();
      assert.equal(await scale.inputValue(), '150');
      assert.equal(await width.inputValue(), changedWidth);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(await scale.inputValue(), '100');
      await size.press('Home');
      assert.equal(await paint.getAttribute('aria-selected'), 'true');
      await paint.press('ArrowLeft');
      assert.equal(await size.getAttribute('aria-selected'), 'true');
      await editor.locator('.lang select').selectOption('ja');
      assert.equal(
        await editor.getByRole('tab', { name: 'サイズとプロパティ' }).getAttribute('aria-selected'),
        'true',
      );
      await page.screenshot({ path: '/tmp/pptx-pr287-format-tabs.png', fullPage: true });
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

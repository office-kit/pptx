import assert from 'node:assert/strict';
import { copyFile, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { startPreview } from '../helpers/server.mjs';

for (const locale of ['en', 'ja']) {
  test(
    `outline view changes tolerate queued selection events (${locale})`,
    { timeout: 60000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-outline-selection-'));
      let preview, browser;
      try {
        await copyFile(
          new URL('../../../../test/fixtures/minimal/one-text-slide.pptx', import.meta.url),
          join(dir, 'template.pptx'),
        );
        const file = join(dir, 'deck.tsx');
        await writeFile(
          file,
          `import {readFileSync} from 'node:fs'; import {Presentation,Slide,Fill} from '@office-kit/pptx-dsl'; const source = new Uint8Array(readFileSync(new URL('./template.pptx', import.meta.url))); export default <Presentation source={source} mode="compose"><Slide layout={{name:'Title and Content'}}><Fill target={{placeholder:{type:'title'}}}>日本語 title</Fill><Fill target={{placeholder:{idx:1}}}>Body</Fill></Slide></Presentation>;`,
        );
        preview = await startPreview(file);
        browser = await chromium.launch({ headless: true });
        const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
        const errors = [];
        page.on('pageerror', (error) => errors.push(error.message));
        await page.addInitScript(
          (language) => localStorage.setItem('ok-editor-locale', language),
          locale,
        );
        await page.goto(preview.url);
        await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
        const editor = page.frameLocator('#editor-frame');
        const labels =
          locale === 'en'
            ? { view: 'View', outline: 'Outline View' }
            : { view: '表示', outline: 'アウトライン表示' };
        await editor.getByRole('tab', { name: labels.view, exact: true }).click();
        await editor
          .getByRole('tabpanel', { name: labels.view, exact: true })
          .getByRole('button', { name: labels.outline, exact: true })
          .click();
        const outline = editor.getByRole('navigation', { name: labels.outline, exact: true });
        const title = outline.getByRole('textbox').first();
        await title.focus();
        const detached = await title.elementHandle();
        await editor
          .getByRole('tabpanel', { name: labels.view, exact: true })
          .getByRole('button', { name: locale === 'en' ? 'Normal' : '標準', exact: true })
          .click();
        await outline.waitFor({ state: 'detached' });
        await detached.evaluate((node) => node.dispatchEvent(new Event('select')));
        await detached.dispose();
        assert.deepEqual(errors, []);
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );
}

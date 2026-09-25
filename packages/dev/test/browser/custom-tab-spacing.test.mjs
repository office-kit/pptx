import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { startPreview } from '../helpers/server.mjs';

test('custom tab alignment uses painted browser font widths', { timeout: 90000 }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-custom-tabs-'));
  let preview, browser;
  try {
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={5} height={3}>{'A\\t12.34'}</Text></Slide></Presentation>`,
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
    await editor.locator('.hit').first().click();
    for (const alignment of ['Left', 'Center', 'Right', 'Decimal']) {
      await editor
        .locator('.ribbon')
        .getByRole('button', { name: 'Line spacing', exact: true })
        .click();
      await editor.getByRole('menuitem', { name: 'Line Spacing Options...', exact: true }).click();
      const paragraph = editor.getByRole('dialog', { name: 'Paragraph', exact: true });
      await paragraph.getByRole('button', { name: 'Tabs...', exact: true }).click();
      const tabs = editor.getByRole('dialog', { name: 'Tabs', exact: true });
      await tabs.getByLabel('Tab stop position:', { exact: true }).fill('5.08');
      await tabs.getByRole('radio', { name: alignment, exact: true }).check();
      await tabs.getByRole('button', { name: 'Set', exact: true }).click();
      await tabs.getByRole('button', { name: 'OK', exact: true }).click();
      await paragraph.getByRole('button', { name: 'OK', exact: true }).click();
      await saved();
      const placement = await editor
        .locator('.paint [data-pptx-paragraph]')
        .first()
        .evaluate((element) => {
          const spans = [...element.querySelectorAll('tspan')];
          const start = spans.find((span) => span.textContent === 'A');
          const field = spans.find((span) => span.textContent === '12.34');
          if (!start || !field) throw new Error('Missing tab field glyphs');
          return {
            start: field.getStartPositionOfChar(0).x - start.getStartPositionOfChar(0).x,
            width: field.getComputedTextLength(),
            decimal: field.getSubStringLength(0, 2),
          };
        });
      const shift =
        alignment === 'Center'
          ? placement.width / 2
          : alignment === 'Right'
            ? placement.width
            : alignment === 'Decimal'
              ? placement.decimal
              : 0;
      assert.ok(
        Math.abs(placement.start + shift - 192) < 1,
        `${alignment}: ${JSON.stringify(placement)}`,
      );
    }
    assert.deepEqual(errors, []);
  } finally {
    await browser?.close();
    await preview?.close();
    await rm(dir, { recursive: true, force: true });
  }
});

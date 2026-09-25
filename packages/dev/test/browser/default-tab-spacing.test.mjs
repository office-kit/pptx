import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { startPreview } from '../helpers/server.mjs';

test('default tab spacing reaches preview and direct editing', { timeout: 90000 }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-default-tabs-'));
  let preview, browser;
  try {
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={5} height={3}>{'A\\tB'}</Text></Slide></Presentation>`,
    );
    preview = await startPreview(file);
    browser = await chromium.launch({ headless: true });
    const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
    await page.goto(preview.url);
    await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
    const editor = page.frameLocator('#editor-frame');
    await editor.getByText('Saved to this project', { exact: true }).waitFor();
    const painted = editor.locator('.paint [data-pptx-paragraph]').first();
    const offset = (locator) =>
      locator.evaluate((paragraph) => {
        const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
        let text;
        while ((text = walker.nextNode())) {
          const index = text.textContent.indexOf('B');
          if (index < 0) continue;
          const range = document.createRange();
          range.setStart(text, index);
          range.setEnd(text, index + 1);
          const box = paragraph.getBoundingClientRect();
          const zoom = box.width / paragraph.offsetWidth;
          return (range.getBoundingClientRect().left - box.left) / zoom;
        }
        throw new Error('Tab field not found');
      });
    assert.ok(Math.abs((await offset(painted)) - 96) < 1);
    await editor.locator('.hit').first().click();
    await editor
      .locator('.ribbon')
      .getByRole('button', { name: 'Line spacing', exact: true })
      .click();
    await editor.getByRole('menuitem', { name: 'Line Spacing Options...', exact: true }).click();
    const paragraph = editor.getByRole('dialog', { name: 'Paragraph', exact: true });
    await paragraph.getByRole('button', { name: 'Tabs...', exact: true }).click();
    const tabs = editor.getByRole('dialog', { name: 'Tabs', exact: true });
    await tabs.getByLabel('Default tab stops:', { exact: true }).fill('1.27');
    await tabs.getByRole('button', { name: 'OK', exact: true }).click();
    await paragraph.getByRole('button', { name: 'OK', exact: true }).click();
    await editor.getByText('Saved to this project', { exact: true }).waitFor();
    assert.ok(Math.abs((await offset(painted)) - 48) < 1);
    await editor.locator('.hit').first().dblclick();
    const editing = editor.locator('[contenteditable="true"] [data-text-paragraph]').first();
    await editing.waitFor();
    const result = await editing.evaluate((element) => ({
      tab: getComputedStyle(element).tabSize,
      zoom: getComputedStyle(element).getPropertyValue('--text-zoom'),
    }));
    assert.ok(Math.abs(parseFloat(result.tab) / Number(result.zoom) - 48) < 0.1);
    assert.ok(Math.abs((await offset(editing)) / Number(result.zoom) - 48) < 1);
  } finally {
    await browser?.close();
    await preview?.close();
    await rm(dir, { recursive: true, force: true });
  }
});

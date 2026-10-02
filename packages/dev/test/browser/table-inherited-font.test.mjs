import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  createPresentation,
  addBlankSlide,
  addSlideTable,
  inches,
  savePresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';
import { installRichTextSelection } from '../helpers/rich-text.mjs';

test(
  'table paragraph default font survives entering text editing',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-table-inherited-font-'));
    let preview, browser;
    try {
      const pres = createPresentation();
      addSlideTable(addBlankSlide(pres), {
        x: inches(1),
        y: inches(1),
        w: inches(6),
        h: inches(2),
        rows: [['Inherited text']],
      });
      const parts = unzipSync(await savePresentation(pres));
      const name = 'ppt/slides/slide1.xml';
      const xml = strFromU8(parts[name]);
      const defaults =
        '<a:defRPr sz="2800" b="1"><a:solidFill><a:srgbClr val="AA2244"/></a:solidFill><a:latin typeface="Courier New"/></a:defRPr>';
      parts[name] = strToU8(
        xml
          .replace('</a:pPr>', `${defaults}</a:pPr>`)
          .replace(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/, '<a:rPr/>'),
      );
      assert.notEqual(strFromU8(parts[name]), xml, 'fixture has a paragraph default');
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
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.hit').first().dblclick();
      const input = editor.locator('.inline-edit');
      await input.waitFor();
      const style = await input
        .locator('span')
        .filter({ hasText: 'Inherited text' })
        .last()
        .evaluate((node) => {
          const s = getComputedStyle(node);
          return {
            family: s.fontFamily,
            weight: s.fontWeight,
            color: s.color,
            size: node.style.fontSize,
          };
        });
      assert.match(style.family, /Courier New/);
      assert.equal(style.weight, '700');
      assert.equal(style.color, 'rgb(170, 34, 68)');
      assert.equal(style.size, 'calc(28pt * var(--text-zoom))');
      const copied = await input.evaluate((node) => {
        node.focus();
        window.selectEditorText(node, 0, node.textContent.length);
        node.dispatchEvent(new Event('select', { bubbles: true }));
        const data = new DataTransfer();
        node.dispatchEvent(
          new ClipboardEvent('copy', { clipboardData: data, bubbles: true, cancelable: true }),
        );
        return JSON.parse(data.getData('application/x-office-kit-text+json'));
      });
      assert.equal(copied.text, 'Inherited text');
      assert.equal(copied.formats[0].format.font, 'Courier New');
      assert.equal(copied.formats[0].format.size, 28);
      assert.equal(copied.formats[0].format.bold, true);
      assert.equal(copied.formats[0].format.color, '#AA2244');
      await page.keyboard.press('Control+T');
      const dialog = editor.getByRole('dialog', { name: 'Font', exact: true });
      await dialog.waitFor();
      assert.equal(
        await dialog.getByLabel('Latin text font', { exact: true }).inputValue(),
        'Courier New',
      );
      assert.equal(await dialog.getByLabel('Font size', { exact: true }).inputValue(), '28');
      assert.equal(
        await dialog.getByRole('combobox', { name: /^Font style/ }).inputValue(),
        'true:false',
      );
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

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
  addSlideTextBox,
  inches,
  savePresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';
import { installRichTextSelection } from '../helpers/rich-text.mjs';

for (const { kind, field, lineBreak = false } of [
  { kind: 'table', field: false },
  { kind: 'table', field: true },
  { kind: 'shape', field: true },
  { kind: 'shape', field: false, lineBreak: true },
  { kind: 'table', field: false, lineBreak: true },
]) {
  test(
    `${kind} paragraph default font survives entering text editing (field: ${field}, break: ${lineBreak})`,
    { timeout: 60000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-table-inherited-font-'));
      let preview, browser;
      try {
        const pres = createPresentation();
        const slide = addBlankSlide(pres);
        const bounds = { x: inches(1), y: inches(1), w: inches(6), h: inches(2) };
        if (kind === 'table') addSlideTable(slide, { ...bounds, rows: [['Inherited text']] });
        else addSlideTextBox(slide, { ...bounds, text: 'Inherited text' });
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
        if (kind === 'shape')
          parts[name] = strToU8(
            xml.replace(
              /<a:p>[\s\S]*?<\/a:p>/,
              `<a:p><a:pPr>${defaults}</a:pPr><a:r><a:rPr/><a:t>Inherited text</a:t></a:r></a:p>`,
            ),
          );
        if (field)
          parts[name] = strToU8(
            strFromU8(parts[name])
              .replace(
                '<a:r>',
                '<a:fld id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}" type="datetime1">',
              )
              .replace('</a:r>', '</a:fld>'),
          );
        if (lineBreak)
          parts[name] = strToU8(
            strFromU8(parts[name]).replace('<a:r>', '<a:br><a:rPr i="1"/></a:br><a:r>'),
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
        if (lineBreak) {
          const breakStyle = await input
            .locator('span')
            .first()
            .evaluate((node) => ({
              text: node.textContent,
              font: node.style.fontFamily,
              size: node.style.fontSize,
              weight: getComputedStyle(node).fontWeight,
            }));
          assert.equal(breakStyle.text, '\n');
          assert.match(breakStyle.font, /Courier New/);
          assert.equal(breakStyle.size, 'calc(28pt * var(--text-zoom))');
          assert.equal(breakStyle.weight, '700');
        }
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
        if (lineBreak) return;
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
}

test(
  'table style text formatting survives entering and copying text editing',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-table-style-font-'));
    let preview, browser;
    try {
      const pres = createPresentation();
      addSlideTable(addBlankSlide(pres), {
        x: inches(1),
        y: inches(1),
        w: inches(6),
        h: inches(2),
        rows: [['Header style'], ['Body style']],
      });
      const parts = unzipSync(await savePresentation(pres));
      const styleId = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';
      parts['ppt/tableStyles.xml'] = strToU8(
        `<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="${styleId}"><a:tblStyle styleId="${styleId}" styleName="Regression"><a:wholeTbl><a:tcTxStyle b="on"><a:font><a:latin typeface="Arial"/><a:ea typeface="Arial"/><a:cs typeface="Arial"/></a:font><a:srgbClr val="008800"/></a:tcTxStyle></a:wholeTbl><a:firstRow><a:tcTxStyle b="off"><a:font><a:latin typeface="Courier New"/><a:ea typeface="Courier New"/><a:cs typeface="Courier New"/></a:font><a:srgbClr val="CC0000"/></a:tcTxStyle></a:firstRow></a:tblStyle></a:tblStyleLst>`,
      );
      const slideName = 'ppt/slides/slide1.xml';
      parts[slideName] = strToU8(
        strFromU8(parts[slideName])
          .replace('<a:tblPr>', '<a:tblPr firstRow="1">')
          .replace(/(<a:r><a:rPr[^>]*>)<a:solidFill>[\s\S]*?<\/a:solidFill>/g, '$1'),
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
        .filter({ hasText: 'Header style' })
        .last()
        .evaluate((node) => {
          const s = getComputedStyle(node);
          return { family: s.fontFamily, weight: s.fontWeight, color: s.color };
        });
      assert.match(style.family, /Courier New/);
      assert.equal(style.weight, '400');
      assert.equal(style.color, 'rgb(204, 0, 0)');
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
      assert.equal(copied.formats[0].format.font, 'Courier New');
      assert.equal(copied.formats[0].format.bold, false);
      assert.equal(copied.formats[0].format.color, '#CC0000');
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

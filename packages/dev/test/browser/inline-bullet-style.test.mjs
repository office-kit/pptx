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
  inches,
  loadPresentation,
  savePresentation,
  setParagraphBullet,
  setShapeRunFormat,
  getPresentationFonts,
} from '@office-kit/pptx';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { startPreview } from '../helpers/server.mjs';

test(
  'inline editing preserves explicit and inherited bullet color, size and font',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-inline-bullet-style-'));
    let preview;
    let browser;
    try {
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      const shape = addSlideTextBox(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(6),
        h: inches(3),
        text: 'Bullet text\nPoints bullet\nTheme bullet\nInherited bullet\nFollow text bullet',
      });
      setParagraphBullet(shape, 0, 'bullet');
      setParagraphBullet(shape, 1, 'bullet');
      setParagraphBullet(shape, 2, 'bullet');
      setParagraphBullet(shape, 3, 'bullet');
      setParagraphBullet(shape, 4, 'bullet');
      setShapeRunFormat(shape, 4, 0, { font: 'Georgia', color: '#CC5500', size: 24 });
      const parts = unzipSync(await savePresentation(pres));
      const name = 'ppt/slides/slide1.xml';
      const xml = strFromU8(parts[name]).replace(
        '<a:lstStyle/>',
        '<a:lstStyle><a:lvl1pPr><a:buClr><a:srgbClr val="336699"/></a:buClr><a:buSzPct val="125000"/><a:buFont typeface="Verdana"/></a:lvl1pPr></a:lstStyle>',
      );
      let paragraphIndex = 0;
      parts[name] = strToU8(
        xml.replace(/<a:pPr([^>]*)><a:buChar char="•"\/><\/a:pPr>/g, (match, attrs) => {
          if (paragraphIndex++ === 0) {
            return `<a:pPr${attrs}><a:buClr><a:srgbClr val="AA00CC"/></a:buClr><a:buSzPct val="50%"/><a:buFont typeface="Courier New"/><a:buChar char="•"/></a:pPr>`;
          }
          if (paragraphIndex === 2) {
            return `<a:pPr${attrs}><a:buClr><a:srgbClr val="0088FF"/></a:buClr><a:buSzPts val="2000"/><a:buFont typeface="Georgia"/><a:buChar char="•"/></a:pPr>`;
          }
          if (paragraphIndex === 3) {
            return `<a:pPr${attrs}><a:buFont typeface="+mj-lt"/><a:buChar char="•"/></a:pPr>`;
          }
          if (paragraphIndex === 5) {
            return `<a:pPr${attrs}><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buChar char="•"/></a:pPr>`;
          }
          return match;
        }),
      );
      assert.notEqual(strFromU8(parts[name]), xml);
      assert.match(strFromU8(parts[name]), /AA00CC/);
      const imported = await loadPresentation(zipSync(parts));
      const expectedThemeFont = getPresentationFonts(imported)?.majorLatin;
      assert.ok(expectedThemeFont);
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
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.hit').first().dblclick();
      const input = editor.locator('.canvas-shell .inline-edit');
      await input.waitFor();
      const markers = input.locator('[data-list-marker]');
      assert.equal(await markers.count(), 5);
      const styles = await markers.evaluateAll((nodes) =>
        nodes.map((node) => {
          const pseudo = getComputedStyle(node, '::before');
          const text = node.querySelector('span');
          return {
            color: pseudo.color,
            font: pseudo.fontFamily,
            size: pseudo.fontSize,
            markerSize: node.style.getPropertyValue('--marker-size'),
            textSize: text && getComputedStyle(text).fontSize,
          };
        }),
      );
      assert.equal(styles[0].color, 'rgb(170, 0, 204)');
      assert.match(styles[0].font, /Courier New/);
      assert.equal(styles[0].markerSize, 'calc(9pt * var(--text-zoom))');
      assert.ok(
        Math.abs(Number.parseFloat(styles[0].size) - Number.parseFloat(styles[0].textSize) * 0.5) <
          0.01,
      );
      assert.equal(styles[1].color, 'rgb(0, 136, 255)');
      assert.match(styles[1].font, /Georgia/);
      assert.equal(styles[1].markerSize, 'calc(20pt * var(--text-zoom))');
      assert.match(
        styles[2].font,
        new RegExp(expectedThemeFont.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')),
      );
      assert.doesNotMatch(styles[2].font, /\+mj-lt/);
      assert.equal(styles[3].color, 'rgb(51, 102, 153)');
      assert.match(styles[3].font, /Verdana/);
      assert.equal(styles[3].markerSize, 'calc(22.5pt * var(--text-zoom))');
      assert.equal(styles[4].color, 'rgb(204, 85, 0)');
      assert.match(styles[4].font, /Georgia/);
      assert.equal(styles[4].markerSize, 'calc(24pt * var(--text-zoom))');
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

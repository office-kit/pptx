import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';
import { chromium } from 'playwright';
import { startPreview } from '../helpers/server.mjs';

const styles = [
  'none',
  'sng',
  'dbl',
  'heavy',
  'dotted',
  'dottedHeavy',
  'dash',
  'dashHeavy',
  'dashLong',
  'dashLongHeavy',
  'dotDash',
  'dotDashHeavy',
  'dotDotDash',
  'dotDotDashHeavy',
  'wavy',
  'wavyHeavy',
  'wavyDbl',
  'words',
];

function decorations(locator) {
  return locator.evaluate((root) =>
    [...root.querySelectorAll('span, u')].flatMap((node) => {
      const css = getComputedStyle(node);
      if (!css.textDecorationLine.includes('underline') && css.backgroundImage === 'none')
        return [];
      // Preview uses SVG scaling; the editor scales font sizes directly.
      const em = (value) =>
        value.replace(
          /([\d.]+)px/g,
          (_, px) => `${(Number(px) / parseFloat(css.fontSize)).toFixed(4)}em`,
        );
      return [
        {
          text: node.textContent,
          line: css.textDecorationLine.replace('line-through', '').trim(),
          style: css.textDecorationStyle,
          thickness: em(css.textDecorationThickness),
          image: css.backgroundImage,
          size: em(css.backgroundSize),
          position: css.backgroundPosition,
        },
      ];
    }),
  );
}
function geometry(locator) {
  return locator.evaluate((root) => {
    const range = document.createRange();
    range.selectNodeContents(root);
    const rect = range.getBoundingClientRect();
    return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
  });
}

test(
  'all underline styles survive edit entry without shifting text or altering solid strike',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-inline-underlines-'));
    let browser;
    let preview;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation>${styles.map((underline) => `<Slide><Text x={1} y={1} width={7} height={1} paragraphs={[{runs:[{text:'Underline words',format:{font:'Arial',size:36,color:'#154687',underline:'${underline}',strike:true}}]}]}/></Slide>`).join('')}</Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const images = new Map();
      for (const [index, underline] of styles.entries()) {
        await editor.locator(`.thumb-row[data-slide-index="${index}"]`).click();
        const glyph = editor
          .locator('.paint foreignObject > div')
          .filter({ hasText: 'Underline words' })
          .first();
        const before = await geometry(glyph);
        const expected = await decorations(glyph);
        const point = await glyph.evaluate((root) => {
          const text = document.createTreeWalker(root, NodeFilter.SHOW_TEXT).nextNode();
          const range = document.createRange();
          range.setStart(text, 1);
          range.setEnd(text, 3);
          const r = range.getBoundingClientRect();
          return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
        });
        const hit = editor.locator('.hit').first();
        const position = await hit.evaluate((node, point) => {
          const rect = node.getBoundingClientRect();
          return { x: point.x - rect.x, y: point.y - rect.y };
        }, point);
        await hit.click({ position });
        const input = editor.locator('.inline-edit').first();
        await input.waitFor();
        await input.evaluate((node) => (node.style.caretColor = 'transparent'));
        assert.deepEqual(
          await decorations(input),
          expected,
          `${underline}: decoration changed on edit entry`,
        );
        const after = await geometry(input);
        for (const dimension of ['x', 'y', 'width', 'height'])
          assert.ok(
            Math.abs(before[dimension] - after[dimension]) < 2,
            `${underline}: ${dimension} shifted`,
          );
        const strike = await input
          .locator('span')
          .evaluateAll((nodes) =>
            nodes
              .filter((node) => getComputedStyle(node).textDecorationLine.includes('line-through'))
              .map((node) => getComputedStyle(node).textDecorationStyle),
          );
        assert.deepEqual(strike, ['solid']);
        const png = (await input.screenshot()).toString('base64');
        assert.ok(!images.has(png), `${underline} looks identical to ${images.get(png)}`);
        images.set(png, underline);
        await input.press('Escape');
        await input.waitFor({ state: 'hidden' });
      }
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test('table editing uses the resolved theme color for text and patterned underlines', async () => {
  const bundle = await build({
    stdin: {
      contents: `import {createPresentation,addBlankSlide,addSlideTable,getTableCells,setTableCellTextFormat,setPresentationTheme,inches} from '@office-kit/pptx';
        import {inlineTextHtml} from './inline-text-html.ts';
        export function html() {
          const pres=createPresentation();
          setPresentationTheme(pres,{dark1:'#123456'});
          const slide=addBlankSlide(pres);
          const table=addSlideTable(slide,{x:inches(1),y:inches(1),w:inches(4),h:inches(1),rows:[['Theme text']]});
          setTableCellTextFormat(getTableCells(table)[0][0],{color:'scheme:tx1',underline:'dashLong'});
          return inlineTextHtml(pres,table,undefined,{row:0,col:0});
        }`,
      loader: 'ts',
      resolveDir: fileURLToPath(new URL('../../../../site/src/lib/editor/core/', import.meta.url)),
    },
    bundle: true,
    format: 'esm',
    platform: 'browser',
    write: false,
  });
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage();
    const result = await page.evaluate(async (code) => {
      const module = await import(
        URL.createObjectURL(new Blob([code], { type: 'text/javascript' }))
      );
      document.body.innerHTML = module.html();
      const underline = document.querySelector('u');
      return {
        text: document.body.textContent,
        color: getComputedStyle(underline).color,
        image: decodeURIComponent(getComputedStyle(underline).backgroundImage),
      };
    }, bundle.outputFiles[0].text);
    assert.equal(result.text, 'Theme text');
    assert.equal(result.color, 'rgb(18, 52, 86)');
    assert.ok(result.image.includes('stroke="#123456"'));
  } finally {
    await browser.close();
  }
});

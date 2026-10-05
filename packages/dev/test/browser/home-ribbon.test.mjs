import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getParagraphPropertiesEffective,
  getShapeFill,
  getSlideLayout,
  getSlideLayoutName,
  getSlides,
  getSlideShapes,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

// Mac PowerPoint 16 collapses Home groups as its window narrows: Drawing
// first (ribbon < 1300px), then Slides, Paragraph and Insert (< 1080px), then
// Font (< 840px). Clipboard always stays expanded.
const collapsedAt = (width) => [
  ...(width < 1080 ? ['Slides'] : []),
  ...(width < 840 ? ['Font'] : []),
  ...(width < 1080 ? ['Paragraph', 'Insert'] : []),
  ...(width < 1300 ? ['Drawing'] : []),
];
const JA = {
  Slides: 'スライド',
  Font: 'フォント',
  Paragraph: '段落',
  Insert: '挿入',
  Drawing: '図形描画',
};

async function withDeck(source, run) {
  const dir = await mkdtemp(join(tmpdir(), 'office-home-ribbon-'));
  const file = join(dir, 'deck.tsx');
  await writeFile(file, source);
  let preview;
  let browser;
  try {
    preview = await startPreview(file);
    browser = await chromium.launch({ headless: true });
    await run(preview, browser);
  } finally {
    await browser?.close();
    await preview?.close();
    await rm(dir, { recursive: true, force: true });
  }
}

test(
  'home ribbon collapses groups in PowerPoint order without horizontal scrolling',
  { timeout: 120000 },
  () =>
    withDeck(
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={6} height={1}>Ribbon test</Text></Slide></Presentation>`,
      async (preview, browser) => {
        for (const width of [756, 900, 1200, 1512, 2100]) {
          const page = await browser.newPage({ viewport: { width, height: 800 } });
          try {
            await page.goto(preview.url + '/editor');
            await page.getByText('Saved to this project', { exact: true }).waitFor();
            for (const locale of ['en', 'ja']) {
              await page.locator('.lang select').selectOption(locale);
              const label = (name) => (locale === 'ja' ? JA[name] : name);
              const panel = page.locator('#ribbon-panel');
              assert.equal(
                await panel.evaluate((node) => node.scrollWidth <= node.clientWidth + 1),
                true,
                `${width}px ${locale} Home ribbon scrolls`,
              );
              const homeWidth = await page.locator('.home').evaluate((node) => node.clientWidth);
              const triggers = await page
                .locator('.home .group-trigger')
                .evaluateAll((nodes) =>
                  nodes.map((node) => node.closest('section').getAttribute('aria-label')),
                );
              assert.deepEqual(
                triggers,
                collapsedAt(homeWidth).map(label),
                `${width}px ${locale} (ribbon ${homeWidth}px) collapsed groups`,
              );
              await page
                .getByRole('button', { name: locale === 'ja' ? '貼り付け' : 'Paste', exact: true })
                .waitFor();

              const drawing = page
                .locator('.home .group-trigger')
                .filter({ hasText: label('Drawing') });
              if (await drawing.count()) await drawing.click();
              await page
                .getByRole('button', { name: locale === 'ja' ? '配置' : 'Arrange', exact: true })
                .click();
              const selectionPane = page.getByRole('menuitemcheckbox', {
                name: locale === 'ja' ? '選択ウィンドウ...' : 'Selection Pane...',
                exact: true,
              });
              await selectionPane.waitFor();
              await page.keyboard.press('Escape');
              await selectionPane.waitFor({ state: 'hidden' });
              if (await drawing.count()) {
                await page.keyboard.press('Escape');
                await page.locator('.group-popup').waitFor({ state: 'hidden' });
              }
            }
          } finally {
            await page.close();
          }
        }
      },
    ),
);

test(
  'home ribbon paragraph, layout and format painter commands save and undo',
  { timeout: 120000 },
  () =>
    withDeck(
      `import {Presentation,Slide,Shape} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Shape preset="rect" x={1} y={1} width={3} height={1} fill="#FF0000" text="Source" /><Shape preset="rect" x={6} y={1} width={3} height={1} text="Target" /></Slide></Presentation>`,
      async (preview, browser) => {
        const page = await browser.newPage({ viewport: { width: 1512, height: 860 } });
        await page.goto(preview.url + '/editor');
        const saved = () => page.getByText('Saved to this project', { exact: true }).waitFor();
        const deck = async () =>
          loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          );
        const shapes = async () => getSlideShapes(getSlides(await deck())[0]);
        const changed = async (action) => {
          const before = await (await fetch(preview.url + '/editor/state')).json();
          await action();
          for (let i = 0; i < 200; i += 1) {
            const state = await (await fetch(preview.url + '/editor/state')).json();
            if (!state.building && state.revision !== before.revision) return;
            await new Promise((resolve) => setTimeout(resolve, 50));
          }
          throw new Error('The edit was not saved');
        };
        await saved();
        const hits = page.locator('.hit');

        await hits.nth(1).click();
        await changed(() => page.getByRole('button', { name: 'Bullets', exact: true }).click());
        let target = (await shapes())[1];
        const bullet = getParagraphPropertiesEffective(await deck(), target, 0).bullet;
        assert.ok(bullet === 'bullet' || bullet?.char === '•', JSON.stringify(bullet));
        assert.equal(
          await page
            .getByRole('button', { name: 'Bullets', exact: true })
            .getAttribute('aria-pressed'),
          'true',
        );
        await changed(() =>
          page.getByRole('button', { name: 'Increase List Level', exact: true }).click(),
        );
        target = (await shapes())[1];
        assert.equal(getParagraphPropertiesEffective(await deck(), target, 0).level, 1);
        await changed(() => page.getByTitle('Undo (Ctrl+Z)', { exact: true }).click());
        target = (await shapes())[1];
        assert.equal(getParagraphPropertiesEffective(await deck(), target, 0).level, 0);

        await hits.nth(0).click();
        const painter = page.getByRole('button', { name: 'Format Painter', exact: true });
        await painter.click();
        assert.equal(await painter.getAttribute('aria-pressed'), 'true');
        await changed(() => hits.nth(1).click());
        assert.deepEqual(getShapeFill((await shapes())[1]), getShapeFill((await shapes())[0]));
        assert.equal(await painter.getAttribute('aria-pressed'), 'false');

        await changed(async () => {
          await page.getByRole('button', { name: 'Layout', exact: true }).click();
          await page.getByRole('menuitemradio', { name: 'Title and Content', exact: true }).click();
        });
        assert.equal(
          getSlideLayoutName(getSlideLayout(getSlides(await deck())[0])),
          'Title and Content',
        );
      },
    ),
);

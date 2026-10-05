import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { startPreview } from '../helpers/server.mjs';

const DECK = `import { Presentation, Slide, Shape } from '@office-kit/pptx-dsl';
const rows = ['a', 'b'];
export default (
  <Presentation>
    <Slide>
      <Shape preset="rect" x={1} y={1} width={2} height={1} />
      {rows.map((row, i) => <Shape preset="ellipse" x={5} y={1 + i * 1.5} width={1} height={1} />)}
    </Slide>
  </Presentation>
);
`;

test('editor moves are written back into the TSX source', { timeout: 120000 }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-write-back-'));
  const file = join(dir, 'deck.tsx');
  await writeFile(file, DECK);
  process.env.OFFICE_KIT_TSX_WRITE_BACK = '1';
  let preview;
  let browser;
  try {
    preview = await startPreview(file);
    browser = await chromium.launch({ headless: true });
    const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
    await page.goto(preview.url + '/editor');
    const state = async () => (await fetch(preview.url + '/editor/state')).json();
    const settled = async (predicate) => {
      for (let i = 0; i < 300; i += 1) {
        const current = await state();
        if (!current.building && predicate(current)) return current;
        await new Promise((resolve) => setTimeout(resolve, 50));
      }
      throw new Error('The preview did not settle: ' + JSON.stringify(await state()));
    };
    await page.getByText('Saved to this project', { exact: true }).waitFor();
    const before = (await state()).previewRevision;

    // The rectangle is a single literal element: its x is rewritten in place
    // and no sidecar is kept.
    await page.locator('.hit').nth(0).click();
    await page.keyboard.press('Escape');
    await page.locator('.hit').nth(0).click();
    await page.keyboard.press('Shift+ArrowRight');
    await settled((current) => current.previewRevision > before && !current.hasEdits);
    const written = await readFile(file, 'utf8');
    assert.match(
      written,
      /<Shape preset="rect" x=\{(?!1\})[\d.]+\} y=\{1\} width=\{2\} height=\{1\} \/>/,
    );

    // An ellipse comes from rows.map, so its element makes two shapes; the
    // move stays in the sidecar and the source is untouched.
    const revision = (await state()).previewRevision;
    await page.locator('.hit').nth(1).click();
    await page.keyboard.press('Shift+ArrowDown');
    await settled((current) => current.previewRevision > revision && current.hasEdits);
    assert.equal(await readFile(file, 'utf8'), written);
  } finally {
    delete process.env.OFFICE_KIT_TSX_WRITE_BACK;
    await browser?.close();
    await preview?.close();
    await rm(dir, { recursive: true, force: true });
  }
});

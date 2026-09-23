// Records the editor demo used in the pull request. Not a test: run it with
// `node packages/dev/test/browser/demo-recording.mjs <output-dir>`.
//
// It drives the same preview server the browser tests use, so what it records
// is the editor as it actually runs, not a mock-up.

import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { chromium } from 'playwright';
import { startPreview } from '../helpers/server.mjs';

const outDir = process.argv[2] ?? join(tmpdir(), 'office-kit-demo');

const DECK = `import {Presentation,Slide,Text,Shape,Table} from '@office-kit/pptx-dsl';

export default (
  <Presentation>
    <Slide>
      <Text x={0.8} y={0.7} width={8.4} height={1.1} size={40} bold color="#1F3864">
        @office-kit/pptx editor
      </Text>
      <Text x={0.8} y={1.9} width={8.4} height={0.8} size={20} color="#44546A">
        PowerPoint スタイルの編集を、プレビューの中で
      </Text>
      <Shape preset="roundRect" x={0.8} y={3.0} width={3.6} height={1.6} fill="#2E75B6" />
      <Text x={1.1} y={3.4} width={3.0} height={0.8} size={18} color="#FFFFFF">
        図形もテキストも
      </Text>
    </Slide>
    <Slide>
      <Text x={0.8} y={0.7} width={8.4} height={1.0} size={34} bold color="#1F3864">
        Slide two
      </Text>
      <Table
        x={0.8}
        y={2.0}
        width={8.4}
        height={2.0}
        rows={[
          ['Feature', 'Status'],
          ['Layout editing', 'New'],
          ['Slide numbers', 'New'],
        ]}
      />
    </Slide>
  </Presentation>
);
`;

const pause = (page, ms) => page.waitForTimeout(ms);

const run = async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-kit-demo-src-'));
  const file = join(dir, 'deck.tsx');
  await writeFile(file, DECK);
  const preview = await startPreview(file);
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({
    viewport: { width: 1440, height: 900 },
    recordVideo: { dir: outDir, size: { width: 1440, height: 900 } },
  });
  const page = await context.newPage();
  try {
    await page.goto(preview.url);
    await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
    const editor = page.frameLocator('#editor-frame');
    const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
    await saved();
    await pause(page, 1200);

    // 1. Edit text in place. Ctrl+Enter commits; Escape would discard it.
    await editor.locator('.hit').first().dblclick();
    await pause(page, 500);
    await page.keyboard.press('ControlOrMeta+A');
    await page.keyboard.type('Editing, live', { delay: 60 });
    await page.keyboard.press('ControlOrMeta+Enter');
    await saved();
    await pause(page, 900);

    // 2. Character effects from the ribbon's Font group.
    await editor.locator('.hit').first().click();
    await editor.getByRole('button', { name: 'Home', exact: true }).click();
    await pause(page, 400);
    await editor
      .locator('.ribbon')
      .getByRole('button', { name: 'Text format', exact: true })
      .click();
    const dialog = editor.getByRole('dialog');
    await dialog.waitFor();
    await pause(page, 600);
    await dialog.getByLabel('Italic', { exact: true }).check();
    await dialog.getByLabel('Color', { exact: true }).fill('C00000');
    await pause(page, 500);
    await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
    await saved();
    await pause(page, 800);

    // 3. Move an object on the canvas.
    const shape = editor.locator('.hit').nth(2);
    const box = await shape.boundingBox();
    if (box) {
      await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
      await page.mouse.down();
      for (let step = 1; step <= 12; step++) {
        await page.mouse.move(box.x + box.width / 2 + step * 14, box.y + box.height / 2 + step * 4);
        await pause(page, 30);
      }
      await page.mouse.up();
      await saved();
    }
    await pause(page, 800);

    // 4. Deck-wide slide numbers.
    await editor.locator('.paint').click({ position: { x: 5, y: 5 } });
    await editor.getByLabel('Slide numbers', { exact: true }).check();
    await saved();
    await pause(page, 1000);

    // 5. Layout editing — the shared design behind the slide.
    const pane = editor.getByRole('region', { name: 'Layout', exact: true });
    await pane.getByLabel('Layout name', { exact: true }).fill('Brand base');
    await pane.getByLabel('Layout name', { exact: true }).press('Tab');
    await saved();
    await pane.getByLabel('Layout background color', { exact: true }).fill('#eef2f8');
    await saved();
    await pause(page, 1200);

    // 6. Japanese, live.
    await editor.locator('.lang select').selectOption('ja');
    await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
    await pause(page, 1500);

    // 7. The second slide, and undo.
    await editor.locator('.thumb-row').nth(1).click();
    await pause(page, 1200);
    await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
    await pause(page, 1500);
  } finally {
    await context.close();
    await browser.close();
    await preview.close();
    await rm(dir, { recursive: true, force: true });
  }
  console.log('video written under', outDir);
};

await run();

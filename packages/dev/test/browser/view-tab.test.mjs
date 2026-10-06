import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getSlideLayout, getSlideLayoutName, getSlides, loadPresentation } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const DECK = `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={4} height={1}>One</Text></Slide></Presentation>`;

test(
  'View tab: Notes Page, Slide Master tab and unavailable masters',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-view-tab-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(file, DECK);
    let preview;
    let browser;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1512, height: 900 } });
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
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
      const panel = editor.locator('#ribbon-panel');
      const button = (name) => panel.getByRole('button', { name, exact: true });
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.getByRole('tab', { name: 'View', exact: true }).click();
      for (const name of ['Handout Master', 'Notes Master', 'Macros'])
        assert.equal(await button(name).isDisabled(), true, name);

      // Notes Page shows the slide above its editable notes.
      await button('Notes Page').click();
      const notesPage = editor.getByRole('region', { name: 'Notes Page' });
      await notesPage.locator('svg').first().waitFor();
      assert.equal((await notesPage.getByRole('textbox').count()) > 0, true);
      assert.equal(await editor.locator('.canvas-shell').count(), 0);
      await button('Normal').click();
      await editor.locator('.canvas-shell').waitFor();

      // Slide Master opens its own tab; Rename edits the current layout.
      await button('Slide Master').click();
      assert.equal(
        await editor
          .getByRole('tab', { name: 'Slide Master', exact: true })
          .getAttribute('aria-selected'),
        'true',
      );
      await button('Rename').click();
      const field = editor.getByRole('dialog').getByRole('textbox').first();
      await field.fill('Custom Layout');
      await changed(() => field.press('Enter'));
      const deck = await loadPresentation(
        new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
      );
      assert.equal(getSlideLayoutName(getSlideLayout(getSlides(deck)[0])), 'Custom Layout');
      await button('Close Master').click();
      assert.equal(await editor.getByRole('tab', { name: 'Slide Master', exact: true }).count(), 0);
      // Layout editing is no longer at the end of Design.
      await editor.getByRole('tab', { name: 'Design', exact: true }).click();
      assert.equal(await button('Rename layout').count(), 0);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

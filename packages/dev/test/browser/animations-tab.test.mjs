import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getSlideAnimations, getSlides, loadPresentation } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const DECK = `import {Presentation,Slide,Shape} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Shape preset="rect" x={1} y={1} width={2} height={1} text="A" /><Shape preset="ellipse" x={5} y={3} width={2} height={1} text="B" /></Slide></Presentation>`;

test(
  'Animations tab: Exit Effects menu, Preview and Animation Painter',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-animations-tab-'));
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
      const animations = async () =>
        getSlideAnimations(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        );
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
      await editor.getByRole('tab', { name: 'Animations', exact: true }).click();
      for (const name of ['Path Animation', 'Trigger', 'Preview'])
        assert.equal(await button(name).isDisabled(), true, name);
      // Exit effects are not in the inline gallery.
      assert.equal(await panel.getByRole('radio', { name: 'Fade Out' }).count(), 0);

      const hits = editor.locator('.hit');
      await hits.nth(0).click();
      await button('Exit Effects').click();
      await changed(() => panel.getByRole('menuitemradio', { name: 'Fade Out' }).click());
      assert.deepEqual(
        (await animations()).map((step) => step.effect),
        ['fadeOut'],
      );

      // Animation Painter gives the next selected object the same effect.
      await button('Animation Painter').click();
      assert.equal(await button('Animation Painter').getAttribute('aria-pressed'), 'true');
      await changed(() => hits.nth(1).click());
      const steps = await animations();
      assert.deepEqual(
        steps.map((step) => step.effect),
        ['fadeOut', 'fadeOut'],
      );
      assert.equal(new Set(steps.flatMap((step) => step.targetShapeIds)).size, 2);
      assert.equal(await button('Animation Painter').getAttribute('aria-pressed'), 'false');

      await button('Preview').click();
      const player = editor.getByRole('dialog', { name: 'Play animations' });
      await player.waitFor();
      await page.keyboard.press('Escape');
      await player.waitFor({ state: 'detached' });
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

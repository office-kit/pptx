// The editor's animation pane, driven the way a person drives it: objects are
// picked by name, never by the id the file uses as a handle.

import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { chromium } from 'playwright';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  getShapeId,
  getSlideAnimations,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeAnimation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'animations are added, changed, reordered, removed and played from the panel',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-animation-panel-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide>` +
          `<Text x={1} y={1} width={4} height={1}>Alpha</Text>` +
          `<Text x={1} y={3} width={4} height={1}>Beta</Text>` +
          `<Text x={1} y={5} width={4} height={2}>{'One\\nTwo'}</Text>` +
          `</Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      // The editor runs in an iframe, and a request made from one is routed by
      // the context rather than the page.
      const context = await browser.newContext({ viewport: { width: 1500, height: 1000 } });
      const page = await context.newPage();
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      let ja = false;
      const label = (en, jp) => (ja ? jp : en);
      const saved = () =>
        editor.getByText(label('Saved to this project', 'このプロジェクトに保存済み')).waitFor();
      /** What the deck on disk says, which is what a reload would show. */
      const stored = async () =>
        getSlideAnimations(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        ).map((step) => ({
          target: step.targetShapeIds,
          effect: step.effect,
          start: step.start,
          durationMs: step.durationMs,
          delayMs: step.delayMs,
          byParagraph: step.buildByParagraph,
        }));
      const pane = () =>
        editor.getByRole('region', { name: label('Animations', 'アニメーション') });
      const add = () =>
        pane().getByRole('button', { name: label('Add animation', 'アニメーションを追加') });
      await saved();
      assert.deepEqual(await stored(), []);

      // --- Adding, by name ------------------------------------------------
      const objects = pane().getByLabel(label('Object', '対象'));
      assert.deepEqual(
        (await objects.locator('option').allTextContents()).map((text) => text.trim()),
        ['TextBox 2 — Alpha', 'TextBox 3 — Beta', 'TextBox 4 — One Two'],
        'objects are offered by name and the text they carry',
      );
      await objects.selectOption({ label: 'TextBox 3 — Beta' });
      await pane().getByLabel(label('New effect', '追加する効果')).selectOption('fadeIn');
      await pane().getByLabel(label('New duration (ms)', '追加する長さ (ミリ秒)')).fill('600');
      await add().click();
      await saved();
      const first = await stored();
      assert.equal(first.length, 1);
      assert.equal(first[0].effect, 'fadeIn');
      assert.equal(first[0].start, 'click');
      assert.equal(first[0].durationMs, 600);
      const beta = first[0].target[0];

      // The row names the object it animates, not a number from the file.
      await assert.doesNotReject(pane().getByText('TextBox 3 — Beta').first().waitFor());

      // --- Changing --------------------------------------------------------
      await pane()
        .getByLabel(`${label('Effect', '効果')} 1`)
        .selectOption('appear');
      await saved();
      assert.equal((await stored())[0].effect, 'appear');
      await pane()
        .getByLabel(`${label('Start', '開始')} 1`)
        .selectOption('afterPrevious');
      await saved();
      // The slide's first effect has nothing before it, so 'after previous'
      // means it starts as the slide appears.
      assert.equal((await stored())[0].start, 'afterPrevious');
      await pane()
        .getByLabel(`${label('Delay (ms)', '遅延 (ミリ秒)')} 1`)
        .fill('250');
      await pane()
        .getByLabel(`${label('Delay (ms)', '遅延 (ミリ秒)')} 1`)
        .blur();
      await saved();
      assert.equal((await stored())[0].delayMs, 250);

      // --- A second effect, and the order between them ---------------------
      await objects.selectOption({ label: 'TextBox 2 — Alpha' });
      await pane().getByLabel(label('New effect', '追加する効果')).selectOption('fadeOut');
      await add().click();
      await saved();
      const two = await stored();
      assert.equal(two.length, 2);
      assert.equal(two[1].effect, 'fadeOut');
      const alpha = two[1].target[0];
      assert.deepEqual([two[0].target[0], two[1].target[0]], [beta, alpha]);

      await pane()
        .getByRole('button', { name: `${label('Move earlier', '前へ移動')} 2` })
        .click();
      await saved();
      assert.deepEqual(
        (await stored()).map((step) => step.target[0]),
        [alpha, beta],
        'the second effect moved ahead of the first',
      );

      // --- Paragraph builds ------------------------------------------------
      await objects.selectOption({ label: 'TextBox 4 — One Two' });
      await pane().getByLabel(label('New effect', '追加する効果')).selectOption('fadeIn');
      await pane().getByLabel(label('New by paragraph', '追加時に段落ごと')).check();
      await add().click();
      await saved();
      const built = await stored();
      assert.equal(built.length, 4, 'a paragraph build is one effect per paragraph');
      assert.ok(built.slice(2).every((step) => step.byParagraph));

      // --- Undo, redo, reload ----------------------------------------------
      await editor.getByTitle(label('Undo (Ctrl+Z)', '元に戻す (Ctrl+Z)'), { exact: true }).click();
      await saved();
      assert.equal((await stored()).length, 2);
      await editor.getByTitle(label('Redo (Ctrl+Y)', 'やり直し (Ctrl+Y)'), { exact: true }).click();
      await saved();
      assert.equal((await stored()).length, 4);
      await page.reload();
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      await saved();
      assert.equal((await stored()).length, 4);

      // --- Removing --------------------------------------------------------
      const before = await stored();
      await pane()
        .getByRole('button', { name: `${label('Delete animation', 'アニメーションの削除')} 1` })
        .click();
      await saved();
      const after = await stored();
      assert.equal(after.length, before.length - 1);
      assert.ok(after.every((step) => step.target[0] !== alpha));

      // --- Japanese --------------------------------------------------------
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await assert.doesNotReject(pane().waitFor());
      await assert.doesNotReject(
        pane().getByRole('button', { name: 'アニメーションを再生' }).waitFor(),
      );
      await pane().getByLabel('効果 1').selectOption('fadeOut');
      await saved();
      assert.equal((await stored())[0].effect, 'fadeOut');

      // --- Playing what is in hand, not what is saved -----------------------
      // The save is held open, so the deck on disk provably still carries the
      // old length while the panel plays the new one.
      // The callback itself hands the held route over, so awaiting it is proof
      // the save is stopped — a request event would only say it was sent.
      let hold;
      const held = new Promise((resolve) => {
        hold = resolve;
      });
      const savePath = (url) => url.pathname === '/editor/document';
      await context.route(savePath, async (route) => {
        if (route.request().method() !== 'PUT' || hold === undefined) {
          await route.continue();
          return;
        }
        const take = hold;
        hold = undefined;
        take(route);
      });
      const lengthBox = pane().getByLabel('長さ (ミリ秒) 1');
      const wasStored = (await stored())[0].durationMs;
      assert.notEqual(wasStored, 1200);
      await lengthBox.fill('1200');
      await lengthBox.blur();
      const heldSave = await Promise.race([
        held,
        new Promise((_, reject) =>
          setTimeout(() => reject(new Error('the save was never held')), 20000),
        ),
      ]);
      assert.equal((await stored())[0].durationMs, wasStored, 'the save has not landed');

      const play = pane().getByRole('button', { name: 'アニメーションを再生' });
      await play.click();
      const dialog = editor.getByRole('dialog', { name: 'アニメーションを再生' });
      await dialog.waitFor();
      // The paragraphs still to be revealed are the ones held back; the object
      // whose only effect takes it *off* the slide is still on it.
      const shown = () =>
        dialog
          .locator('[data-pptx-paragraph]')
          .evaluateAll((nodes) => nodes.map((node) => node.style.visibility || 'visible'));
      assert.ok((await shown()).includes('hidden'), 'the show starts before the entrances');
      // This slide's first effect runs 'after previous' with nothing before it,
      // so the show starts it without a click.
      await assert.doesNotReject(dialog.getByText('クリック 1 / 3').waitFor());

      // The effect that runs is the one the panel is holding: 1200ms, not the
      // length the file still has.
      // It waits on its delay first, so this waits for the effect to start
      // rather than for a stopwatch.
      const running = await dialog.locator('.slide').evaluateAll(async (nodes) => {
        const stage = nodes[0];
        for (let i = 0; i < 300; i++) {
          const found = stage.getAnimations({ subtree: true });
          if (found.length > 0)
            return found.map((animation) => animation.effect.getTiming().duration);
          await new Promise((resolve) => setTimeout(resolve, 20));
        }
        return [];
      });
      assert.deepEqual(running, [1200], 'the unsaved length is the one being played');

      // The editor behind the show takes neither keys nor edits.
      await page.keyboard.press('Control+z');
      await page.keyboard.press('Control+k');
      await page.waitForTimeout(150);
      await assert.doesNotReject(dialog.waitFor());
      assert.equal(await editor.locator('.palette').count(), 0, 'no command palette behind it');

      // The arrows drive the show: the dialog takes the focus itself rather
      // than leaving it on a control that would swallow them.
      await page.keyboard.press('ArrowRight');
      await assert.doesNotReject(dialog.getByText('クリック 2 / 3').waitFor());
      await page.keyboard.press('ArrowLeft');
      await assert.doesNotReject(dialog.getByText('クリック 1 / 3').waitFor());
      await dialog.getByRole('button', { name: '最初から' }).click();
      // A control with the focus keeps its own keyboard: Enter presses it,
      // rather than counting as the show's next click.
      await dialog.getByRole('button', { name: '次へ' }).focus();
      await page.keyboard.press('Enter');
      await assert.doesNotReject(dialog.getByText('クリック 2 / 3').waitFor());
      await dialog.getByRole('button', { name: '停止' }).focus();
      await page.keyboard.press('Enter');
      await dialog.waitFor({ state: 'detached' });
      // …and the focus goes back where the show was started from.
      assert.equal(await editor.locator(':focus').textContent(), 'アニメーションを再生');

      // Escape closes it as well — the dialog's own way out is not swallowed by
      // the keyboard the show holds — and gives the focus back the same way.
      await play.click();
      await dialog.waitFor();
      await page.keyboard.press('Escape');
      await dialog.waitFor({ state: 'detached' });
      assert.equal(await editor.locator(':focus').textContent(), 'アニメーションを再生');

      // The undo that was pressed behind the show did not happen: the panel
      // still holds the change, and it is what is saved once the save lands.
      assert.equal(await lengthBox.inputValue(), '1200');
      await heldSave.continue();
      await context.unroute(savePath);
      await saved();
      assert.equal((await stored())[0].durationMs, 1200);

      // A document changed from outside the show — not by the viewer, whose
      // keys and clicks the dialog holds — leaves it playing a deck that no
      // longer exists, so it closes. Pressed programmatically, since a real
      // click cannot reach the editor while the dialog is up.
      await play.click();
      await dialog.waitFor();
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).dispatchEvent('click');
      await dialog.waitFor({ state: 'detached' });
      await saved();
      assert.equal((await stored())[0].durationMs, wasStored, 'the outside change went through');

      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'the panel keeps an effect it cannot change, tells two slides apart, and shows a 4:3 slide as one',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-animation-kept-'));
    let preview, browser;
    try {
      const pres = createPresentation({ size: '4:3' });
      const one = addBlankSlide(pres);
      const at = (y) => ({ x: inches(1), y: inches(y), w: inches(4), h: inches(1) });
      const alpha = addSlideTextBox(one, { ...at(1), text: 'Alpha' });
      const beta = addSlideTextBox(one, { ...at(3), text: 'Beta' });
      setShapeAnimation(alpha, { effect: 'fadeIn' });
      setShapeAnimation(beta, { effect: 'fadeIn', durationMs: 600 });
      const two = addBlankSlide(pres);
      // Shape ids run per slide, so these two carry the same ids as Alpha and
      // Beta do — which is what makes a remembered id the wrong way to name an
      // object across slides.
      const gamma = getShapeId(addSlideTextBox(two, { ...at(1), text: 'Gamma' }));
      const delta = getShapeId(addSlideTextBox(two, { ...at(3), text: 'Delta' }));
      assert.deepEqual([gamma, delta], [getShapeId(alpha), getShapeId(beta)], 'the ids repeat');

      const parts = unzipSync(await savePresentation(pres));
      const slidePath = 'ppt/slides/slide1.xml';
      const slideXml = strFromU8(parts[slidePath]);
      // Alpha's effect is written first, so this is its node: its value is
      // taken away again when it ends, which this library reads but does not
      // model. The panel has to keep it exactly as it is.
      const patched = slideXml.replace('fill="hold" grpId', 'fill="remove" grpId');
      assert.notEqual(patched, slideXml, 'the effect node was found');
      parts[slidePath] = strToU8(patched);
      // The effect's `<p:par>` holds only behaviours, so its own closing tag is
      // the first one after it.
      const keptEffect = (xml) => {
        const at = xml.indexOf('fill="remove"');
        assert.notEqual(at, -1, 'the kept effect is still there');
        return xml.slice(xml.lastIndexOf('<p:par>', at), xml.indexOf('</p:par>', at) + 8);
      };
      const keptBefore = keptEffect(patched);
      const source = join(dir, 'source.pptx');
      await writeFile(source, zipSync(parts));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';` +
          `export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
      );

      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const context = await browser.newContext({ viewport: { width: 1500, height: 1000 } });
      const page = await context.newPage();
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project').waitFor();
      const stored = async (index) =>
        getSlideAnimations(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[index],
        );
      const pane = () => editor.getByRole('region', { name: 'Animations' });
      await saved();

      // --- What it cannot change, it keeps and explains --------------------
      const first = await stored(0);
      assert.equal(first.length, 2);
      assert.equal(first[0].playable, false, 'the patched effect is read, not played');
      assert.equal(first[0].editable, false);
      assert.equal(first[1].editable, true);
      await assert.doesNotReject(
        pane()
          .getByText(
            'Kept as it is: Uses an effect, target or start condition this library only reads.',
          )
          .waitFor(),
      );
      // It has no controls of its own, and the effect below it still has.
      assert.equal(await pane().getByLabel('Effect 1').count(), 0);
      await pane().getByLabel('Effect 2').selectOption('fadeOut');
      await saved();
      const edited = await stored(0);
      assert.equal(edited.length, 2, 'editing its neighbour did not drop it');
      assert.equal(edited[0].playable, false);
      assert.equal(edited[0].effect, 'fadeIn');
      assert.equal(edited[1].effect, 'fadeOut');
      // Not merely still listed: the node the deck arrived with is the node it
      // leaves with, attribute for attribute.
      const savedXml = strFromU8(
        unzipSync(new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()))[
          slidePath
        ],
      );
      assert.equal(keptEffect(savedXml), keptBefore);

      // --- A 4:3 slide is shown as one -------------------------------------
      await pane().getByRole('button', { name: 'Play animations' }).click();
      const dialog = editor.getByRole('dialog', { name: 'Play animations' });
      await dialog.waitFor();
      const stage = await dialog.locator('.slide').boundingBox();
      assert.ok(
        Math.abs(stage.width / stage.height - 4 / 3) < 0.02,
        `shown at the deck's own ratio (${stage.width} x ${stage.height})`,
      );
      // The kept effect is named rather than passed over in silence.
      await assert.doesNotReject(
        dialog.getByText('1 × Uses an effect, target or start this library only reads').waitFor(),
      );
      await page.keyboard.press('Escape');
      await dialog.waitFor({ state: 'detached' });

      // --- The next slide's objects are its own -----------------------------
      const objects = pane().getByLabel('Object');
      await objects.selectOption({ label: 'TextBox 3 — Beta' });
      await editor.locator('.thumb-row').nth(1).click();
      assert.deepEqual(
        (await objects.locator('option').allTextContents()).map((text) => text.trim()),
        ['TextBox 2 — Gamma', 'TextBox 3 — Delta'],
        'the second slide offers its own objects',
      );
      await pane().getByRole('button', { name: 'Add animation' }).click();
      await saved();
      const second = await stored(1);
      assert.equal(second.length, 1);
      // Beta's id was chosen on the slide before this one. Carrying it over
      // would have animated Delta, which is drawn with that very id here.
      assert.equal(second[0].target.shapeId, gamma);
      assert.deepEqual(second[0].targetShapeIds, [gamma]);

      await pane().getByRole('button', { name: 'Play animations' }).click();
      await dialog.waitFor();
      const shown = await dialog
        .locator('[data-pptx-shape-id]')
        .evaluateAll((nodes) =>
          nodes.map((node) => [
            node.getAttribute('data-pptx-shape-id'),
            node.style.visibility || 'visible',
          ]),
        );
      // Only the object the new effect names is held back.
      assert.deepEqual(shown, [
        [String(gamma), 'hidden'],
        [String(delta), 'visible'],
      ]);
      await page.keyboard.press('Escape');
      await dialog.waitFor({ state: 'detached' });

      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'a fly is given an edge in the panel, and only a fly is offered one',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-animation-direction-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide>` +
          `<Text x={1} y={1} width={4} height={1}>Alpha</Text>` +
          `</Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const context = await browser.newContext({ viewport: { width: 1500, height: 1000 } });
      const page = await context.newPage();
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      let ja = false;
      const label = (en, jp) => (ja ? jp : en);
      const saved = () =>
        editor.getByText(label('Saved to this project', 'このプロジェクトに保存済み')).waitFor();
      const stored = async () =>
        getSlideAnimations(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        ).map((step) => ({ effect: step.effect, direction: step.direction }));
      const pane = () =>
        editor.getByRole('region', { name: label('Animations', 'アニメーション') });
      /**
       * What the deck on disk says once the edit has landed. The save
       * indicator is already showing "saved" from the edit before this one, so
       * waiting on it can step straight past the write being checked.
       */
      const settles = async (expected) => {
        const deadline = Date.now() + 20000;
        let last;
        for (;;) {
          last = await stored();
          if (JSON.stringify(last) === JSON.stringify(expected)) return;
          if (Date.now() > deadline) break;
          await new Promise((resolve) => setTimeout(resolve, 100));
        }
        assert.deepEqual(last, expected);
      };
      await saved();

      // --- Adding a fly ----------------------------------------------------
      // The edge is only offered once the effect that uses it is picked.
      const newDirection = () => pane().getByLabel(label('New direction', '追加する方向'));
      assert.equal(await newDirection().count(), 0, 'no edge for a fade');
      await pane().getByLabel(label('New effect', '追加する効果')).selectOption('flyIn');
      await newDirection().selectOption('left');
      await pane()
        .getByRole('button', { name: label('Add animation', 'アニメーションを追加') })
        .click();
      await settles([{ effect: 'flyIn', direction: 'left' }]);

      // --- Changing the edge -----------------------------------------------
      const direction = () => pane().getByLabel(`${label('Direction', '方向')} 1`);
      assert.equal(await direction().inputValue(), 'left');
      await direction().selectOption('top');
      await settles([{ effect: 'flyIn', direction: 'top' }]);

      // --- Changing the preset ---------------------------------------------
      // A spin states no edge, so the file stops carrying one and the row stops
      // offering one. The panel must not hand the old edge to the new preset.
      await pane()
        .getByLabel(`${label('Effect', '効果')} 1`)
        .selectOption('spin');
      await settles([{ effect: 'spin', direction: null }]);
      assert.equal(await direction().count(), 0, 'no edge for a spin');

      // Back to a fly: it takes this library's own default rather than an edge
      // the row was showing before the spin.
      await pane()
        .getByLabel(`${label('Effect', '効果')} 1`)
        .selectOption('flyOut');
      await settles([{ effect: 'flyOut', direction: 'bottom' }]);

      // --- Japanese --------------------------------------------------------
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await assert.doesNotReject(pane().waitFor());
      assert.deepEqual(
        (await direction().locator('option').allTextContents()).map((text) => text.trim()),
        ['下辺', '上辺', '左辺', '右辺'],
      );
      await direction().selectOption('right');
      await settles([{ effect: 'flyOut', direction: 'right' }]);

      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

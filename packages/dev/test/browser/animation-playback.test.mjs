import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile, mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import * as pptx from '@office-kit/pptx';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';

test('native filter fade metadata drives intermediate playback opacity', async () => {
  const deck = pptx.createPresentation();
  const slide = pptx.addBlankSlide(deck);
  pptx.addSlideTextBox(slide, { x: 0, y: 0, w: 914400, h: 914400, text: 'Fade' });
  const id = pptx.getShapeId(pptx.getSlideShapes(slide)[0]);
  const parts = unzipSync(await pptx.savePresentation(deck));
  const xml = strFromU8(parts['ppt/slides/slide1.xml']);
  const timing = `<p:timing><p:tnLst><p:par><p:cTn id="1" nodeType="tmRoot"><p:childTnLst><p:seq><p:cTn id="2" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3" presetID="10" presetClass="entr" nodeType="clickEffect"><p:childTnLst><p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="4" dur="2000"/><p:tgtEl><p:spTgt spid="${id}"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
  parts['ppt/slides/slide1.xml'] = strToU8(xml.replace('</p:sld>', timing + '</p:sld>'));
  const loaded = await pptx.loadPresentation(zipSync(parts));
  pptx.setSlideAnimationDelay(pptx.getSlides(loaded)[0], '3', 500);
  const groups = pptx.getSlideAnimationSequence(pptx.getSlides(loaded)[0]);
  assert.equal(groups[0][0].durationMs, 2000);
  const source = await readFile(
    new URL('../../src/animation-playback.ts', import.meta.url),
    'utf8',
  );
  const script = source.slice(source.indexOf('`') + 1, source.lastIndexOf('`'));
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage();
    await page.setContent(
      `<svg><g data-pptx-shape-id="${id}"><rect width="40" height="40"/></g></svg>`,
    );
    const opacities = await page.evaluate(
      ({ script, groups }) => {
        window.eval(
          `const canvas=document;const presenting=true,index=0;const state={editor:{slides:[{animations:${JSON.stringify(groups)}}]}};function cancelSlideAdvance(){}function scheduleSlideAdvance(){}${script};initializeShapeAnimations();advanceShapeAnimation();cancelAnimationFrame(shapeAnimationState.frame);window.seekAnimation=renderShapeAnimationState;`,
        );
        return [0, 499, 500, 1000, 2500].map((time) => {
          window.seekAnimation(time);
          return document.querySelector('g').style.opacity;
        });
      },
      { script, groups },
    );
    assert.deepEqual(opacities, ['0', '0', '0', '0.25', '1']);
  } finally {
    await browser.close();
  }
});

test('shape effects advance, finish, rewind and restore original artwork', async () => {
  const source = await readFile(
    new URL('../../src/animation-playback.ts', import.meta.url),
    'utf8',
  );
  const script = source.slice(source.indexOf('`') + 1, source.lastIndexOf('`'));
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage();
    await page.setContent('<div id="slide"></div>');
    await page.evaluate((script) => {
      window.eval(`
        const canvas=document.querySelector('#slide').attachShadow({mode:'open'});
        canvas.innerHTML='<svg><g data-pptx-shape-id="2" style="fill:red"><rect width="50" height="50"/></g><g data-pptx-shape-id="3"/></svg>';
        let presenting=true,index=0;
        const fx=(effect,id,durationMs=0,trigger='clickEffect',delayMs=0)=>({effect,shapeIds:[id],durationMs,trigger,delayMs});
        const state={editor:{slides:[{animations:[[fx('fadeIn','2',10000)],[fx('fadeOut','2',10000)],[fx('appear','2'),fx('disappear','3',0,'afterEffect',200)]]}]}};
        function cancelSlideAdvance(){} function scheduleSlideAdvance(){}
        ${script}
        window.init=initializeShapeAnimations;window.next=advanceShapeAnimation;window.prev=previousShapeAnimation;
        window.seek=t=>{cancelAnimationFrame(shapeAnimationState.frame);renderShapeAnimationState(t);};
        window.stop=()=>{presenting=false;initializeShapeAnimations();};
        window.values=()=>[...canvas.querySelectorAll('g')].map(el=>({opacity:el.style.opacity,visibility:el.style.visibility,style:el.getAttribute('style')}));
      `);
    }, script);
    await page.evaluate(() => window.init());
    assert.equal((await page.evaluate(() => window.values()))[0].visibility, 'hidden');
    assert.equal(await page.evaluate(() => window.next()), true);
    await page.evaluate(() => window.seek(5000));
    assert.equal((await page.evaluate(() => window.values()))[0].opacity, '0.5');
    // A second next finishes the running effect instead of skipping the following one.
    await page.evaluate(() => window.next());
    assert.equal((await page.evaluate(() => window.values()))[0].opacity, '1');
    await page.evaluate(() => {
      window.next();
      window.seek(2500);
    });
    assert.equal((await page.evaluate(() => window.values()))[0].opacity, '0.75');
    await page.evaluate(() => window.next());
    assert.equal((await page.evaluate(() => window.values()))[0].visibility, 'hidden');
    await page.evaluate(() => {
      window.next();
      window.seek(0);
    });
    let values = await page.evaluate(() => window.values());
    assert.equal(values[0].opacity, '1');
    assert.equal(values[1].opacity, '1');
    await page.evaluate(() => window.seek(200));
    assert.equal((await page.evaluate(() => window.values()))[1].visibility, 'hidden');
    await page.evaluate(() => window.next());
    assert.equal(await page.evaluate(() => window.next()), false);
    await page.evaluate(() => window.prev());
    values = await page.evaluate(() => window.values());
    assert.equal(values[0].visibility, 'hidden');
    assert.equal(values[1].visibility, 'visible');
    await page.evaluate(() => window.init());
    assert.equal((await page.evaluate(() => window.values()))[0].visibility, 'hidden');
    await page.evaluate(() => window.stop());
    assert.deepEqual(await page.evaluate(() => window.values()), [
      { opacity: '', visibility: '', style: 'fill:red' },
      { opacity: '', visibility: '', style: null },
    ]);
  } finally {
    await browser.close();
  }
});

test('selected preview skips unselected effects and restores artwork', async () => {
  const source = await readFile(
    new URL('../../src/animation-playback.ts', import.meta.url),
    'utf8',
  );
  const script = source.slice(source.indexOf('`') + 1, source.lastIndexOf('`'));
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage();
    await page.setContent(
      '<svg><g data-pptx-shape-id="1" style="fill:red"/><g data-pptx-shape-id="2"/><g data-pptx-shape-id="3"/></svg>',
    );
    await page.evaluate((script) => {
      window.eval(`
        const canvas=document;let presenting=false;const index=0;
        const fx=(timingId,shape,effect,trigger,delayMs=0)=>({timingId,shapeIds:[shape],effect,trigger,delayMs,durationMs:1000});
        const state={editor:{slides:[{animations:[
          [fx('a','1','fadeOut','clickEffect'),fx('b','2','fadeIn','afterEffect',200),fx('c','3','fadeIn','withEffect',200)],
          [fx('d','2','fadeOut','clickEffect')]
        ]}]}};
        function cancelSlideAdvance(){}function scheduleSlideAdvance(){}
        ${script}
        window.preview=previewAnimations;window.stop=stopAnimationPreview;
        window.seek=t=>{cancelAnimationFrame(shapeAnimationState.frame);renderShapeAnimationState(t);};
        window.next=()=>{cancelAnimationFrame(shapeAnimationState.frame);shapeAnimationState.running=false;advanceShapeAnimation();};
        window.values=()=>[...document.querySelectorAll('g')].map(el=>el.style.cssText);
      `);
    }, script);
    // Input order must not alter slide order. The unselected predecessor has no preview delay.
    await page.evaluate(() => {
      window.preview(['d', 'c', 'b'], true);
      window.seek(700);
    });
    assert.equal(await page.locator('g').nth(0).getAttribute('style'), 'fill:red');
    assert.deepEqual(
      await page.locator('g').evaluateAll((nodes) => nodes.slice(1).map((el) => el.style.opacity)),
      ['0.5', '0.5'],
    );
    await page.evaluate(() => {
      window.next();
      window.seek(500);
    });
    assert.deepEqual(
      await page.locator('g').evaluateAll((nodes) => nodes.slice(1).map((el) => el.style.opacity)),
      ['0.5', '1'],
    );
    await page.evaluate(() => window.stop());
    assert.deepEqual(await page.evaluate(() => window.values()), ['fill: red;', '', '']);
    await page.evaluate(() => {
      window.preview(['a', 'b'], true);
      window.seek(1500);
    });
    assert.deepEqual(
      await page
        .locator('g')
        .evaluateAll((nodes) => nodes.slice(0, 2).map((el) => el.style.opacity)),
      ['0', '0.3'],
    );
    assert.equal(
      await page
        .locator('g')
        .nth(2)
        .evaluate((el) => el.style.opacity),
      '',
    );
    await page.evaluate(() => window.stop());
    await page.evaluate(() => window.preview(['missing'], true));
    assert.deepEqual(await page.evaluate(() => window.values()), ['fill: red;', '', '']);
  } finally {
    await browser.close();
  }
});

test(
  'authored click effects run before navigating and reset after leaving the show',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-animation-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={8} height={1}>Animated</Text></Slide><Slide><Text x={1} y={1} width={8} height={1}>Next slide</Text></Slide></Presentation>`,
    );
    const proc = spawn(process.execPath, [
      fileURLToPath(new URL('../../dist/cli.mjs', import.meta.url)),
      'dev',
      file,
      '--port',
      '0',
    ]);
    let browser;
    try {
      const url = await new Promise((resolve, reject) => {
        let output = '';
        proc.stdout.on('data', (data) => {
          output += data;
          const match = output.match(/Preview: (http:\/\/\S+)/);
          if (match) resolve(match[1]);
        });
        proc.on('error', reject);
        proc.on('exit', (code) => reject(new Error('Server exited ' + code)));
      });
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage();
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(url);
      await page.waitForFunction(() => window.office?.getState().editor?.slides.length === 2);
      await page.locator('#tab-animations').click();
      assert.equal(
        await page.getByRole('button', { name: 'Entrance Appear', exact: true }).isDisabled(),
        true,
      );
      await page
        .locator('.shape-hit')
        .first()
        .click({ position: { x: 4, y: 4 } });
      await page.locator('#tab-animations').click();
      await page.getByRole('button', { name: 'Entrance Appear', exact: true }).click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations.length === 1,
      );
      await page.getByRole('button', { name: 'Exit Fade', exact: true }).click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations.length === 2,
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations.length === 1,
      );
      await page.locator('[data-edit="redo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations.length === 2,
      );
      const markers = page.locator('.animation-marker');
      assert.deepEqual(await markers.allTextContents(), ['1', '2']);
      await page.locator('#tab-home').click();
      assert.equal(await markers.count(), 0);
      await page.locator('#tab-animations').click();
      assert.deepEqual(await markers.allTextContents(), ['1', '2']);
      await page.evaluate(() => window.office.selectSlide(1));
      assert.equal(await markers.count(), 0);
      await page.evaluate(() => window.office.selectSlide(0));
      assert.deepEqual(await markers.allTextContents(), ['1', '2']);
      await page.locator('#animation-pane-toggle').click();
      const duration = page.locator('#animation-duration');
      assert.equal(await duration.isDisabled(), true);
      await markers.first().click();
      assert.equal(await duration.isDisabled(), true);
      await markers.nth(1).click();
      assert.equal(await markers.nth(1).getAttribute('aria-pressed'), 'true');
      assert.equal(
        await page.locator('#animation-list [role="option"]').nth(1).getAttribute('aria-selected'),
        'true',
      );
      const start = page.locator('#animation-start');
      const rows = page.locator('#animation-list [role="option"]');
      await rows.first().click();
      await rows.nth(1).click({ modifiers: ['Shift'] });
      assert.equal(await page.locator('#animation-list [aria-selected="true"]').count(), 2);
      assert.equal(await start.isDisabled(), false);
      assert.deepEqual(
        await markers.evaluateAll((nodes) =>
          nodes.map((node) => node.getAttribute('aria-pressed')),
        ),
        ['true', 'true'],
      );
      await rows.first().click({ modifiers: ['Meta'] });
      assert.equal(await page.locator('#animation-list [aria-selected="true"]').count(), 1);
      await rows.nth(1).press('Meta+a');
      assert.equal(await page.locator('#animation-list [aria-selected="true"]').count(), 2);
      await rows.nth(1).click({ button: 'right' });
      assert.equal(await page.locator('#animation-list [aria-selected="true"]').count(), 2);
      assert.equal(
        await page
          .getByRole('menuitemcheckbox', { name: 'Start On Click', exact: true })
          .getAttribute('aria-checked'),
        'true',
      );
      await page.keyboard.press('Escape');
      await rows.nth(1).press('Delete');
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations.flat().length === 0,
      );
      assert.equal(await page.locator('.shape-hit').count(), 1);
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations.flat().length === 2,
      );
      await page.locator('[data-edit="redo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations.flat().length === 0,
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations.flat().length === 2,
      );
      await rows.first().click();
      await rows.first().press('Shift+ArrowDown');
      assert.equal(await page.locator('#animation-list [aria-selected="true"]').count(), 2);
      await rows.nth(1).press('Shift+F10');
      await page
        .getByRole('menuitemcheckbox', { name: 'Start After Previous', exact: true })
        .click();

      await page.waitForFunction(() =>
        window.office
          .getState()
          .editor.slides[0].animations.flat()
          .every((effect) => effect.trigger === 'afterEffect'),
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(() =>
        window.office
          .getState()
          .editor.slides[0].animations.flat()
          .every((effect) => effect.trigger === 'clickEffect'),
      );
      await page.getByRole('button', { name: 'Entrance Fade', exact: true }).click();
      await page.waitForFunction(() =>
        window.office
          .getState()
          .editor.slides[0].animations.flat()
          .every((effect) => effect.effect === 'fadeIn'),
      );
      await page.locator('#animation-duration').fill('1.7');
      await page.locator('#animation-duration').press('Tab');
      await page.waitForFunction(() =>
        window.office
          .getState()
          .editor.slides[0].animations.flat()
          .every((effect) => effect.durationMs === 1700),
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(() =>
        window.office
          .getState()
          .editor.slides[0].animations.flat()
          .every((effect) => effect.durationMs !== 1700),
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations.flat()[0].effect === 'appear',
      );
      assert.equal(await page.locator('#animation-duration').inputValue(), '');
      await page.locator('#animation-delay').fill('0.35');
      await page.locator('#animation-delay').press('Tab');
      await page.waitForFunction(() =>
        window.office
          .getState()
          .editor.slides[0].animations.flat()
          .every((effect) => effect.delayMs === 350),
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(() =>
        window.office
          .getState()
          .editor.slides[0].animations.flat()
          .every((effect) => !effect.delayMs),
      );
      await page.locator('[data-edit="redo"]').click();
      await page.waitForFunction(() =>
        window.office
          .getState()
          .editor.slides[0].animations.flat()
          .every((effect) => effect.delayMs === 350),
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(() =>
        window.office
          .getState()
          .editor.slides[0].animations.flat()
          .every((effect) => !effect.delayMs),
      );
      await markers.nth(1).click();
      await rows.first().click({ button: 'right' });
      assert.equal(await page.locator('#animation-list [aria-selected="true"]').count(), 1);
      assert.equal(await rows.first().getAttribute('aria-selected'), 'true');
      await page.getByRole('menuitem', { name: 'Remove', exact: true }).click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations.flat().length === 1,
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations.flat().length === 2,
      );
      await markers.nth(1).click();
      const selectedEffectId = await page.evaluate(
        () => window.office.getState().editor.slides[0].animations[1][0].timingId,
      );
      await page.getByRole('button', { name: 'Entrance Fade', exact: true }).click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations[1][0].effect === 'fadeIn',
      );
      assert.equal(
        await page.evaluate(
          () => window.office.getState().editor.slides[0].animations.flat().length,
        ),
        2,
      );
      assert.equal(
        await page.evaluate(
          () => window.office.getState().editor.slides[0].animations[1][0].timingId,
        ),
        selectedEffectId,
      );
      assert.equal(
        await page
          .getByRole('button', { name: 'Entrance Fade', exact: true })
          .getAttribute('aria-pressed'),
        'true',
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations[1][0].effect === 'fadeOut',
      );
      await page.locator('[data-edit="redo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations[1][0].effect === 'fadeIn',
      );
      await page.getByRole('button', { name: 'Exit Fade', exact: true }).click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations[1][0].effect === 'fadeOut',
      );
      // Reselecting the object changes the gallery operation back to adding.
      await page.waitForFunction(() => !document.querySelector('#animation-start').disabled);
      await page
        .locator('.shape-hit')
        .first()
        .click({ position: { x: 4, y: 4 } });
      assert.equal(await start.isDisabled(), true);
      await page.getByRole('button', { name: 'Entrance Fade', exact: true }).click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations.flat().length === 3,
      );
      const originalOrder = await page.evaluate(() =>
        window.office
          .getState()
          .editor.slides[0].animations.flat()
          .map((effect) => effect.timingId),
      );
      await rows.first().click();
      await rows.nth(1).click({ modifiers: ['Shift'] });
      assert.equal(
        await page.locator('#animation-pane [data-edit="animation-earlier"]').isDisabled(),
        true,
      );
      const beforeNoopDrag = await page.evaluate(() => window.office.getState().history);
      await rows.first().dragTo(rows.nth(1), { targetPosition: { x: 10, y: 2 } });
      assert.deepEqual(await page.evaluate(() => window.office.getState().history), beforeNoopDrag);
      const dropBox = await rows.nth(2).boundingBox();
      await rows
        .first()
        .dragTo(rows.nth(2), { targetPosition: { x: dropBox.width / 2, y: dropBox.height - 2 } });
      const movedOrder = [originalOrder[2], originalOrder[0], originalOrder[1]];
      await page.waitForFunction(
        (order) =>
          JSON.stringify(
            window.office
              .getState()
              .editor.slides[0].animations.flat()
              .map((effect) => effect.timingId),
          ) === JSON.stringify(order),
        movedOrder,
      );
      assert.equal(await page.locator('#animation-list [aria-selected="true"]').count(), 2);
      assert.equal(
        await page.locator('#animation-pane [data-edit="animation-later"]').isDisabled(),
        true,
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        (order) =>
          JSON.stringify(
            window.office
              .getState()
              .editor.slides[0].animations.flat()
              .map((effect) => effect.timingId),
          ) === JSON.stringify(order),
        originalOrder,
      );
      await page.locator('[data-edit="redo"]').click();
      await page.waitForFunction(
        (order) =>
          JSON.stringify(
            window.office
              .getState()
              .editor.slides[0].animations.flat()
              .map((effect) => effect.timingId),
          ) === JSON.stringify(order),
        movedOrder,
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        (order) =>
          JSON.stringify(
            window.office
              .getState()
              .editor.slides[0].animations.flat()
              .map((effect) => effect.timingId),
          ) === JSON.stringify(order),
        originalOrder,
      );
      await rows.first().click();
      await rows.nth(1).click({ modifiers: ['Shift'] });
      const pane = page.locator('#animation-pane');
      const tailDropBox = await pane.boundingBox();
      await rows.first().dragTo(pane, {
        targetPosition: { x: tailDropBox.width / 2, y: tailDropBox.height - 20 },
      });
      const animationOrder = () =>
        window.office
          .getState()
          .editor.slides[0].animations.flat()
          .map((effect) => effect.timingId);
      await page.waitForFunction(
        (order) =>
          JSON.stringify(
            window.office
              .getState()
              .editor.slides[0].animations.flat()
              .map((effect) => effect.timingId),
          ) === JSON.stringify(order),
        movedOrder,
      );
      assert.equal(await page.locator('#animation-list [aria-selected="true"]').count(), 2);
      assert.equal(await page.locator('#animation-list [data-drop-position]').count(), 0);
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        (order) =>
          JSON.stringify(
            window.office
              .getState()
              .editor.slides[0].animations.flat()
              .map((effect) => effect.timingId),
          ) === JSON.stringify(order),
        originalOrder,
      );
      assert.deepEqual(await page.evaluate(animationOrder), originalOrder);
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations.flat().length === 2,
      );
      await markers.nth(1).click();
      await start.selectOption('withEffect');
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations[0].length === 2,
      );
      assert.deepEqual(await markers.allTextContents(), ['1', '1']);
      await page.locator('[data-edit="animation-earlier"]').click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.slides[0].animations[0][0].trigger === 'withEffect' &&
          document.querySelector('#animation-start').value === 'withEffect',
      );
      assert.deepEqual(await markers.allTextContents(), ['0', '1']);
      assert.equal(await page.locator('[data-edit="animation-earlier"]').isDisabled(), true);
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations[0].length === 2,
      );
      await page.locator('[data-edit="redo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations.length === 2,
      );
      await page.locator('[data-edit="animation-later"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations[0].length === 2,
      );
      assert.deepEqual(await markers.allTextContents(), ['1', '1']);
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations.length === 2,
      );
      await page.locator('[data-edit="redo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations[0].length === 2,
      );
      await start.selectOption('afterEffect');
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations[0][1].trigger === 'afterEffect',
      );
      // Removing the first click effect reconnects the remaining automatic effect.
      await markers.first().click();
      await page.locator('#animation-remove').click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.slides[0].animations.flat().length === 1 &&
          document.querySelector('#animation-start').value === 'afterEffect',
      );
      assert.deepEqual(await markers.allTextContents(), ['0']);
      assert.equal(await start.inputValue(), 'afterEffect');
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations.flat().length === 2,
      );
      assert.deepEqual(await markers.allTextContents(), ['1', '1']);
      await page.locator('[data-edit="redo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations.flat().length === 1,
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations.flat().length === 2,
      );
      await markers.nth(1).click();
      await start.selectOption('clickEffect');
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations.length === 2,
      );
      const originalDuration = await duration.inputValue();
      await duration.fill('2.4');
      await duration.press('Tab');
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations[1][0].durationMs === 2400,
      );
      assert.equal(
        await page.evaluate(
          () => window.office.getState().editor.slides[0].animations[0][0].durationMs,
        ),
        0,
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        (ms) => window.office.getState().editor.slides[0].animations[1][0].durationMs === ms,
        Number(originalDuration) * 1000,
      );
      await page.locator('[data-edit="redo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations[1][0].durationMs === 2400,
      );
      await page.waitForFunction(() => !document.querySelector('#animation-duration').disabled);
      const stageBox = await page.locator('#stage').boundingBox();
      const beforePreview = await page.evaluate(() => ({
        revision: window.office.getState().revision,
        history: window.office.getState().history,
      }));
      await page.locator('#animation-list button').first().click();
      await page
        .locator('#animation-list button')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      assert.equal(await page.locator('#animation-list [aria-selected="true"]').count(), 2);
      await page.locator('[data-edit="animation-preview-selected"]').click();
      await page.waitForFunction(() => window.office.getState().animationPreview);
      assert.equal(
        await page
          .locator('#slide')
          .evaluate((el) => el.shadowRoot.querySelector('.editor-layer').hidden),
        true,
      );
      await page.keyboard.press('Escape');
      assert.equal(await page.evaluate(() => window.office.getState().animationPreview), false);
      assert.equal(
        await page
          .locator('#slide')
          .evaluate((el) => el.shadowRoot.querySelector('.editor-layer').hidden),
        false,
      );
      await page.locator('[data-edit="animation-preview-from"]').click();
      await page.waitForFunction(() => window.office.getState().animationPreview);
      await page.locator('[data-edit="animation-preview-stop"]').click();
      assert.equal(await page.evaluate(() => window.office.getState().animationPreview), false);
      await page.locator('[data-edit="animation-preview-selected"]').click();
      await page.waitForFunction(() => !window.office.getState().animationPreview);
      assert.deepEqual(
        await page.evaluate(() => ({
          revision: window.office.getState().revision,
          history: window.office.getState().history,
        })),
        beforePreview,
      );
      await page.locator('#animation-preview').click();
      await page.evaluate(() => window.office.selectSlide(1));
      assert.equal(await page.evaluate(() => window.office.getState().animationPreview), false);
      await page.evaluate(() => window.office.selectSlide(0));
      await page.locator('#animation-list button').nth(1).click();
      const ribbonBox = await page.locator('#ribbon-animations').boundingBox();
      for (const id of ['animation-start', 'animation-duration', 'animation-delay']) {
        const box = await page.locator('#' + id).boundingBox();
        assert.ok(box.y >= ribbonBox.y && box.y + box.height <= ribbonBox.y + ribbonBox.height);
      }
      const paneBox = await page.locator('#animation-pane').boundingBox();
      assert.ok(stageBox.x + stageBox.width <= paneBox.x + 1);
      await page.screenshot({ path: '/tmp/pptx-animation-pane.png' });
      const earlier = page.locator('[data-edit="animation-earlier"]');
      const later = page.locator('[data-edit="animation-later"]');
      assert.equal(await later.isDisabled(), true);
      await earlier.click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations[0][0].effect === 'fadeOut',
      );
      assert.equal(
        await page
          .locator('#animation-list [aria-selected="true"]')
          .getAttribute('data-animation-id'),
        await page.evaluate(
          () => window.office.getState().editor.slides[0].animations[0][0].timingId,
        ),
      );
      assert.equal(await earlier.isDisabled(), true);
      assert.equal(await markers.first().getAttribute('aria-pressed'), 'true');
      assert.equal(await markers.nth(1).getAttribute('aria-pressed'), 'false');
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations[1][0].effect === 'fadeOut',
      );
      await page.locator('[data-edit="redo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations[0][0].effect === 'fadeOut',
      );
      await later.click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations[1][0].effect === 'fadeOut',
      );
      const delay = page.locator('#animation-delay');
      await delay.fill('1.25');
      await delay.press('Tab');
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations[1][0].delayMs === 1250,
      );
      assert.equal(
        await page.evaluate(
          () => window.office.getState().editor.slides[0].animations[0][0].delayMs,
        ),
        0,
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations[1][0].delayMs === 0,
      );
      await page.locator('[data-edit="redo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations[1][0].delayMs === 1250,
      );
      await delay.fill('0');
      await delay.press('Tab');
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations[1][0].delayMs === 0,
      );
      await page.waitForFunction(() => !document.querySelector('#animation-remove').disabled);
      await markers.nth(1).focus();
      await page.keyboard.press('Delete');
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations.length === 1,
      );
      assert.equal(
        await page.evaluate(() => window.office.getState().editor.slides[0].shapes.length),
        1,
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations.length === 2,
      );
      assert.equal(
        await page.evaluate(
          () => window.office.getState().editor.slides[0].animations[1][0].durationMs,
        ),
        2400,
      );
      await page.locator('[data-edit="redo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations.length === 1,
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations.length === 2,
      );
      await page.locator('#animation-list [role="option"]').nth(1).click();
      // A zero-length exit keeps subsequent navigation assertions deterministic.
      await duration.fill('0');
      await duration.press('Tab');
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations[1][0].durationMs === 0,
      );
      await page.reload();
      await page.waitForFunction(
        () => window.office?.getState().editor?.slides[0]?.animations.length === 2,
      );
      assert.equal(
        await page.evaluate(
          () => window.office.getState().editor.slides[0].animations[1][0].durationMs,
        ),
        0,
      );
      const target = page.locator('#slide [data-pptx-shape-id]').first();
      await page.locator('#present').click();
      await page.waitForFunction(() => window.office.getState().presenting);
      assert.equal(await markers.count(), 0);
      assert.equal(await target.evaluate((el) => getComputedStyle(el).visibility), 'hidden');
      await page.keyboard.press('ArrowRight');
      assert.equal(await target.evaluate((el) => getComputedStyle(el).visibility), 'visible');
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      assert.equal(await page.locator('#present-prev').isDisabled(), false);
      await page.keyboard.press('ArrowRight');
      assert.equal(await target.evaluate((el) => getComputedStyle(el).visibility), 'hidden');
      await page.keyboard.press('ArrowLeft');
      assert.equal(await target.evaluate((el) => getComputedStyle(el).visibility), 'visible');
      await page.keyboard.press('ArrowRight');
      await page.keyboard.press('ArrowRight');
      assert.equal(await page.evaluate(() => window.office.getState().index), 1);
      await page.keyboard.press('Escape');
      await page.evaluate(() => window.office.selectSlide(0));
      assert.equal(await target.evaluate((el) => getComputedStyle(el).visibility), 'visible');
      await page.locator('#tab-animations').click();
      await markers.first().click();
      await start.selectOption('withEffect');
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].animations[0][0].trigger === 'withEffect',
      );
      assert.deepEqual(await markers.allTextContents(), ['0', '1']);
      await page.reload();
      await page.waitForFunction(
        () =>
          window.office?.getState().editor?.slides[0]?.animations[0]?.[0]?.trigger === 'withEffect',
      );
      await page.locator('#present').click();
      await page.waitForFunction(() => window.office.getState().presenting);
      assert.equal(await target.evaluate((el) => getComputedStyle(el).visibility), 'visible');
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      await page.keyboard.press('ArrowRight');
      assert.equal(await target.evaluate((el) => getComputedStyle(el).visibility), 'hidden');
      await page.keyboard.press('Escape');
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      proc.kill('SIGTERM');
      await new Promise((resolve) =>
        proc.exitCode !== null ? resolve() : proc.once('exit', resolve),
      );
      await rm(dir, { recursive: true, force: true });
    }
  },
);

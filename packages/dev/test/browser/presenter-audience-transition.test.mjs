import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

test('audience shares transition layers, timing, resize and cancellation', async () => {
  const scripts = await Promise.all(
    ['transition-playback', 'animation-playback', 'presenter-audience'].map(async (name) => {
      const source = await readFile(new URL('../../src/' + name + '.ts', import.meta.url), 'utf8');
      return source.slice(source.indexOf('`') + 1, source.lastIndexOf('`'));
    }),
  );
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage();
    const errors = [];
    page.on('pageerror', (error) => errors.push(error.message));
    await page.setContent(
      '<div id="slide" style="position:relative;width:640px;height:360px"></div><svg id="ink"></svg><div id="laser" hidden></div><div id="show-screen" hidden></div>',
    );
    await page.evaluate((scripts) => {
      window.eval(`
        const byId=id=>document.getElementById(id),slide=byId('slide'),stage=slide;
        const canvas=slide.attachShadow({mode:'open'}),inkLayer=byId('ink'),laserDot=byId('laser');
        let index=0,presenting=true,screenMode='',showPointerPoint=null,showPointerColor='red',activeInk=null;
        function bindShowInkSurface(){}
        function cancelSlideAdvance(){} function scheduleSlideAdvance(){}
        let displayedSvg='<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 360"><defs><linearGradient id="paint"><stop stop-color="red"/></linearGradient></defs><rect width="640" height="360" fill="url(#paint)"/></svg>';
        canvas.innerHTML=displayedSvg;
        const state={aspectRatio:16/9,editor:{slides:[{animations:[[{effect:'fadeIn',shapeIds:['2'],durationMs:10000,delayMs:0,trigger:'clickEffect'}]]}]}},officeMenu=document.createElement('div');
        ${scripts.join('\n')}
        window.openOutput=openAudience;window.closeOutput=closeAudience;
        window.begin=effect=>playTransition(displayedSvg.replace('red','blue'),{effect,durationMs:10000});
        window.seek=time=>{for(const a of transitionRun.animations){a.pause();a.currentTime=time;}syncAudienceTransition();};
        window.sample=()=>({
          source:transitionRun.animations.map(a=>({time:a.currentTime,frames:a.effect.getKeyframes(),clip:getComputedStyle(a.effect.target).clipPath,transform:getComputedStyle(a.effect.target).transform})),
          target:audienceTransition.animations.map(({animation:a})=>({time:a.currentTime,frames:a.effect.getKeyframes(),clip:audienceWindow.getComputedStyle(a.effect.target).clipPath,transform:audienceWindow.getComputedStyle(a.effect.target).transform})),
          isolated:[...audienceTransition.layer.children].every(node=>node.shadowRoot?.querySelector('#paint')),
          width:audienceTransition.layer.clientWidth,
          scale:audienceTransition.layer.style.transform,
        });
        window.cancel=()=>{cancelTransitionPlayback();syncAudienceTransition();};
        window.beginShape=()=>{cancelTransitionPlayback();canvas.innerHTML='<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 360"><g data-pptx-shape-id="2"><rect width="640" height="360" fill="red"/></g><a href="#focus" tabindex="0"><text x="20" y="20">Focus target</text></a></svg>';initializeShapeAnimations();advanceShapeAnimation();cancelAnimationFrame(shapeAnimationState.frame);renderShapeAnimationState(2500);renderAudience();};
        window.seekShape=time=>{renderShapeAnimationState(time);renderAudience();};
        window.restoreShape=()=>{cancelShapeAnimationPlayback();renderAudience();};
        window.changeArtwork=()=>{canvas.querySelector('rect').setAttribute('fill','blue');renderAudience();};
      `);
    }, scripts);
    const popupReady = page.waitForEvent('popup');
    await page.evaluate(() => window.openOutput());
    const output = await popupReady;
    output.on('pageerror', (error) => errors.push(error.message));
    await output.setViewportSize({ width: 960, height: 540 });
    for (const effect of ['fade', 'push', 'strips', 'checker', 'dissolve']) {
      await page.evaluate((effect) => {
        window.begin(effect);
        window.seek(4000);
      }, effect);
      const sample = await page.evaluate(() => window.sample());
      assert.deepEqual(sample.target, sample.source, effect);
      assert.equal(sample.isolated, true, effect);
      assert.equal(sample.width, 640);
      assert.equal(sample.scale, 'scale(1.5, 1.5)');
      await page.evaluate(() => window.seek(7000));
      const later = await page.evaluate(() => window.sample());
      assert.deepEqual(later.target, later.source, effect + ' seek');
      await page.evaluate(() => window.cancel());
      assert.equal(await output.locator('#audience-transition').count(), 0);
    }
    await page.evaluate(() => {
      window.begin('checker');
      window.seek(5000);
      document.getElementById('slide').style.width = '800px';
    });
    await page.waitForTimeout(100);
    const resized = await page.evaluate(() => window.sample());
    assert.equal(resized.width, 800);
    assert.deepEqual(resized.target, resized.source);
    await output.screenshot({ path: '/tmp/pptx-audience-transition.png' });
    await page.evaluate(() => window.beginShape());
    assert.equal(
      await output.locator('[data-pptx-shape-id="2"]').evaluate((el) => el.style.opacity),
      '0.25',
    );
    await output.locator('a').focus();
    await output.evaluate(() => {
      window.savedLink = document.querySelector('a');
    });
    for (const time of [5000, 7500, 10000]) {
      await page.evaluate((time) => window.seekShape(time), time);
      assert.equal(
        await output.locator('[data-pptx-shape-id="2"]').evaluate((el) => el.style.opacity),
        String(time / 10000),
      );
      assert.equal(
        await output.evaluate(
          () =>
            document.activeElement === window.savedLink &&
            document.querySelector('a') === window.savedLink,
        ),
        true,
      );
    }
    await page.evaluate(() => window.restoreShape());
    assert.equal(await output.locator('[data-pptx-shape-id="2"]').getAttribute('style'), null);
    assert.equal(await output.evaluate(() => document.activeElement === window.savedLink), true);
    await page.evaluate(() => window.changeArtwork());
    assert.equal(await output.locator('rect').getAttribute('fill'), 'blue');

    const closed = output.waitForEvent('close');
    await page.evaluate(() => window.closeOutput());
    await closed;
    assert.equal(output.isClosed(), true);
    assert.deepEqual(errors, []);
  } finally {
    await browser.close();
  }
});

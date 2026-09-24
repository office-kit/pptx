import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test(
  'Presenter view shares navigation, notes, timer and exit lifecycle',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-link-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text,Raw} from '@office-kit/pptx-dsl';import {setSlideTitle} from '@office-kit/pptx';export default <Presentation><Slide><Text x={1} y={1} width={8} height={1}>Link target</Text></Slide><Slide><Text x={1} y={1} width={8} height={1}>Hidden</Text></Slide><Slide layout={{type:"title"}}><Raw scope="slide" apply={({slide})=>setSlideTitle(slide,"Last")}/></Slide></Presentation>`,
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
      const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(url);
      await page.waitForFunction(() => window.office?.getState().editor?.slides.length === 3);
      await page.evaluate(async () => {
        const response = await fetch('/edit', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            revision: window.office.getState().revision,
            command: { type: 'slide-hidden', slide: 1, hidden: true },
          }),
        });
        if (!response.ok) throw new Error(await response.text());
        await window.office.refresh();
      });
      await page.keyboard.press('Alt+Enter');
      await page.waitForFunction(
        () => document.body.classList.contains('presenter-view') && !!document.fullscreenElement,
      );
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      assert.equal(await page.locator('#presenter-next').getAttribute('alt'), 'Next slide: 3');
      assert.equal(await page.locator('#presenter-navigation button').count(), 3);
      assert.equal(await page.locator('#presenter-navigation .hidden-slide').count(), 1);
      const notes = page.getByRole('textbox', { name: 'Slide notes', exact: true });
      await notes.fill('Private presenter notes');
      await page.locator('#presenter-notes-larger').click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.slides[0].notes === 'Private presenter notes' &&
          !window.office.getState().building,
      );
      assert.equal(await notes.evaluate((node) => getComputedStyle(node).fontSize), '24px');
      await page.locator('#presenter-pause').click();
      const elapsed = await page.locator('#presenter-elapsed').textContent();
      await page.waitForTimeout(1100);
      assert.equal(await page.locator('#presenter-elapsed').textContent(), elapsed);
      await page.locator('#presenter-reset').click();
      assert.equal(await page.locator('#presenter-elapsed').textContent(), '0:00');
      await page.getByRole('button', { name: 'Resume timer', exact: true }).click();
      await page.waitForFunction(
        () => document.getElementById('presenter-elapsed').textContent !== '0:00',
      );
      const before = await page.locator('#stage').boundingBox();
      await page.locator('#presenter-divider').focus();
      await page.keyboard.press('ArrowLeft');
      assert.ok((await page.locator('#stage').boundingBox()).width < before.width);
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      await page.screenshot({ path: '/tmp/pptx-presenter-view.png' });
      await page
        .locator('#presenter-navigation')
        .getByRole('button', { name: 'Slide 2 (hidden)', exact: true })
        .click();
      assert.equal(await page.evaluate(() => window.office.getState().index), 1);
      assert.equal(await notes.inputValue(), '');
      assert.equal(await page.locator('#presenter-next').getAttribute('alt'), 'Next slide: 3');
      await page.keyboard.press('ArrowRight');
      assert.equal(await page.evaluate(() => window.office.getState().index), 2);
      assert.equal(await page.locator('#presenter-next').isVisible(), false);
      assert.equal(await page.locator('#presenter-next-label').textContent(), 'End of slide show');
      await page.keyboard.press('b');
      assert.equal(await page.locator('#show-screen').isVisible(), true);
      assert.equal(await page.locator('#presenter-navigation').isVisible(), true);
      await page
        .locator('#presenter-navigation')
        .getByRole('button', { name: 'Slide 1', exact: true })
        .click();
      assert.equal(await page.locator('#show-screen').isVisible(), false);
      assert.equal(await notes.inputValue(), 'Private presenter notes');
      await page.getByRole('button', { name: 'Use Slide Show', exact: true }).focus();
      await page.keyboard.press('Enter');
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      assert.equal(await page.evaluate(() => window.office.getState().presenting), true);
      assert.equal(await page.locator('#presenter-view').isVisible(), false);
      await page.keyboard.press('Escape');
      await page.waitForFunction(
        () => !window.office.getState().presenting && !document.fullscreenElement,
      );
      assert.equal(await page.locator('#notes-pane #speaker-notes').count(), 1);
      await page.locator('#tab-show').click();
      await page.evaluate(() =>
        Object.defineProperty(window.screen, 'isExtended', { value: true, configurable: true }),
      );
      // Simulate ScreenDetailed permission/placement; physical monitor behavior is a separate check.
      await page.evaluate(() => {
        const first = { availLeft: 0, availTop: 0, availWidth: 1600, availHeight: 1000 };
        const second = { availLeft: 1600, availTop: 0, availWidth: 1920, availHeight: 1080 };
        const details = Object.assign(new EventTarget(), {
          screens: [first, second],
          currentScreen: first,
        });
        window.displayFixture = { details, requests: [], reject: false };
        window.getScreenDetails = async () => details;
        document.documentElement.requestFullscreen = async (options) => {
          if (!options?.screen) throw new Error('Activation unavailable');
          window.displayFixture.requests.push(details.screens.indexOf(options.screen));
          if (window.displayFixture.reject) throw new Error('Display permission denied');
          if (!window.displayFixture.ignore) details.currentScreen = options.screen;
        };
      });
      const outputReady = page.waitForEvent('popup');
      await page.locator('#presenter-view-start').click();
      const audience = await outputReady;
      audience.on('pageerror', (error) => errors.push(error.message));
      await audience.waitForFunction(() => document.querySelector('#audience-artwork svg'));
      assert.equal(await audience.locator('textarea').count(), 0);
      assert.equal(
        (await audience.locator('body').textContent()).includes('Private presenter notes'),
        false,
      );
      await audience.evaluate(() => {
        window.moves = [];
        window.moveTo = (x, y) => window.moves.push(['move', x, y]);
        window.resizeTo = (w, h) => window.moves.push(['resize', w, h]);
      });
      const swap = page.getByRole('button', { name: 'Swap Displays', exact: true });
      await swap.click();
      await page.waitForFunction(() => window.displayFixture.requests.length === 1);
      assert.deepEqual(await audience.evaluate(() => window.moves), [
        ['move', 0, 0],
        ['resize', 1600, 1000],
      ]);
      await swap.click();
      assert.deepEqual(await page.evaluate(() => window.displayFixture.requests), [1, 0]);
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      assert.equal(await notes.inputValue(), 'Private presenter notes');
      await page.evaluate(() => {
        window.displayFixture.ignore = true;
      });
      const movesBeforeIgnored = await audience.evaluate(() => window.moves.length);
      await swap.click();
      await page.waitForFunction(() =>
        document.getElementById('presenter-display-status').textContent.includes('did not move'),
      );
      assert.equal(await audience.evaluate(() => window.moves.length), movesBeforeIgnored);
      await page.evaluate(() => {
        window.displayFixture.ignore = false;
      });
      await page.evaluate(() => {
        window.displayFixture.reject = true;
      });
      const movesBeforeFailure = await audience.evaluate(() => window.moves.length);
      await swap.click();
      await page.waitForFunction(
        () =>
          document.getElementById('presenter-display-status').textContent ===
          'Display permission denied',
      );
      assert.equal(await audience.evaluate(() => window.moves.length), movesBeforeFailure);
      assert.equal(await swap.isEnabled(), true);
      await page.evaluate(() => {
        window.displayFixture.details.screens = [window.displayFixture.details.currentScreen];
        window.displayFixture.details.dispatchEvent(new Event('screenschange'));
      });
      assert.equal(await swap.isVisible(), false);
      await page.locator('#presenter-advance').click();
      await audience.waitForFunction(() =>
        document.querySelector('#audience-artwork').textContent.includes('Last'),
      );
      await page.locator('#stage').focus();
      await page.keyboard.press('b');
      await audience.waitForFunction(() => !document.getElementById('audience-screen').hidden);
      assert.equal(
        await audience
          .locator('#audience-screen')
          .evaluate((node) => getComputedStyle(node).backgroundColor),
        'rgb(0, 0, 0)',
      );
      await page.keyboard.press('w');
      await audience.waitForFunction(
        () =>
          getComputedStyle(document.getElementById('audience-screen')).backgroundColor ===
          'rgb(255, 255, 255)',
      );
      await page.keyboard.press('w');
      await page.keyboard.press('Meta+p');
      const slideBox = await page.locator('#slide').boundingBox();
      await page.mouse.move(slideBox.x + slideBox.width * 0.3, slideBox.y + slideBox.height * 0.3);
      await page.mouse.down();
      await page.mouse.move(slideBox.x + slideBox.width * 0.6, slideBox.y + slideBox.height * 0.6, {
        steps: 4,
      });
      await page.mouse.up();
      await audience.waitForFunction(
        () => document.querySelectorAll('#audience-ink path').length === 1,
      );
      assert.equal(
        await audience.locator('#audience-ink path').getAttribute('d'),
        await page.locator('#show-ink path').getAttribute('d'),
      );
      await page.keyboard.press('b');
      await audience.waitForFunction(() => !document.getElementById('audience-screen').hidden);
      assert.equal(
        await audience.evaluate(
          () => document.elementFromPoint(innerWidth / 2, innerHeight / 2).id,
        ),
        'audience-screen',
      );
      await page.keyboard.press('b');
      await page.keyboard.press('Shift+e');
      await audience.waitForFunction(() => !document.querySelector('#audience-ink path'));
      await audience.keyboard.press('ArrowLeft');
      await audience.waitForFunction(() =>
        document.querySelector('#audience-artwork').textContent.includes('Link target'),
      );

      await page.waitForFunction(
        () =>
          document.body.classList.contains('presenter-view') && window.office.getState().presenting,
      );
      await audience
        .locator('#audience-slide')
        .click({ button: 'right', position: { x: 80, y: 80 } });
      await audience.getByRole('menuitem', { name: 'By Title', exact: true }).hover();
      const titleMenu = audience.getByRole('menu', { name: 'By Title', exact: true });
      assert.deepEqual(await titleMenu.getByRole('menuitem').allTextContents(), [
        '1 Slide 1',
        '2 Slide 2',
        '3 Last',
      ]);
      await titleMenu.getByRole('menuitem', { name: '3 Last', exact: true }).click();
      await page.waitForFunction(() => window.office.getState().index === 2);
      await audience.keyboard.press('ArrowLeft');
      await page.waitForFunction(() => window.office.getState().index === 0);
      await audience
        .locator('#audience-slide')
        .click({ button: 'right', position: { x: 80, y: 80 } });
      await audience.screenshot({ path: '/tmp/pptx-audience-context-menu.png' });
      await audience.getByRole('menuitem', { name: 'Pointer Options', exact: true }).hover();
      await audience.getByRole('menuitemcheckbox', { name: 'Arrow', exact: false }).click();
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      assert.equal(await audience.locator('#office-menu').count(), 0);
      await audience.keyboard.press('Shift+F10');
      await audience.getByRole('menuitem', { name: 'Screen', exact: true }).focus();
      await audience.keyboard.press('ArrowRight');
      await audience.keyboard.press('Enter');
      await audience.waitForFunction(() => !document.getElementById('audience-screen').hidden);
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      await audience.locator('#audience-screen').click({ button: 'right' });
      await audience.getByRole('menuitem', { name: 'Next', exact: true }).click();
      await audience.waitForFunction(() => document.getElementById('audience-screen').hidden);
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      await audience.keyboard.press('Shift+F10');
      await audience.keyboard.press('Escape');
      assert.equal(await audience.locator('#office-menu').count(), 0);
      assert.equal(await audience.evaluate(() => document.activeElement.id), 'audience-slide');
      await page.locator('#stage').focus();
      await page.keyboard.press('Meta+a');
      const insertAudienceAction = async (href) =>
        page.evaluate((href) => {
          const svg = document.getElementById('slide').shadowRoot.querySelector('svg');
          svg.querySelector('#audience-action-fixture')?.remove();
          const action = document.createElementNS(svg.namespaceURI, href ? 'a' : 'g');
          if (href?.startsWith('https:')) action.setAttribute('target', '_blank');
          if (href) action.setAttribute('href', href);
          else action.setAttribute('data-click-stop-sound', 'true');
          action.id = 'audience-action-fixture';
          const rect = document.createElementNS(svg.namespaceURI, 'rect');
          rect.setAttribute('width', String(svg.viewBox.baseVal.width));
          rect.setAttribute('height', String(svg.viewBox.baseVal.height));
          rect.setAttribute('fill', 'transparent');
          action.append(rect);
          svg.append(action);
        }, href);
      await insertAudienceAction('#slide-lastSlide');
      await audience.locator('#audience-action-fixture').click();
      await page.waitForFunction(() => window.office.getState().index === 2);
      await page.locator('#presenter-previous').click();
      await insertAudienceAction('#slide-lastSlide');
      await audience.keyboard.press('Tab');
      await audience.waitForFunction(
        () => document.activeElement?.getAttribute('href') === '#slide-lastSlide',
      );
      await audience.keyboard.press('Enter');
      await page.waitForFunction(() => window.office.getState().index === 2);
      await page.locator('#presenter-previous').click();
      await insertAudienceAction(null);
      await audience.locator('#audience-action-fixture').click();
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      await page
        .context()
        .route('https://audience-link.test/', (route) =>
          route.fulfill({ status: 200, body: 'Audience link target' }),
        );
      await insertAudienceAction('https://audience-link.test/');
      const linkedPageReady = page.context().waitForEvent('page');
      await audience.locator('#audience-action-fixture').click();
      const linkedPage = await linkedPageReady;
      await linkedPage.waitForURL('https://audience-link.test/');
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      assert.equal(audience.url(), 'about:blank');
      await linkedPage.close();
      await audience.keyboard.press('Meta+l');
      const audienceBox = await audience.locator('#audience-slide').boundingBox();
      await audience.mouse.move(
        audienceBox.x + audienceBox.width * 0.25,
        audienceBox.y + audienceBox.height * 0.75,
      );
      await audience.waitForFunction(() => !document.getElementById('audience-laser').hidden);
      const laserPosition = await audience
        .locator('#audience-laser')
        .evaluate((node) => ({ x: parseFloat(node.style.left), y: parseFloat(node.style.top) }));
      assert.ok(Math.abs(laserPosition.x - 25) < 0.5);
      assert.ok(Math.abs(laserPosition.y - 75) < 0.5);
      assert.equal(
        await audience
          .locator('#audience-artwork svg')
          .evaluate((node) => getComputedStyle(node).cursor),
        'none',
      );
      await audience.keyboard.press('b');
      assert.equal(await audience.locator('#audience-laser').isVisible(), false);
      await audience.keyboard.press('b');
      await audience.evaluate(() =>
        document.dispatchEvent(new PointerEvent('pointerout', { relatedTarget: null })),
      );
      assert.equal(await audience.locator('#audience-laser').isVisible(), false);
      await audience.keyboard.press('Meta+a');
      assert.equal(
        await audience
          .locator('#audience-artwork svg')
          .evaluate((node) => getComputedStyle(node).cursor),
        'default',
      );
      await audience.mouse.move(-10, -10);
      await insertAudienceAction(null);
      await page.evaluate(() => {
        const action = document
          .getElementById('slide')
          .shadowRoot.querySelector('#audience-action-fixture');
        action.removeAttribute('data-click-stop-sound');
        action.setAttribute('data-hover-href', '#slide-lastSlide');
      });
      await audience.waitForFunction(() => document.querySelector('[data-hover-href]'));
      const hoverAction = (pointerType = 'mouse') =>
        audience
          .locator('#audience-action-fixture rect')
          .dispatchEvent('pointerover', { pointerType, clientX: 123, clientY: 145 });
      await hoverAction('touch');
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      await audience.keyboard.press('Meta+p');
      await hoverAction();
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      await audience.keyboard.press('Meta+a');
      await hoverAction();
      await page.waitForFunction(() => window.office.getState().index === 2);
      await audience.keyboard.press('Meta+p');
      const inkBox = await audience.locator('#audience-slide').boundingBox();
      await audience.mouse.move(inkBox.x + inkBox.width * 0.2, inkBox.y + inkBox.height * 0.3);
      await audience.mouse.down();
      await audience.mouse.move(inkBox.x + inkBox.width * 0.7, inkBox.y + inkBox.height * 0.6, {
        steps: 5,
      });
      await audience.mouse.up();
      await page.waitForFunction(() => document.querySelectorAll('#show-ink path').length === 1);
      const audienceInk = await audience.locator('#audience-ink path').getAttribute('d');
      assert.equal(audienceInk, await page.locator('#show-ink path').getAttribute('d'));
      const firstPoint = audienceInk.match(/^M(\d+) (\d+)/);
      const dimensions = await page.evaluate(() => ({
        width: window.office.getState().editor.width,
        height: window.office.getState().editor.height,
      }));
      assert.ok(Math.abs(Number(firstPoint[1]) / dimensions.width - 0.2) < 0.005);
      assert.ok(Math.abs(Number(firstPoint[2]) / dimensions.height - 0.3) < 0.005);
      assert.equal(await page.evaluate(() => window.office.getState().index), 2);
      await audience.keyboard.press('ArrowLeft');
      await page.waitForFunction(() => window.office.getState().index === 0);
      await audience.waitForFunction(() => !document.querySelector('#audience-ink path'));
      await audience.keyboard.press('ArrowRight');
      await page.waitForFunction(() => window.office.getState().index === 2);
      await audience.waitForFunction(() => document.querySelector('#audience-ink path'));
      assert.equal(await audience.locator('#audience-ink path').getAttribute('d'), audienceInk);
      await audience.keyboard.press('Shift+F10');
      await audience.getByRole('menuitem', { name: 'End Show', exact: true }).click();
      await page.waitForFunction(
        () => !window.office.getState().presenting && !document.fullscreenElement,
      );
      assert.equal(await page.locator('#presenter-view').isVisible(), false);
      assert.equal(audience.isClosed(), true);
      assert.equal(await page.locator('#office-menu').count(), 1);
      await page
        .getByRole('dialog', { name: 'Keep ink annotations' })
        .getByRole('button', { name: 'Keep', exact: true })
        .click();
      await page.waitForFunction(() =>
        window.office.getState().editor.slides[2].shapes.some((shape) => shape.kind === 'ink'),
      );
      await page.reload();
      await page.waitForFunction(
        () => window.office?.getState().editor?.slides[0]?.notes === 'Private presenter notes',
      );
      await page.waitForFunction(() =>
        window.office.getState().editor.slides[2].shapes.some((shape) => shape.kind === 'ink'),
      );
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

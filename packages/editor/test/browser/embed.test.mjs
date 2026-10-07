// Embeds the built package (dist/) the way a host application does: a plain HTML
// page with hostile global CSS and no framework, driving the editor only through
// `mountEditor` and its handle. Run `pnpm --filter @office-kit/pptx-editor... build` first.
import assert from 'node:assert/strict';
import { after, before, test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';
import { chromium } from 'playwright';
import {
  addTitleSlide,
  createPresentation,
  getShapeBoundsResolved,
  getSlideShapes,
  getSlideText,
  getSlideXmlString,
  getSlides,
  loadPresentation,
  savePresentation,
} from '@office-kit/pptx';
import { ja } from '../../src/i18n/ja.ts';

const ORIGIN = 'http://editor.test';

// Rules a careless host stylesheet might carry: none may reach the editor, and
// the editor's own class names must not restyle the page.
const HOST_CSS = `
  * { box-sizing: content-box; }
  body { margin: 0; font: 31px/3 serif; letter-spacing: 3px; text-transform: uppercase; color: rgb(255, 0, 0); }
  div, span, button, section, input, p { padding: 9px; border: 3px solid rgb(0, 255, 0); background: rgb(255, 255, 0); font-size: 40px; }
  .ok-shell, .ribbon, .ok-editor, .ok-btn { display: none; }
  .host-target { height: 820px; padding: 0; border: 0; }
`;
const PAGE = `<!doctype html><html lang="en"><meta charset="utf-8"><title>Host</title>
<style>${HOST_CSS}</style>
<p class="ok-btn" id="host-paragraph">Host content</p>
<div class="host-target" id="a"></div>
<div class="host-target" id="b"></div>
<script type="module" src="/editor.js"></script></html>`;

let browser;
let script;

before(async () => {
  const bundle = await build({
    stdin: {
      contents: `import { mountEditor } from ${JSON.stringify(
        fileURLToPath(new URL('../../dist/index.js', import.meta.url)),
      )}; window.mountEditor = mountEditor;`,
      resolveDir: fileURLToPath(new URL('../..', import.meta.url)),
    },
    bundle: true,
    format: 'esm',
    platform: 'browser',
    target: 'es2022',
    conditions: ['browser'],
    write: false,
  });
  script = bundle.outputFiles[0].text;
  browser = await chromium.launch({
    channel: process.env.PLAYWRIGHT_CHANNEL || undefined,
    headless: true,
  });
});

after(async () => {
  await browser?.close();
});

async function deck(title) {
  const pres = createPresentation();
  addTitleSlide(pres, title);
  return [...(await savePresentation(pres))];
}

/** A host page that counts the window and document listeners still registered. */
async function hostPage() {
  const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
  const errors = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.addInitScript(() => {
    const live = new Set();
    const key = (target, type, listener, options) => [
      target === window ? 'window' : 'document',
      type,
      listener,
      !!(typeof options === 'boolean' ? options : options?.capture),
    ];
    const add = EventTarget.prototype.addEventListener;
    const remove = EventTarget.prototype.removeEventListener;
    const ids = new WeakMap();
    let next = 0;
    const id = (listener) => {
      if (!ids.has(listener)) ids.set(listener, ++next);
      return ids.get(listener);
    };
    EventTarget.prototype.addEventListener = function (type, listener, options) {
      if ((this === window || this === document) && listener) {
        const [t, ty, l, c] = key(this, type, listener, options);
        live.add(`${t}:${ty}:${id(l)}:${c}`);
      }
      return add.call(this, type, listener, options);
    };
    EventTarget.prototype.removeEventListener = function (type, listener, options) {
      if ((this === window || this === document) && listener) {
        const [t, ty, l, c] = key(this, type, listener, options);
        live.delete(`${t}:${ty}:${id(l)}:${c}`);
      }
      return remove.call(this, type, listener, options);
    };
    window.liveListeners = () => [...live].sort();
  });
  await page.route(`${ORIGIN}/**`, (route) => {
    const path = new URL(route.request().url()).pathname;
    if (path === '/editor.js')
      return route.fulfill({ contentType: 'text/javascript', body: script });
    return route.fulfill({ contentType: 'text/html', body: PAGE });
  });
  await page.goto(ORIGIN + '/');
  await page.waitForFunction(() => typeof window.mountEditor === 'function');
  return { page, errors };
}

/** Mounts an editor in `#id` with `bytes`, recording what onSave receives. */
async function mount(page, id, bytes, options = {}) {
  await page.evaluate(
    async ({ id, bytes, options }) => {
      window.saved ??= {};
      window.handles ??= {};
      const handle = window.mountEditor(document.getElementById(id), {
        source: new Uint8Array(bytes),
        ...options,
        onSave: (pptx) => {
          window.saved[id] = [...pptx];
        },
      });
      window.handles[id] = handle;
      await handle.ready;
    },
    { id, bytes, options },
  );
  return page.locator(`#${id}`);
}

const savedDeck = async (page, id) =>
  loadPresentation(new Uint8Array(await page.evaluate((id) => window.saved[id], id)));
const handleDeck = async (page, id) =>
  loadPresentation(
    new Uint8Array(await page.evaluate(async (id) => [...(await window.handles[id].save())], id)),
  );
const titleOf = (presentation, index = 0) => getSlideText(getSlides(presentation)[index]);

test('isolates styles in a shadow root and saves edits through onSave and save()', async () => {
  const { page, errors } = await hostPage();
  const headStyles = await page.locator('head style').count();
  const editor = await mount(page, 'a', await deck('Embedded title'), { locale: 'en' });

  // Neither the host's rules nor its inherited font reach the editor …
  const styles = await page.evaluate(() => {
    const root = document.querySelector('#a > office-kit-pptx-editor').shadowRoot;
    const shell = getComputedStyle(root.querySelector('.ok-shell'));
    const button = getComputedStyle(root.querySelector('.ok-btn'));
    return {
      shell: [
        shell.display,
        shell.fontSize,
        shell.textTransform,
        shell.letterSpacing,
        shell.lineHeight,
      ],
      font: shell.fontFamily,
      button: [button.borderTopColor, button.paddingTop, button.boxSizing],
      height: root.querySelector('.ok-shell').getBoundingClientRect().height,
      host: getComputedStyle(document.getElementById('host-paragraph')).display,
    };
  });
  assert.deepEqual(styles.shell, ['grid', '13px', 'none', 'normal', 'normal']);
  assert.match(styles.font, /^system-ui/);
  assert.notEqual(styles.button[0], 'rgb(0, 255, 0)');
  assert.notEqual(styles.button[1], '9px');
  assert.equal(styles.button[2], 'border-box');
  // … the editor fills its target …
  assert.equal(styles.height, 820);
  // … and the editor's class names do not restyle the page.
  assert.equal(styles.host, 'none');
  assert.equal(await page.locator('head style').count(), headStyles);

  // A ribbon menu opens and takes a click inside the shadow root.
  await editor.getByRole('button', { name: 'New Slide options', exact: true }).click();
  await editor.getByRole('menu', { name: 'New Slide', exact: true }).waitFor();
  await editor.getByRole('menuitem', { name: 'Title Slide', exact: true }).click();
  await editor.getByRole('menu', { name: 'New Slide', exact: true }).waitFor({ state: 'hidden' });

  // Edit the first slide's title with the keyboard: the caret and a selection
  // made inside the shadow root decide what is replaced.
  await editor.locator('.thumb-row').first().click();
  await editor.locator('.hit').first().dblclick();
  const input = editor.locator('.canvas-shell .inline-edit');
  await input.waitFor();
  // Arrow keys rather than End: at this zoom the title wraps after "Embedded".
  for (let i = 0; i < 'Embedded title'.length; i++) await page.keyboard.press('ArrowRight');
  for (let i = 0; i < 'title'.length; i++) await page.keyboard.press('Shift+ArrowLeft');
  await page.keyboard.type('headline');
  await page.keyboard.press('Control+Enter');
  await editor.locator('.canvas-shell .inline-edit').waitFor({ state: 'hidden' });

  // Drag the title: pointer capture and the window's pointer listeners work
  // through the shadow boundary.
  const box = await editor.locator('.hit').first().boundingBox();
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.down();
  await page.mouse.move(box.x + box.width / 2 + 60, box.y + box.height / 2 + 30, { steps: 5 });
  await page.mouse.up();

  // A context menu and a modal dialog (top layer) work from the shadow root.
  await editor.locator('.hit').first().click({ button: 'right' });
  await editor.getByRole('menuitem', { name: 'Hyperlink...', exact: true }).click();
  const linkDialog = editor.getByRole('dialog', { name: 'Edit link', exact: true });
  await linkDialog.getByLabel('Link address', { exact: true }).fill('https://example.com/embedded');
  await linkDialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await linkDialog.waitFor({ state: 'hidden' });

  await editor.getByRole('button', { name: 'Save', exact: true }).click();
  await page.waitForFunction(() => window.saved?.a);
  const saved = await savedDeck(page, 'a');
  assert.equal(getSlides(saved).length, 2);
  assert.equal(titleOf(saved), 'Embedded headline');
  assert.match(getSlideXmlString(getSlides(saved)[0]), /<a:hlinkClick /);
  const original = await loadPresentation(new Uint8Array(await deck('Embedded title')));
  const bounds = (presentation) =>
    getShapeBoundsResolved(presentation, getSlideShapes(getSlides(presentation)[0])[0]);
  assert.ok(bounds(saved).x > bounds(original).x);
  assert.ok(bounds(saved).y > bounds(original).y);

  // save() serializes the same state without calling onSave again.
  await page.evaluate(() => (window.saved.a = null));
  const pulled = await handleDeck(page, 'a');
  assert.equal(titleOf(pulled), 'Embedded headline');
  assert.equal(getSlides(pulled).length, 2);
  assert.equal(await page.evaluate(() => window.saved.a), null);

  // Ctrl/Cmd+S saves too.
  await editor.locator('.hit').first().click();
  await page.keyboard.press('ControlOrMeta+s');
  await page.waitForFunction(() => window.saved?.a);
  assert.deepEqual(errors, []);
  await page.close();
});

test('destroy() removes the editor and every listener it added', async () => {
  const { page, errors } = await hostPage();
  const baseline = await page.evaluate(() => ({
    listeners: window.liveListeners(),
    globals: Object.keys(window).sort(),
    head: document.head.innerHTML,
    body: document.body.children.length,
  }));
  const editor = await mount(page, 'a', await deck('Destroyed'));
  // Use it first, so menus, dialogs and text editing have registered theirs.
  await editor.getByRole('button', { name: 'New Slide options', exact: true }).click();
  await editor.getByRole('menuitem', { name: 'Title Slide', exact: true }).click();
  await editor.locator('.hit').first().dblclick();
  await editor.locator('.canvas-shell .inline-edit').waitFor();
  await page.keyboard.type('x');
  await page.keyboard.press('Escape');
  await page.evaluate(() => {
    window.handles.a.destroy();
    // Destroying twice is harmless.
    window.handles.a.destroy();
  });
  const leftover = await page.evaluate(() => ({
    // Svelte's runtime registers one page-wide form `reset` listener the first
    // time any component binds an input; it is shared, not the editor's own.
    listeners: window.liveListeners().filter((entry) => !entry.startsWith('document:reset:')),
    globals: Object.keys(window)
      .filter((name) => name !== 'saved' && name !== 'handles')
      .sort(),
    head: document.head.innerHTML,
    body: document.body.children.length,
    target: document.getElementById('a').childNodes.length,
  }));
  assert.deepEqual(leftover, { ...baseline, target: 0 });

  // The target can take a new editor.
  const again = await mount(page, 'a', await deck('Mounted again'));
  await again.locator('.hit').first().waitFor();
  assert.equal(titleOf(await handleDeck(page, 'a')), 'Mounted again');
  assert.deepEqual(errors, []);
  await page.close();
});

test('two editors on one page keep their own documents, selection and keys', async () => {
  const { page, errors } = await hostPage();
  const first = await mount(page, 'a', await deck('First deck'));
  const second = await mount(page, 'b', await deck('Second deck'));

  const shapes = await first.locator('.hit').count();
  assert.equal(await second.locator('.hit').count(), shapes);
  // Clicking the title edits its text; Escape leaves the shape selected and
  // focus on the page, where keys go to the editor used last.
  const select = async (editor) => {
    await editor.locator('.hit').first().click();
    await page.keyboard.press('Escape');
  };

  // Delete in the first editor removes its selected shape only.
  await select(first);
  await page.keyboard.press('Delete');
  await page.waitForFunction(
    (count) =>
      document.querySelector('#a > office-kit-pptx-editor').shadowRoot.querySelectorAll('.hit')
        .length === count,
    shapes - 1,
  );
  assert.equal(await second.locator('.hit').count(), shapes);

  // Undo after using the second editor acts on the second, not the first.
  await select(second);
  await page.keyboard.press('ControlOrMeta+z');
  assert.equal(await first.locator('.hit').count(), shapes - 1);

  // Keys typed into the page's own controls reach neither editor.
  await page.evaluate(() => {
    const input = document.createElement('input');
    input.id = 'host-input';
    document.body.prepend(input);
  });
  await page.locator('#host-input').click();
  await page.keyboard.press('Delete');
  assert.equal(await second.locator('.hit').count(), shapes);

  await second.getByRole('button', { name: 'Save', exact: true }).click();
  await page.waitForFunction(() => window.saved?.b);
  assert.equal(await page.evaluate(() => window.saved.a), undefined);
  const secondDeck = await savedDeck(page, 'b');
  assert.equal(titleOf(secondDeck), 'Second deck');
  assert.equal(getSlideShapes(getSlides(secondDeck)[0]).length, shapes);
  assert.equal(getSlideShapes(getSlides(await handleDeck(page, 'a'))[0]).length, shapes - 1);
  assert.deepEqual(errors, []);
  await page.close();
});

test('shows the Japanese interface and saves through it', async () => {
  const { page, errors } = await hostPage();
  const editor = await mount(page, 'a', await deck('日本語のタイトル'), { locale: 'ja' });
  await editor.getByRole('tab', { name: ja.Home, exact: true }).waitFor();
  await editor.getByRole('button', { name: ja['New Slide options'], exact: true }).click();
  await editor.getByRole('menuitem', { name: ja['Title Slide'], exact: true }).click();
  await editor.getByRole('button', { name: ja.Save, exact: true }).click();
  await page.waitForFunction(() => window.saved?.a);
  const saved = await savedDeck(page, 'a');
  assert.equal(getSlides(saved).length, 2);
  assert.equal(titleOf(saved), '日本語のタイトル');
  // The host's choice is not remembered as the user's own.
  assert.equal(await page.evaluate(() => localStorage.getItem('ok-editor-locale')), null);
  assert.deepEqual(errors, []);
  await page.close();
});

test('a failed onSave is shown, and an unreadable source rejects ready', async () => {
  const { page } = await hostPage();
  const editor = await mount(page, 'a', await deck('Failing save'), { locale: 'en' });
  await page.evaluate(() => {
    window.handles.a.destroy();
    window.handles.a = null;
  });
  await page.evaluate(
    async (bytes) => {
      window.handles.a = window.mountEditor(document.getElementById('a'), {
        source: new Uint8Array(bytes),
        onSave: () => Promise.reject(new Error('Storage is full')),
      });
      await window.handles.a.ready;
    },
    await deck('Failing save'),
  );
  await editor.getByRole('button', { name: 'Save', exact: true }).click();
  await editor.getByText('Save failed: Storage is full').waitFor();

  const rejected = await page.evaluate(async () => {
    const target = document.getElementById('b');
    const handle = window.mountEditor(target, { source: new Uint8Array([1, 2, 3]) });
    const message = await handle.ready.then(
      () => 'resolved',
      (error) => error.message,
    );
    const saved = await handle.save().then(
      () => 'resolved',
      () => 'rejected',
    );
    handle.destroy();
    return { message, saved, children: target.childNodes.length };
  });
  assert.notEqual(rejected.message, 'resolved');
  assert.equal(rejected.saved, 'rejected');
  assert.equal(rejected.children, 0);
  await page.close();
});

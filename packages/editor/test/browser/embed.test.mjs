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
  findShapeById,
  getShapeBoundsResolved,
  getShapeFillColor,
  getShapeName,
  getShapePosition,
  getShapeText,
  getSlidePartName,
  getSlideShapes,
  getSlideText,
  getSlideXmlString,
  getSlides,
  isSlideHidden,
  loadPresentation,
  savePresentation,
  setShapeText,
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
      // The page's "agent" uses the core library the editor depends on.
      contents: `import { mountEditor, resolveShape } from ${JSON.stringify(
        fileURLToPath(new URL('../../dist/index.js', import.meta.url)),
      )}; import * as pptx from '@office-kit/pptx';
      Object.assign(window, { mountEditor, resolveShape, pptx });`,
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
    new Uint8Array(
      await page.evaluate(async (id) => [...(await window.handles[id].snapshot())], id),
    ),
  );
const titleOf = (presentation, index = 0) => getSlideText(getSlides(presentation)[index]);

test('isolates styles in a shadow root and saves edits through onSave and snapshot()', async () => {
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

  // snapshot() serializes the same state without calling onSave again.
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
    const saved = await handle.snapshot().then(
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

// --- Agents in the browser ---------------------------------------------------
// An "agent" here is a script in the host page: it reads the deck and the user's
// selection, and edits through `apply` with the core library's public API.

/** Records the handle's events for `id` in `window.events[id]`. */
const watchEvents = (page, id) =>
  page.evaluate((id) => {
    window.events ??= {};
    const events = (window.events[id] = { change: [], selectionchange: [] });
    const handle = window.handles[id];
    handle.on('change', (event) => events.change.push(event.source));
    handle.on('selectionchange', (event) => events.selectionchange.push(event.selection));
  }, id);
const events = (page, id) => page.evaluate((id) => window.events[id], id);
const waitForChanges = (page, id, count) =>
  page.waitForFunction(({ id, count }) => window.events[id].change.length === count, { id, count });
const firstShape = (presentation) => getSlideShapes(getSlides(presentation)[0])[0];

/**
 * Runs `handle.apply` in the page with `edit` as the body of a function of
 * (presentation, pptx, resolveShape, arg); resolves to its rejection message, or null.
 */
const apply = (page, id, label, edit, arg) =>
  page.evaluate(
    async ({ id, label, edit, arg }) => {
      const run = new Function('presentation', 'pptx', 'resolveShape', 'arg', edit);
      return window.handles[id]
        .apply(label, (presentation) => run(presentation, window.pptx, window.resolveShape, arg))
        .then(
          () => null,
          (error) => error.message,
        );
    },
    { id, label, edit, arg },
  );

/** The Edit menu's Undo item, which names the edit Undo reverts. */
async function undoMenuItem(editor, page, labels = { edit: 'Edit', undo: 'Undo' }) {
  await editor
    .getByRole('menubar')
    .getByRole('menuitem', { name: labels.edit, exact: true })
    .click();
  const item = editor
    .getByRole('menu', { name: labels.edit, exact: true })
    .getByRole('menuitem', { name: new RegExp(`^${labels.undo}`) })
    .first();
  const name = await item.getAttribute('aria-label');
  await page.keyboard.press('Escape');
  return name;
}

test('an agent reads the deck and selection, and edits as one undo step', async () => {
  const { page, errors } = await hostPage();
  const editor = await mount(page, 'a', await deck('Quarterly review'), { locale: 'en' });
  await watchEvents(page, 'a');
  const original = await handleDeck(page, 'a');
  assert.equal(titleOf(original), 'Quarterly review');
  assert.deepEqual(await page.evaluate(() => window.handles.a.selection()), []);

  // The user picks the title; the agent sees which shape "this" is.
  await editor.locator('.hit').first().click();
  await page.keyboard.press('Escape');
  const selected = await page.evaluate(() => window.handles.a.selection());
  const ref = {
    slideIndex: 0,
    slide: getSlidePartName(getSlides(original)[0]),
    shapeId: selected[0]?.shapeId,
    name: getShapeName(firstShape(original)),
  };
  assert.deepEqual(selected, [ref]);
  await page.waitForFunction(() => window.events.a.selectionchange.length > 0);
  assert.deepEqual((await events(page, 'a')).selectionchange.at(-1), [ref]);
  const changesBefore = (await events(page, 'a')).change.length;

  const edit = `const shape = resolveShape(presentation, arg);
    pptx.setShapeText(shape, 'Agent title');
    pptx.setShapeFill(shape, '#FF0000');`;
  assert.equal(await apply(page, 'a', 'Make the title red', edit, ref), null);
  assert.deepEqual((await events(page, 'a')).change.slice(changesBefore), ['agent']);
  const edited = await handleDeck(page, 'a');
  assert.equal(getShapeText(firstShape(edited)), 'Agent title');
  assert.equal(getShapeFillColor(firstShape(edited)), '#FF0000');
  // The canvas shows it at once.
  assert.match(await editor.locator('.canvas-shell .paint').textContent(), /Agent/);
  assert.equal(await undoMenuItem(editor, page), 'Undo Agent: Make the title red');

  // One Undo reverts both the text and the fill; Redo restores both.
  await page.keyboard.press('ControlOrMeta+z');
  await waitForChanges(page, 'a', changesBefore + 2);
  const undone = await handleDeck(page, 'a');
  assert.equal(getShapeText(firstShape(undone)), 'Quarterly review');
  assert.equal(getShapeFillColor(firstShape(undone)), getShapeFillColor(firstShape(original)));
  await page.keyboard.press('ControlOrMeta+Shift+z');
  await waitForChanges(page, 'a', changesBefore + 3);
  assert.equal(getShapeText(firstShape(await handleDeck(page, 'a'))), 'Agent title');
  assert.deepEqual((await events(page, 'a')).change.slice(changesBefore), [
    'agent',
    'user',
    'user',
  ]);

  // An edit that throws leaves nothing behind: no partial text, no undo step,
  // no change event.
  const failed = await apply(
    page,
    'a',
    'Half done',
    `pptx.setShapeText(resolveShape(presentation, arg), 'partial'); throw new Error('model gave up');`,
    ref,
  );
  assert.equal(failed, 'model gave up');
  assert.equal(getShapeText(firstShape(await handleDeck(page, 'a'))), 'Agent title');
  assert.equal(await undoMenuItem(editor, page), 'Undo Agent: Make the title red');
  assert.equal((await events(page, 'a')).change.length, changesBefore + 3);
  // An asynchronous edit is refused rather than committed half-way.
  assert.match(await apply(page, 'a', 'Async', `return Promise.resolve();`), /must be synchronous/);

  // Once the shape is gone, its ref fails loudly instead of editing something else.
  assert.equal(
    await apply(
      page,
      'a',
      'Delete the title',
      `pptx.removeShape(resolveShape(presentation, arg));`,
      ref,
    ),
    null,
  );
  assert.deepEqual(await page.evaluate(() => window.handles.a.selection()), []);
  await page.waitForFunction(() => window.events.a.selectionchange.at(-1).length === 0);
  assert.match(
    await apply(
      page,
      'a',
      'Recolor',
      `pptx.setShapeFill(resolveShape(presentation, arg), '#00FF00');`,
      ref,
    ),
    /is gone/,
  );

  // A user's edit reports itself as such.
  const count = (await events(page, 'a')).change.length;
  await editor.getByRole('button', { name: 'New Slide options', exact: true }).click();
  await editor.getByRole('menuitem', { name: 'Title Slide', exact: true }).click();
  await waitForChanges(page, 'a', count + 1);
  assert.equal((await events(page, 'a')).change.at(-1), 'user');
  assert.deepEqual(errors, []);
  await page.close();
});

test('names agent edits in Japanese', async () => {
  const { page, errors } = await hostPage();
  const editor = await mount(page, 'a', await deck('日本語のタイトル'), { locale: 'ja' });
  assert.equal(
    await apply(
      page,
      'a',
      'タイトルを赤に',
      `pptx.setShapeFill(pptx.getSlideShapes(pptx.getSlides(presentation)[0])[0], '#FF0000');`,
    ),
    null,
  );
  assert.equal(
    await undoMenuItem(editor, page, { edit: ja.Edit, undo: ja.Undo }),
    `${ja.Undo} エージェント: タイトルを赤に`,
  );
  assert.deepEqual(errors, []);
  await page.close();
});

test('agents drive each editor on a page separately', async () => {
  const { page, errors } = await hostPage();
  // Before the deck is open there is nothing to edit, and after destroy() nothing either.
  const early = await page.evaluate(
    async (bytes) => {
      const handle = window.mountEditor(document.getElementById('b'), {
        source: new Uint8Array(bytes),
      });
      const message = await handle
        .apply('Too early', () => {})
        .then(
          () => null,
          (error) => error.message,
        );
      await handle.ready;
      handle.destroy();
      const destroyed = await handle
        .apply('Too late', () => {})
        .then(
          () => null,
          (error) => error.message,
        );
      return { message, destroyed };
    },
    await deck('Early'),
  );
  assert.match(early.message, /not ready/);
  assert.match(early.destroyed, /destroyed/);

  const first = await mount(page, 'a', await deck('First deck'));
  await mount(page, 'b', await deck('Second deck'));
  await watchEvents(page, 'a');
  await watchEvents(page, 'b');
  await first.locator('.hit').first().click();
  await page.keyboard.press('Escape');
  assert.equal((await page.evaluate(() => window.handles.a.selection())).length, 1);
  assert.deepEqual(await page.evaluate(() => window.handles.b.selection()), []);

  const retitle = `pptx.setShapeText(pptx.getSlideShapes(pptx.getSlides(presentation)[0])[0], 'Second, edited');`;
  assert.equal(await apply(page, 'b', 'Retitle', retitle), null);
  assert.equal(titleOf(await handleDeck(page, 'b')), 'Second, edited');
  assert.equal(titleOf(await handleDeck(page, 'a')), 'First deck');
  assert.deepEqual((await events(page, 'b')).change, ['agent']);
  assert.deepEqual((await events(page, 'a')).change, []);

  // Undo in the editor the user last used leaves the other's agent edit alone.
  await page.keyboard.press('ControlOrMeta+z');
  assert.equal(titleOf(await handleDeck(page, 'b')), 'Second, edited');

  // An unsubscribed listener hears nothing more.
  await page.evaluate(() => {
    window.unsubscribed = [];
    const off = window.handles.b.on('change', (event) => window.unsubscribed.push(event));
    off();
  });
  assert.equal(
    await apply(page, 'b', 'Again', `pptx.addTitleSlide(presentation, 'Another');`),
    null,
  );
  assert.deepEqual(await page.evaluate(() => window.unsubscribed), []);
  assert.deepEqual(errors, []);
  await page.close();
});

// --- Tools for a model --------------------------------------------------------
// The "model" is a script that picks tools from `tools()` and sends JSON
// through `run`, as a host routes a model's tool calls.

/** Runs `handle.run(name, input)` in the page; resolves to `{ result }` or `{ error }`. */
const runTool = (page, id, name, input) =>
  page.evaluate(
    ({ id, name, input }) =>
      window.handles[id].run(name, input).then(
        (result) => ({ result }),
        (error) => ({ error: error.message }),
      ),
    { id, name, input },
  );

test('a model calls tools by name with JSON, each edit one undo step', async () => {
  const { page, errors } = await hostPage();
  const editor = await mount(page, 'a', await deck('Tool deck'), { locale: 'en' });
  await watchEvents(page, 'a');

  const tools = await page.evaluate(() => window.handles.a.tools());
  const names = tools.map((tool) => tool.name);
  assert.deepEqual(
    names,
    [...names].sort((a, b) => a.localeCompare(b)),
  );
  for (const name of ['listShapes', 'setShapeText', 'setShapeFill', 'addSlideShape'])
    assert.ok(names.includes(name), name);
  // What a model API takes, as it is.
  const setShapeFill = tools.find((tool) => tool.name === 'setShapeFill');
  assert.deepEqual(Object.keys(setShapeFill), ['name', 'description', 'input_schema']);
  assert.equal(setShapeFill.input_schema.type, 'object');

  // Find "the title" the way a model would, then edit it.
  const [slide] = (await runTool(page, 'a', 'listSlides', {})).result;
  assert.deepEqual(slide, {
    slide: '/ppt/slides/slide1.xml',
    index: 0,
    title: 'Tool deck',
    layout: slide.layout,
    hidden: false,
  });
  const shapes = (await runTool(page, 'a', 'listShapes', { slide: slide.slide })).result;
  const title = shapes.find((entry) => entry.text === 'Tool deck').shape;
  assert.deepEqual(
    await runTool(page, 'a', 'setShapeText', { shape: title, value: 'Set by a tool' }),
    { result: null },
  );
  assert.deepEqual(await runTool(page, 'a', 'setShapeFill', { shape: title, color: '#00B050' }), {
    result: null,
  });
  const added = (
    await runTool(page, 'a', 'addSlideShape', {
      slide: slide.slide,
      opts: { preset: 'ellipse', x: 457200, y: 457200, w: 914400, h: 914400, text: 'New' },
    })
  ).result;
  assert.equal(added.slide, slide.slide);
  assert.deepEqual(
    await runTool(page, 'a', 'setShapePosition', { shape: added, x: 1828800, y: 914400 }),
    { result: null },
  );
  assert.deepEqual((await events(page, 'a')).change, ['agent', 'agent', 'agent', 'agent']);

  // The document says so, read back with the core library.
  const edited = await handleDeck(page, 'a');
  const editedTitle = findShapeById(getSlides(edited)[0], title.shapeId);
  assert.equal(getShapeText(editedTitle), 'Set by a tool');
  assert.equal(getShapeFillColor(editedTitle), '#00B050');
  const circle = findShapeById(getSlides(edited)[0], added.shapeId);
  assert.equal(getShapeText(circle), 'New');
  assert.deepEqual(getShapePosition(circle), { x: 1828800, y: 914400 });
  assert.match(await editor.locator('.canvas-shell .paint').textContent(), /Set by a tool/);

  // One Undo takes back the last tool call only.
  assert.equal(await undoMenuItem(editor, page), 'Undo Agent: setShapePosition');
  await page.keyboard.press('ControlOrMeta+z');
  await waitForChanges(page, 'a', 5);
  const undone = await handleDeck(page, 'a');
  assert.deepEqual(getShapePosition(findShapeById(getSlides(undone)[0], added.shapeId)), {
    x: 457200,
    y: 457200,
  });
  assert.equal(getShapeText(findShapeById(getSlides(undone)[0], title.shapeId)), 'Set by a tool');

  // Input that does not fit is refused with what to fix, and nothing changes.
  const refused = await runTool(page, 'a', 'setShapeFill', { shape: title, color: 'green' });
  assert.match(refused.error, /^Invalid input for setShapeFill:\n- \/color: matches none/);
  assert.match(
    (await runTool(page, 'a', 'setShapePosition', { shape: title, x: '1in' })).error,
    /- \(input\): missing required "y"\n- \/x: expected integer, got string "1in"$/,
  );
  assert.match((await runTool(page, 'a', 'paintItBlack', {})).error, /There is no tool/);
  assert.equal((await events(page, 'a')).change.length, 5);
  assert.equal(await undoMenuItem(editor, page), 'Undo Agent: addSlideShape');
  assert.deepEqual(errors, []);
  await page.close();
});

test('names tool calls in Japanese', async () => {
  const { page, errors } = await hostPage();
  const editor = await mount(page, 'a', await deck('ツール'), { locale: 'ja' });
  const [slide] = (await runTool(page, 'a', 'listSlides', {})).result;
  assert.deepEqual(
    await runTool(page, 'a', 'setSlideHidden', { slide: slide.slide, hidden: true }),
    {
      result: null,
    },
  );
  assert.equal(
    await undoMenuItem(editor, page, { edit: ja.Edit, undo: ja.Undo }),
    `${ja.Undo} エージェント: setSlideHidden`,
  );
  assert.equal(isSlideHidden(getSlides(await handleDeck(page, 'a'))[0]), true);
  assert.deepEqual(errors, []);
  await page.close();
});

// --- Proposals ----------------------------------------------------------------
// A long-running agent snapshots the deck, works on its copy and hands it back
// with `propose`; the user may have edited meanwhile.

/** The deck's bytes with shape `index` of the first slide set to `text`. */
async function withText(bytes, index, text) {
  const presentation = await loadPresentation(new Uint8Array(bytes));
  setShapeText(getSlideShapes(getSlides(presentation)[0])[index], text);
  return [...(await savePresentation(presentation))];
}
const snapshot = (page, id) =>
  page.evaluate(async (id) => [...(await window.handles[id].snapshot())], id);
const firstTexts = async (page, id) =>
  getSlideShapes(getSlides(await handleDeck(page, id))[0]).map(getShapeText);

/** Starts `handle.propose` in the page; `settled(page)` resolves to its result. */
const propose = (page, id, base, edited, options) =>
  page.evaluate(
    ({ id, base, edited, options }) => {
      window.proposal = window.handles[id]
        .propose(base && new Uint8Array(base), new Uint8Array(edited), options)
        .then(
          (result) => result,
          (error) => ({ error: error.message }),
        );
    },
    { id, base, edited, options },
  );
const settled = (page) => page.evaluate(() => window.proposal);
const dirty = (page) => page.evaluate(() => window.handles.a.dirty);

/** The user edits the first shape's text on the canvas (not saved). */
async function userEdits(page, editor, text) {
  await editor.locator('.hit').first().dblclick();
  await editor.locator('.canvas-shell .inline-edit').fill(text);
  await editor.locator('.canvas-shell .inline-edit').press('Control+Enter');
  await page.waitForFunction(() => window.handles.a.dirty);
}

test('a proposal merges with unsaved user edits as one undo step', async () => {
  const { page, errors } = await hostPage();
  const editor = await mount(page, 'a', await deck('Plan'), { locale: 'en' });
  await watchEvents(page, 'a');
  const base = await snapshot(page, 'a');
  // While the agent works on its copy, the user retitles the slide.
  await userEdits(page, editor, 'User title');
  const edited = await withText(base, 1, 'Agent subtitle');

  await propose(page, 'a', base, edited, { label: 'Write the subtitle' });
  assert.deepEqual(await settled(page), { status: 'applied' });
  assert.deepEqual(await firstTexts(page, 'a'), ['User title', 'Agent subtitle']);
  assert.equal((await events(page, 'a')).change.at(-1), 'agent');
  assert.equal(await dirty(page), true);
  assert.equal(await undoMenuItem(editor, page), 'Undo Agent: Write the subtitle');
  // Undo takes back the agent's step only.
  await page.keyboard.press('ControlOrMeta+z');
  await page.waitForFunction(() => window.events.a.change.at(-1) === 'user');
  assert.deepEqual(await firstTexts(page, 'a'), ['User title', '']);

  // Handing back what the deck already has changes nothing.
  const count = (await events(page, 'a')).change.length;
  await propose(page, 'a', base, await snapshot(page, 'a'), { label: 'No-op' });
  assert.deepEqual(await settled(page), { status: 'applied' });
  assert.equal((await events(page, 'a')).change.length, count);
  assert.deepEqual(errors, []);
  await page.close();
});

test('a colliding proposal waits for the user, who keeps their edits or takes the agent version', async () => {
  const { page, errors } = await hostPage();
  const editor = await mount(page, 'a', await deck('Plan'), { locale: 'en' });
  const base = await snapshot(page, 'a');
  await userEdits(page, editor, 'User title');
  const edited = await withText(base, 0, 'Agent title');
  const conflicts = [
    {
      part: 'ppt/slides/slide1.xml',
      slide: 1,
      shape: { id: '2', name: 'Centered Title 1' },
      reason: 'both-changed',
    },
  ];

  await propose(page, 'a', base, edited, { label: 'Retitle' });
  const bar = editor.getByRole('alert').filter({ hasText: 'Keep my edits' });
  await bar
    .getByText('Slide 1: Centered Title 1 was changed both here and by the agent.', {
      exact: true,
    })
    .waitFor();
  assert.equal(
    await bar
      .getByText(
        "These edits could not be merged with the agent's changes. Your edits are still here.",
        {
          exact: true,
        },
      )
      .count(),
    1,
  );
  assert.equal(
    await bar.getByRole('link', { name: "Download the agent's version" }).getAttribute('download'),
    'agent.pptx',
  );
  // Until the user chooses, nothing of the proposal is applied.
  assert.deepEqual(await firstTexts(page, 'a'), ['User title', '']);
  await bar.getByRole('button', { name: 'Keep my edits', exact: true }).click();
  assert.deepEqual(await settled(page), { status: 'kept-mine', conflicts });
  assert.equal(await bar.count(), 0);
  assert.deepEqual(await firstTexts(page, 'a'), ['User title', '']);
  assert.equal(await dirty(page), true);

  // In Japanese, the user takes the agent's version: one undo step, which Undo takes back.
  await editor.locator('.lang select').selectOption('ja');
  await propose(page, 'a', base, edited, { label: '題名を変更' });
  const jaBar = editor.getByRole('alert').filter({ hasText: ja['Keep my edits'] });
  await jaBar
    .getByText(
      'スライド 1: Centered Title 1 はこのエディターとエージェントの両方で変更されました。',
      { exact: true },
    )
    .waitFor();
  await jaBar.getByRole('button', { name: ja["Use the agent's version"], exact: true }).click();
  assert.deepEqual(await settled(page), { status: 'took-theirs', conflicts });
  assert.deepEqual(await firstTexts(page, 'a'), ['Agent title', '']);
  assert.equal(
    await undoMenuItem(editor, page, { edit: ja.Edit, undo: ja.Undo }),
    `${ja.Undo} エージェント: 題名を変更`,
  );
  // The user's version is one Undo away, not lost.
  await page.keyboard.press('ControlOrMeta+z');
  await page.waitForFunction(async () => {
    const { pptx } = window;
    const presentation = await pptx.loadPresentation(await window.handles.a.snapshot());
    return (
      pptx.getShapeText(pptx.getSlideShapes(pptx.getSlides(presentation)[0])[0]) === 'User title'
    );
  });
  assert.deepEqual(errors, []);
  await page.close();
});

test('a newer source merges without leaving unsaved changes, and asks in its own words', async () => {
  const { page, errors } = await hostPage();
  const editor = await mount(page, 'a', await deck('Plan'), { locale: 'en' });
  await watchEvents(page, 'a');
  const base = await snapshot(page, 'a');
  // The file changed elsewhere and nothing was edited here: the deck stays as saved.
  const rebuilt = await withText(base, 1, 'From the source');
  await propose(page, 'a', base, rebuilt, { label: 'Source changed', from: 'source' });
  assert.deepEqual(await settled(page), { status: 'applied' });
  assert.deepEqual(await firstTexts(page, 'a'), ['Plan', 'From the source']);
  assert.equal((await events(page, 'a')).change.at(-1), 'source');
  assert.equal(await dirty(page), false);
  assert.equal(await undoMenuItem(editor, page), 'Undo Source changed');

  // Without a common version, any difference is the user's choice.
  await userEdits(page, editor, 'User title');
  await propose(page, 'a', null, rebuilt, { label: 'Source changed', from: 'source' });
  const bar = editor.getByRole('alert').filter({ hasText: 'Keep my edits' });
  await bar
    .getByText('The source or saved deck changed. Your edits are still here.', { exact: true })
    .waitFor();
  assert.equal(
    await bar.getByRole('link', { name: 'Download source' }).getAttribute('download'),
    'source.pptx',
  );
  await bar.getByRole('button', { name: 'Use source', exact: true }).click();
  assert.deepEqual(await settled(page), { status: 'took-theirs', conflicts: [] });
  assert.deepEqual(await firstTexts(page, 'a'), ['Plan', 'From the source']);
  // The deck is the saved source again: nothing is left to save.
  assert.equal(await dirty(page), false);

  // destroy() ends a proposal that is still waiting for the user.
  await userEdits(page, editor, 'Mine again');
  await propose(page, 'a', base, await withText(base, 0, 'Theirs'), { label: 'Late' });
  await bar.waitFor();
  await page.evaluate(() => window.handles.a.destroy());
  assert.deepEqual(await settled(page), { error: 'The editor has been destroyed.' });
  assert.deepEqual(errors, []);
  await page.close();
});

// --- Host options ---------------------------------------------------------------

test('a host shows its own status, autosaves, defers saves and opens another deck', async () => {
  const { page, errors } = await hostPage();
  await page.evaluate(
    async (bytes) => {
      window.saves = [];
      window.accept = false;
      window.hostEvents = [];
      const status = document.createElement('span');
      status.id = 'host-status';
      status.textContent = 'Host status';
      const handle = window.mountEditor(document.getElementById('a'), {
        source: new Uint8Array(bytes),
        fileName: 'quarterly.pptx',
        locale: 'en',
        autoSave: true,
        compact: true,
        status,
        // Refuses until the host can save, as a host does while offline.
        onSave: (pptx) => {
          window.saves.push(pptx.length);
          return window.accept;
        },
      });
      window.handles = { a: handle };
      handle.on('dirtychange', ({ dirty }) => window.hostEvents.push(['dirty', dirty]));
      handle.on('localechange', ({ locale }) => window.hostEvents.push(['locale', locale]));
      await handle.ready;
    },
    await deck('Hosted'),
  );
  const editor = page.locator('#a');
  // The host's element is slotted into the title bar, styled by the page.
  await editor.getByText('Host status', { exact: true }).waitFor();
  const [bar, own] = await Promise.all([
    editor.locator('.topbar').boundingBox(),
    page.locator('#host-status').boundingBox(),
  ]);
  assert.ok(own.y >= bar.y && own.y + own.height <= bar.y + bar.height);
  assert.equal(
    await page.evaluate(() => document.getElementById('host-status').parentElement.tagName),
    'OFFICE-KIT-PPTX-EDITOR',
  );
  assert.equal(
    await page.evaluate(() => getComputedStyle(document.getElementById('host-status')).fontSize),
    '40px',
  );
  await editor.getByText('quarterly.pptx').first().waitFor();
  await editor.getByRole('switch', { name: 'AutoSave' }).waitFor();

  // AutoSave calls onSave; a refusal leaves the deck unsaved and is not retried.
  await userEdits(page, editor, 'Edited');
  await page.waitForFunction(() => window.saves.length === 1);
  await new Promise((resolve) => setTimeout(resolve, 1000));
  assert.equal(await page.evaluate(() => window.saves.length), 1);
  assert.equal(await dirty(page), true);
  // The host saves once it can.
  await page.evaluate(async () => {
    window.accept = true;
    await window.handles.a.save();
  });
  assert.equal(await dirty(page), false);

  // Opening another deck starts over; a recovered copy counts as unsaved.
  await page.evaluate(
    (bytes) =>
      window.handles.a.open(new Uint8Array(bytes), { fileName: 'recovered.pptx', unsaved: true }),
    await deck('Recovered'),
  );
  assert.equal(titleOf(await handleDeck(page, 'a')), 'Recovered');
  assert.equal(await dirty(page), true);
  await editor.getByText('recovered.pptx').first().waitFor();

  await editor.locator('.lang select').selectOption('ja');
  await page.waitForFunction(() => window.hostEvents.some(([type]) => type === 'locale'));
  assert.equal(await page.evaluate(() => window.handles.a.locale), 'ja');
  assert.deepEqual(await page.evaluate(() => window.hostEvents), [
    ['dirty', true],
    ['dirty', false],
    ['dirty', true],
    ['locale', 'ja'],
  ]);
  await page.evaluate(() => window.handles.a.destroy());
  assert.equal(await page.locator('#host-status').count(), 0);
  assert.deepEqual(errors, []);
  await page.close();
});

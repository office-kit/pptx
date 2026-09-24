// Resolving and reopening a comment thread in the editor.
//
// The threads here are the modern ones PowerPoint 2021 and Microsoft 365
// write ([MS-PPTX] §2.16.1) — the only format that can record that a thread
// has been dealt with. The deck is built part by part the way that
// specification says PowerPoint writes it, because nothing in this repository
// ships a real file with them in.
//
// What is checked beyond the obvious: resolving and changing nothing else
// still enables Apply (the status is a change like any other), Cancel throws
// it away, the whole dialog is one step of history, and a resolved thread can
// be brought back into view and reopened — in either language.

import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  _internalPackageOf,
  addBlankSlide,
  createPresentation,
  getCommentStatus,
  getCommentText,
  getSlideComments,
  getSlides,
  loadPresentation,
  savePresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const AUTHOR = '{A0000000-0000-0000-0000-000000000001}';
const OPEN_THREAD = '{1D1EC5A0-0000-0000-0000-000000000001}';
const DONE_THREAD = '{1D1EC5A0-0000-0000-0000-000000000002}';

const AUTHORS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p188:authorLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><p188:author id="${AUTHOR}" name="Reviewer" initials="R" userId="S::reviewer@example.com::0" providerId="AD"/></p188:authorLst>`;

const thread = (id, status, text) =>
  `<p188:cm id="${id}" authorId="${AUTHOR}" created="2026-05-15T12:00:00.000" status="${status}"><pc:sldMkLst><pc:docMk/><pc:sldMk cId="3141592653" sldId="256"/></pc:sldMkLst><p188:txBody><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p188:txBody></p188:cm>`;

const COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p188:cmLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pc="http://schemas.microsoft.com/office/powerpoint/2013/main/command">${thread(OPEN_THREAD, 'active', 'Tighten this headline.')}${thread(DONE_THREAD, 'resolved', 'Fixed the footer already.')}</p188:cmLst>`;

/**
 * A deck whose first slide carries the two threads above, plus one extra
 * slide per entry in `also` carrying a single thread of its own.
 */
const buildDeck = async (also = []) => {
  const pres = createPresentation();
  addBlankSlide(pres);
  for (const _ of also) addBlankSlide(pres);
  const pkg = _internalPackageOf(pres);
  const encode = (s) => new TextEncoder().encode(s);
  pkg.addPart('/ppt/authors.xml', 'application/vnd.ms-powerpoint.authors+xml', encode(AUTHORS_XML));
  const presRels = pkg.getRels('/ppt/presentation.xml');
  presRels.items.push({
    id: 'rId900',
    type: 'http://schemas.microsoft.com/office/2018/10/relationships/authors',
    target: '/ppt/authors.xml',
    targetMode: 'Internal',
  });
  pkg.setRels('/ppt/presentation.xml', presRels);
  const attach = (n, xml) => {
    const target = `/ppt/comments/modernComment${n}.xml`;
    pkg.addPart(target, 'application/vnd.ms-powerpoint.comments+xml', encode(xml));
    const slideRels = pkg.getRels(`/ppt/slides/slide${n}.xml`);
    slideRels.items.push({
      id: 'rId901',
      type: 'http://schemas.microsoft.com/office/2018/10/relationships/comments',
      target,
      targetMode: 'Internal',
    });
    pkg.setRels(`/ppt/slides/slide${n}.xml`, slideRels);
  };
  attach(1, COMMENTS_XML);
  also.forEach((text, at) =>
    attach(
      at + 2,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p188:cmLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pc="http://schemas.microsoft.com/office/powerpoint/2013/main/command">${thread(`{1D1EC5A0-0000-0000-0000-00000000000${at + 3}}`, 'active', text)}</p188:cmLst>`,
    ),
  );
  return savePresentation(pres);
};

const openEditor = async (dir, also) => {
  const source = join(dir, 'source.pptx');
  const file = join(dir, 'deck.tsx');
  await writeFile(source, await buildDeck(also));
  await writeFile(
    file,
    `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';` +
      `export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
  );
  const preview = await startPreview(file);
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await page.goto(preview.url);
  await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
  const editor = page.frameLocator('#editor-frame');
  await editor.getByText('Saved to this project', { exact: true }).waitFor();
  return { preview, browser, page, editor, errors };
};

/** What the deck on disk says: every slide's threads, text and status. */
const stored = async (preview) => {
  const pres = await loadPresentation(
    new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
  );
  return getSlides(pres).map((slide) =>
    getSlideComments(slide).map((comment) => [getCommentText(comment), getCommentStatus(comment)]),
  );
};

/** Waits for the file itself to say what is expected, not for a label. */
const settles = async (preview, expected, what) => {
  const deadline = Date.now() + 20000;
  let last;
  for (;;) {
    last = await stored(preview);
    if (JSON.stringify(last) === JSON.stringify(expected)) return;
    if (Date.now() > deadline) {
      assert.fail(`${what}: the deck never got there, last was ${JSON.stringify(last)}`);
    }
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
};

/** The comment text the dialog is showing, in order. */
const shown = async (dialog) => {
  const boxes = dialog.locator('section textarea');
  const out = [];
  for (let at = 0; at < (await boxes.count()); at += 1) out.push(await boxes.nth(at).inputValue());
  return out;
};

test('a thread is resolved, hidden, brought back and reopened', { timeout: 120000 }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-comment-resolve-'));
  let session;
  try {
    session = await openEditor(dir, []);
    const { preview, editor, errors } = session;
    let ja = false;
    const word = (en, jp) => (ja ? jp : en);
    const open = async () => {
      await editor.getByRole('button', { name: word('Insert', '挿入'), exact: true }).click();
      await editor.locator('button[title$="— addSlideComment"]').click();
      return editor.getByRole('dialog', { name: word('Comments', 'コメント'), exact: true });
    };
    const apply = (dialog) =>
      dialog.getByRole('button', { name: word('Apply', '適用'), exact: true });

    // The deck came through the preview's own save with both threads intact.
    assert.deepEqual(await stored(preview), [
      [
        ['Tighten this headline.', 'active'],
        ['Fixed the footer already.', 'resolved'],
      ],
    ]);

    let dialog = await open();
    // The resolved one is out of the way, and says so rather than vanishing.
    await dialog.getByLabel('Show resolved threads', { exact: true }).waitFor();
    assert.deepEqual(await shown(dialog), ['Tighten this headline.']);
    assert.match(await dialog.locator('.inline').innerText(), /\(1\)/);

    // --- Resolving, and nothing else --------------------------------------
    assert.equal(await apply(dialog).isDisabled(), true);
    await dialog.getByRole('button', { name: 'Resolve 1', exact: true }).click();
    assert.equal(
      await apply(dialog).isDisabled(),
      false,
      'a status change on its own has to be appliable',
    );

    // --- Cancel throws it away --------------------------------------------
    await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
    assert.deepEqual(await stored(preview), [
      [
        ['Tighten this headline.', 'active'],
        ['Fixed the footer already.', 'resolved'],
      ],
    ]);

    // --- And Apply keeps it ------------------------------------------------
    dialog = await open();
    await dialog.getByRole('button', { name: 'Resolve 1', exact: true }).click();
    await apply(dialog).click();
    await settles(
      preview,
      [
        [
          ['Tighten this headline.', 'resolved'],
          ['Fixed the footer already.', 'resolved'],
        ],
      ],
      'resolve',
    );

    // --- One step of history ----------------------------------------------
    await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
    await settles(
      preview,
      [
        [
          ['Tighten this headline.', 'active'],
          ['Fixed the footer already.', 'resolved'],
        ],
      ],
      'undo the resolve',
    );
    await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
    await settles(
      preview,
      [
        [
          ['Tighten this headline.', 'resolved'],
          ['Fixed the footer already.', 'resolved'],
        ],
      ],
      'redo the resolve',
    );

    // --- Bringing a resolved thread back, in Japanese ----------------------
    await editor.locator('.lang select').selectOption('ja');
    ja = true;
    await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
    dialog = await open();
    // Both are resolved now, so the list says so instead of looking empty.
    await dialog
      .getByText('このスライドのスレッドはすべて解決済みです。', { exact: true })
      .waitFor();
    await dialog.getByLabel('解決済みのスレッドを表示', { exact: true }).check();
    assert.deepEqual(await shown(dialog), ['Tighten this headline.', 'Fixed the footer already.']);
    assert.equal(await dialog.getByText('解決済み', { exact: true }).count(), 2);
    await dialog.getByRole('button', { name: '再開する 2', exact: true }).click();
    await apply(dialog).click();
    await settles(
      preview,
      [
        [
          ['Tighten this headline.', 'resolved'],
          ['Fixed the footer already.', 'active'],
        ],
      ],
      'reopen in Japanese',
    );

    assert.deepEqual(errors, []);
  } finally {
    await session?.browser.close();
    await session?.preview.close();
    await rm(dir, { recursive: true, force: true });
  }
});

test(
  'one Apply carries every slide, and one Undo takes it all back',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-comment-resolve-all-'));
    let session;
    try {
      session = await openEditor(dir, ['Check the second slide too.']);
      const { preview, editor, errors } = session;
      const open = async () => {
        await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
        await editor.locator('button[title$="— addSlideComment"]').click();
        return editor.getByRole('dialog', { name: 'Comments', exact: true });
      };

      const before = [
        [
          ['Tighten this headline.', 'active'],
          ['Fixed the footer already.', 'resolved'],
        ],
        [['Check the second slide too.', 'active']],
      ];
      assert.deepEqual(await stored(preview), before);

      const dialog = await open();
      // Slide one: resolve the open thread.
      await dialog.getByRole('button', { name: 'Resolve 1', exact: true }).click();
      // Slide two: resolve its thread and edit the text, without leaving the
      // dialog. The review picker is the only way across, and what was done on
      // slide one has to still be there when Apply is finally pressed.
      await dialog.getByLabel('Review slide', { exact: true }).selectOption('1');
      assert.deepEqual(await shown(dialog), ['Check the second slide too.']);
      await dialog.getByLabel('Comment text', { exact: true }).fill('Second slide, seen to.');
      await dialog.getByRole('button', { name: 'Resolve 1', exact: true }).click();
      await dialog.getByRole('button', { name: 'Apply', exact: true }).click();

      const after = [
        [
          ['Tighten this headline.', 'resolved'],
          ['Fixed the footer already.', 'resolved'],
        ],
        [['Second slide, seen to.', 'resolved']],
      ];
      await settles(preview, after, 'apply across both slides');

      // Everything the dialog did is one step, however many slides it touched.
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await settles(preview, before, 'undo the whole dialog');
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await settles(preview, after, 'redo the whole dialog');

      assert.deepEqual(errors, []);
    } finally {
      await session?.browser.close();
      await session?.preview.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

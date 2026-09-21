import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, readFile, rm, symlink, mkdir } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createTextEditor } from '../src/text-edit.ts';

test('literal edits escape TSX and refuse ambiguous, stale, shared and concurrent edits', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-text-'));
  try {
    const file = join(dir, 'deck.tsx');
    const initial = 'const deck = <Text>Hello</Text>';
    const replacement = 'A < B & "日本語"';
    let state = {
      revision: 1,
      slides: ['svg1', 'svg2'],
      slideTexts: ['Hello', 'Other'],
      dependencies: [file],
    };
    let behavior = 'normal';
    const editor = createTextEditor(
      file,
      () => state,
      async () => {
        const source = await readFile(file, 'utf8');
        if (behavior === 'concurrent') {
          await writeFile(file, source + '\n// external edit');
          throw new Error('concurrent');
        }
        const text = source === initial ? 'Hello' : replacement;
        state = {
          ...state,
          revision: state.revision + 1,
          slides: [text, behavior === 'shared' ? text : 'svg2'],
          slideTexts: [text, 'Other'],
        };
        return null;
      },
    );
    await writeFile(file, initial);
    const request = () => ({
      revision: state.revision,
      slide: 0,
      before: 'Hello',
      after: replacement,
    });
    await assert.rejects(editor.edit({ ...request(), revision: 0 }), /Preview changed/);
    await editor.edit(request());
    assert.equal(
      await readFile(file, 'utf8'),
      'const deck = <Text>{' + JSON.stringify(replacement) + '}</Text>',
    );
    await editor.undo();
    assert.equal(await readFile(file, 'utf8'), initial);
    await writeFile(file, initial + '; const second = "Hello"');
    await assert.rejects(editor.edit(request()), /multiple source matches/);
    await writeFile(file, initial);
    behavior = 'shared';
    await assert.rejects(editor.edit(request()), /other text or slides/);
    assert.equal(await readFile(file, 'utf8'), initial);
    state.slides = ['svg1', 'svg2'];
    behavior = 'concurrent';
    await assert.rejects(editor.edit(request()), /newer changes were preserved/);
    assert.match(await readFile(file, 'utf8'), /external edit/);
    behavior = 'normal';
    await writeFile(file, initial);
    await editor.edit(request());
    await writeFile(file, 'newer user source');
    await assert.rejects(editor.undo(), /newer changes were preserved/);
    assert.equal(await readFile(file, 'utf8'), 'newer user source');
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('text edits cannot follow a dependency symlink outside the project', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-text-scope-'));
  try {
    await writeFile(join(dir, 'outside.tsx'), '<Text>Hello</Text>');
    const project = join(dir, 'project');
    await mkdir(project);
    const entry = join(project, 'deck.tsx');
    await symlink(join(dir, 'outside.tsx'), entry);
    const editor = createTextEditor(
      entry,
      () => ({ revision: 1, slides: ['svg'], slideTexts: ['Hello'], dependencies: [entry] }),
      async () => null,
    );
    await assert.rejects(
      editor.edit({ revision: 1, slide: 0, before: 'Hello', after: 'Changed' }),
      /computed or has multiple/,
    );
    assert.equal(await readFile(join(dir, 'outside.tsx'), 'utf8'), '<Text>Hello</Text>');
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('multiline JSX literals retain their rendered whitespace when edited', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-text-lines-'));
  try {
    const file = join(dir, 'deck.tsx');
    await writeFile(file, '<Text>\n  Hello\n  world\n</Text>');
    let text = 'Hello world';
    const editor = createTextEditor(
      file,
      () => ({ revision: 1, slides: [text], slideTexts: [text], dependencies: [file] }),
      async () => {
        text = 'Changed';
        return null;
      },
    );
    await editor.edit({ revision: 1, slide: 0, before: 'Hello world', after: 'Changed' });
    assert.equal(await readFile(file, 'utf8'), '<Text>{"Changed"}</Text>');
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

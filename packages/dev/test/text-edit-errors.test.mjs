import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createTextEditor } from '../src/text-edit.ts';

test('a refused edit reports its reason even when the source changed before it could be restored', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-text-errors-'));
  try {
    const file = join(dir, 'deck.tsx');
    await writeFile(file, 'const deck = <Text>Hello</Text>');
    const state = {
      revision: 1,
      slides: ['svg1'],
      slideTexts: ['Hello'],
      dependencies: [file],
    };
    const editor = createTextEditor(
      file,
      () => state,
      async () => {
        await writeFile(file, (await readFile(file, 'utf8')) + '\n// external edit');
        return 'build failed';
      },
    );
    await assert.rejects(
      editor.edit({ revision: 1, slide: 0, before: 'Hello', after: 'Changed' }),
      (error) => {
        assert.match(error.message, /did not build/);
        assert.match(error.message, /newer changes were preserved/);
        assert.match(error.cause.message, /did not build/);
        return true;
      },
    );
    assert.match(await readFile(file, 'utf8'), /external edit/);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

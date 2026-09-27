import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm, symlink } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { buildDeck, initProject } from '../dist/index.mjs';

const execute = promisify(execFile);
const script = fileURLToPath(new URL('../../../skill/scripts/deck-text.mjs', import.meta.url));
const devModules = fileURLToPath(new URL('../node_modules', import.meta.url));

// The skill tells agents to name headlines "Headline" and run this script from the
// slide project; this keeps the documented workflow working end to end.
test('the skill deck-text script reads named headlines and flags wording', async (t) => {
  const directory = await mkdtemp(join(tmpdir(), 'pptx-deck-text-test-'));
  t.after(() => rm(directory, { recursive: true, force: true }));
  const project = await initProject(join(directory, 'slides'));
  await symlink(devModules, join(project, 'node_modules'), 'dir');
  const deck = join(project, 'deck.tsx');
  await writeFile(
    deck,
    `import { Presentation, Slide, Text } from '@office-kit/pptx-dsl';
export default (
  <Presentation>
    <Slide>
      <Text name="Headline" x={1} y={0.5} width={10} height={1} size={32}>権限はどこで決まるの？</Text>
      <Text x={1} y={2} width={10} height={1} size={60}>1</Text>
      <Text x={1} y={3.5} width={10} height={1}>行レベルの権限が効く</Text>
    </Slide>
    <Slide notes="ここで RLS が効いてくる">
      <Text name="Headline" x={1} y={0.5} width={10} height={1}>目次</Text>
      <Text x={1} y={4} width={10} height={1}>値は有効期間つきで積む</Text>
      <Text x={1} y={5} width={10} height={1}>DB 制約で守り切る</Text>
      <Text x={1} y={2} width={10} height={1}>—  前提</Text>
    </Slide>
  </Presentation>
);
`,
  );
  const built = await buildDeck(deck);
  await writeFile(join(project, 'deck.pptx'), built.bytes);

  const { stdout } = await execute(process.execPath, [script, 'deck.pptx', '--mode', 'talk'], {
    cwd: project,
  });
  assert.match(stdout, /^p01  権限はどこで決まるの？$/m);
  assert.match(stdout, /^p02  目次$/m);
  assert.match(stdout, /？ 50%/);
  assert.match(stdout, /^p01 \[効く\] 行レベルの権限が効く$/m);
  assert.match(stdout, /^p02 \[dash\] —  前提$/m);
  assert.match(stdout, /^p02 notes \[効く\] ここで RLS が効いてくる$/m);
  assert.match(stdout, /^p02 \[intensifying 切る\] DB 制約で守り切る$/m);
  assert.doesNotMatch(stdout, /有効期間/);

  await assert.rejects(
    execute(process.execPath, [script, 'deck.pptx', '--mode', 'slides'], { cwd: project }),
    /Usage/,
  );
});

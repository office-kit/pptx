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
    `import { Chart, Presentation, Slide, Table, Text } from '@office-kit/pptx-dsl';
export default (
  <Presentation>
    <Slide>
      <Text name="Headline" x={1} y={0.5} width={10} height={1} size={32}>権限はどこで決まるの？</Text>
      <Text x={1} y={2} width={10} height={1} size={60}>1</Text>
      <Text x={1} y={3.5} width={10} height={1}>行レベルの権限が効く</Text>
      <Text x={1} y={5} width={10} height={1}>{'導入は [要確認: 導入社数] 社'}</Text>
    </Slide>
    <Slide notes={'ここで RLS が効いてくる\\n[要確認: 移行日]'}>
      <Text name="Headline" x={1} y={0.5} width={10} height={1}>目次</Text>
      <Text x={1} y={4} width={10} height={1}>値は有効期間つきで積む</Text>
      <Text x={1} y={5} width={10} height={1}>DB 制約で守り切る</Text>
      <Text x={1} y={2} width={10} height={1}>—  前提</Text>
    </Slide>
    <Slide>
      <Text name="Headline" x={1} y={0.5} width={10} height={1}>{'売上は［要確認: 前年比］伸びた'}</Text>
      <Text x={1} y={1.5} width={10} height={0.5} font="Courier New">{'limit = [要確認: 上限]'}</Text>
      <Table x={1} y={2} width={10} height={1} rows={[['項目', '値'], ['${'長い説明'.repeat(20)}', '[要確認: 単価]']]} />
      <Chart x={1} y={4} width={10} height={3} spec={{ kind: 'column', title: '売上（[要確認: 単位]）', categories: ['Q1'], series: [{ name: '売上', values: [1] }] }} />
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
  assert.match(stdout, /？ 33%/);
  assert.match(stdout, /^p01 \[効く\] 行レベルの権限が効く$/m);
  assert.match(stdout, /^p02 \[dash\] —  前提$/m);
  assert.match(stdout, /^p02 notes \[効く\] ここで RLS が効いてくる$/m);
  assert.match(stdout, /^p02 \[intensifying 切る\] DB 制約で守り切る$/m);
  assert.doesNotMatch(stdout, /有効期間/);
  assert.doesNotMatch(stdout, /label: description/);
  const openQuestions = stdout.split('# [要確認] left on slides')[1].split('\n# ')[0];
  assert.match(openQuestions, /^p01 導入は \[要確認: 導入社数\] 社$/m);
  assert.doesNotMatch(openQuestions, /移行日/);
  assert.match(openQuestions, /^p03 売上は［要確認: 前年比］伸びた$/m);
  assert.match(openQuestions, /^p03 limit = \[要確認: 上限\]$/m);
  assert.match(openQuestions, /^p03 ….*\[要確認: 単価\]$/m);
  assert.match(openQuestions, /^p03 売上（\[要確認: 単位\]）$/m);
  assert.match(stdout, /^p02 notes \[要確認: 移行日\]$/m);

  await writeFile(
    deck,
    `import { Presentation, Slide, Text } from '@office-kit/pptx-dsl';
export default (
  <Presentation>
    <Slide>
      <Text name="Headline" x={1} y={0.5} width={10} height={1}>確認済みの資料</Text>
    </Slide>
  </Presentation>
);
`,
  );
  await writeFile(join(project, 'deck.pptx'), (await buildDeck(deck)).bytes);
  const clean = await execute(process.execPath, [script, 'deck.pptx'], { cwd: project });
  assert.match(clean.stdout, /# \[要確認\] left on slides[^\n]*\nnone\n/);

  await assert.rejects(
    execute(process.execPath, [script, 'deck.pptx', '--mode', 'slides'], { cwd: project }),
    /Usage/,
  );
});

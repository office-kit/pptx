// `changeset version`, keeping 0.x packages on 0.x.
//
// pptx-dsl and pptx-preview peer-depend on `@office-kit/pptx` with a caret
// range. A minor bump of a 0.x pptx leaves that range, and changesets then
// bumps every peer dependent a major — 0.x → 1.0.0 — even though no changeset
// asked for it (it shipped 1.0.0 by accident once). Such implied majors become
// minors here; a changeset that names the package as `major` still wins.
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const run = (command, args) => execFileSync(command, args, { encoding: 'utf8' });

const dir = mkdtempSync(join(tmpdir(), 'changeset-status-'));
const statusFile = join(dir, 'status.json');
run('pnpm', ['changeset', 'status', `--output=${statusFile}`]);
const status = JSON.parse(readFileSync(statusFile, 'utf8'));
rmSync(dir, { recursive: true });

const requestedMajor = new Set(
  status.changesets.flatMap((c) => c.releases.filter((r) => r.type === 'major').map((r) => r.name)),
);
const implied = status.releases.filter(
  (r) => r.type === 'major' && r.oldVersion.startsWith('0.') && !requestedMajor.has(r.name),
);

execFileSync('pnpm', ['changeset', 'version'], { stdio: 'inherit' });

const packages = JSON.parse(run('pnpm', ['ls', '-r', '--depth', '-1', '--json']));
const dirOf = new Map(packages.map((p) => [p.name, p.path]));
const changelogs = packages.map((p) => join(p.path, 'CHANGELOG.md'));
const rewrite = (file, edit) => {
  let text;
  try {
    text = readFileSync(file, 'utf8');
  } catch (error) {
    if (error.code === 'ENOENT') return;
    throw error;
  }
  const next = edit(text);
  if (next !== text) writeFileSync(file, next);
};

for (const { name, oldVersion, newVersion } of implied) {
  const minor = Number(oldVersion.split('.')[1]);
  const version = `0.${minor + 1}.0`;
  const root = dirOf.get(name);
  rewrite(join(root, 'package.json'), (text) =>
    text.replace(`"version": "${newVersion}"`, `"version": "${version}"`),
  );
  rewrite(join(root, 'CHANGELOG.md'), (text) =>
    text.replace(`\n## ${newVersion}\n`, `\n## ${version}\n`),
  );
  for (const file of changelogs)
    rewrite(file, (text) => text.replaceAll(`${name}@${newVersion}`, `${name}@${version}`));
  console.log(`${name}: ${newVersion} → ${version} (implied by a peer dependency)`);
}

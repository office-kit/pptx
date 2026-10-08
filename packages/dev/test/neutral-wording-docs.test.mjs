import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import test from 'node:test';

// "PowerPoint" and "Microsoft" are a third party's trademarks, so the docs,
// the editor and dev sources, the site and the changelogs describe behaviour
// neutrally: "the reference desktop app" for verified behaviour of that app,
// "presentation apps" for compatibility in general. This scans the tracked
// files in those areas. The core library, preview, DSL, scripts and core tests
// are covered by the root `test/neutral-wording.test.ts`; the shipped UI
// strings by `brand-free-ui.test.mjs`.
const root = resolve(import.meta.dirname, '../../..');

const SCOPES = [
  'packages/editor',
  'packages/dev',
  'site',
  'docs',
  'skill',
  '.changeset',
  '.github',
  '.claude/skills',
  'test/fixtures/SOURCES.md',
  'packages/preview/CHANGELOG.md',
  'packages/dsl/CHANGELOG.md',
  'samples',
  'tools',
  ':(glob)*.md',
  // Markdown beside the libraries, which the root test (it skips .md) does not read.
  ':(glob)packages/*/**/*.md',
  ':(glob)test/**/*.md',
];

// Files excluded on purpose:
// - The generated editor manifest / tool schema: regenerated from core TSDoc,
//   which the root test covers at its source.
// - THIRD_PARTY_NOTICES.md: legal attribution that must name the third
//   parties and their products as their licences state them.
// - This test and brand-free-ui.test.mjs: they spell the names they forbid.
const EXCLUDED = [
  /^packages\/editor\/src\/manifest\/[^/]*generated/,
  /(^|\/)tools\.generated\.ts$/,
  /^THIRD_PARTY_NOTICES\.md$/,
  /^packages\/dev\/test\/(neutral-wording-docs|brand-free-ui)\.test\.mjs$/,
];
const BINARY =
  /\.(png|jpe?g|gif|webp|ico|pptx|xlsx|docx|pdf|woff2?|ttf|otf|mp3|mp4|wav|webm|zip|emf|wmf)$/i;

const FORBIDDEN = [
  /powerpoint|パワーポイント/i,
  /microsoft|マイクロソフト/i,
  /\bMS[ -]Office\b/i,
  // "Office" as the product name. Case-sensitive: `officeDocument` and
  // `office-kit` are not it; the allowlist below strips the remaining
  // legitimate uses first.
  /\bOffice\b/,
];

// Everything here is stripped before FORBIDDEN is applied. Each entry is a
// use that names the third party for a reason other than describing our
// behaviour.
const ALLOWED = [
  // The trademark disclaimer, verbatim (the READMEs' "Trademarks" sections,
  // the site footer).
  /Not affiliated with or endorsed by Microsoft\. PowerPoint is a trademark of the Microsoft group of companies\./g,
  // OOXML namespace URIs and the vendor content types: wire-format identifiers.
  /(?:https?:\/\/)?schemas\.microsoft\.com\/[^\s"'`<>)\]]*/g,
  /application\/vnd\.ms-(?:powerpoint|office)[\w.+-]*/g,
  // Specification and Open XML SDK reference pages ([MS-OI29500], [MS-PPTX],
  // DocumentFormat.OpenXml API docs): cited like spec section numbers.
  /https?:\/\/learn\.microsoft\.com\/[\w-]+\/(?:openspecs|dotnet\/api\/documentformat\.openxml|previous-versions\/office)[^\s"'`<>)\]]*/g,
  // Third-party open-source repositories that live under that GitHub org.
  /https?:\/\/github\.com\/microsoft\/[^\s"'`<>)\]]*/g,
  // .NET SDK and Open XML SDK API identifiers the validator tool compiles
  // against.
  /Microsoft\.NET\.Sdk|FileFormatVersions\.Microsoft365/g,
  // Part names and font family names inside OOXML packages.
  /Microsoft_Excel_Worksheet/g,
  /Microsoft (?:YaHei|JhengHei|Sans Serif|Himalaya|Yi Baiti|Uighur|Tai Le|New Tai Lue|PhagsPa|MHei|NeoGothic)(?: UI| Light)?/g,
  // The Record tab's Learn More button opens this help article; it is the
  // button's behaviour, not wording.
  /https:\/\/support\.microsoft\.com\/office\/record-a-slide-show-with-narration-and-slide-timings-[\w-]+/g,
  // The fidelity harness drives the app over AppleScript, which needs its name.
  /tell application "Microsoft PowerPoint"/g,
  // Built-in theme, color set and font pair names: data the files and the
  // editor's galleries carry, quoted as names.
  /\bOffice (?:20\d\d - 20\d\d )?(?:Theme\b|テーマ)/g,
  /\bOffice 20\d\d - 20\d\d\b/g,
  /(['"`‘“])Office\1|‘Office’|“Office”/g,
  // Not the product.
  /Office Open XML|LibreOffice|OfficeArt|Office Kit/g,
];

const tracked = execFileSync('git', ['ls-files', '-z', '--', ...SCOPES], {
  cwd: root,
  encoding: 'utf8',
})
  .split('\0')
  .filter((file) => file && !BINARY.test(file) && !EXCLUDED.some((re) => re.test(file)));

test('docs, editor/dev sources, site and changelogs describe third-party apps neutrally', () => {
  // Guards against the pathspecs silently matching nothing.
  for (const file of ['README.md', 'CLAUDE.md', 'packages/dev/NATIVE_PARITY.md']) {
    assert.ok(tracked.includes(file), `${file} was not scanned`);
  }
  const violations = [];
  for (const file of tracked) {
    const lines = readFileSync(resolve(root, file), 'utf8').split('\n');
    lines.forEach((line, index) => {
      const text = ALLOWED.reduce((acc, re) => acc.replace(re, ''), line);
      const pattern = FORBIDDEN.find((re) => re.test(text));
      if (pattern) violations.push(`${file}:${index + 1}: ${line.trim().slice(0, 160)}`);
    });
  }
  assert.deepEqual(violations, []);
});

import assert from 'node:assert/strict';
import { globSync, readFileSync } from 'node:fs';
import { basename, relative, resolve } from 'node:path';
import test from 'node:test';
import ts from 'typescript';

// The editor and the dev tool must not show a third-party trademark
// (PowerPoint, Microsoft, OneDrive, SharePoint, Excel, Word) or our own brand
// (Office Kit) as UI text. Templates, i18n tables, menu data, tooltips and thrown
// messages all end up as string literals in the built bundles, so this scans
// every string literal there — which a source grep cannot do reliably for
// compiled Svelte templates.
//
// Exclusions, each deliberate:
// - `agent-tools-*.js`: the tools() schema, generated from core TSDoc, which
//   is model-facing and states compatibility with the format's owner.
// - Comments inside strings: Svelte ships component CSS (with its comments)
//   as strings, and the dev pages embed scripts with `//` comments.
// - URLs, including OOXML namespace URIs (`schemas.microsoft.com/office/
//   powerpoint/…`), the `application/vnd.ms-powerpoint.*` content types, and
//   links such as the issue tracker.
// - `typeface="…"` values, the `Microsoft_Excel_Worksheet` part name and the
//   creator / application properties: font names, part names and document
//   metadata inside generated OOXML, not UI text.
// - "Word" as the English word in title-case labels (the Change Case
//   command's "Capitalize Each Word").
// - `office-kit` outside HTML text: package specifiers, storage keys, MIME
//   types, element and attribute names. Only "Office Kit" as words, or
//   `office-kit` inside an element's text, is the brand on screen.
const root = resolve(import.meta.dirname, '../../..');
const BUNDLES = [
  'packages/dev/dist/**/*.{js,mjs}',
  'packages/editor/dist/**/*.js',
  'packages/preview/dist/**/*.js',
  'packages/dsl/dist/**/*.js',
  'dist/**/*.js',
];
const isToolSchema = (file) => basename(file).startsWith('agent-tools-');

const FORBIDDEN = [
  /powerpoint|パワーポイント/i,
  /microsoft|マイクロソフト|onedrive|sharepoint/i,
  // Case-sensitive: the `vnd.ms-excel` content types are not UI text, and
  // `\b` skips the `Microsoft_Excel_Worksheet` part name, `WordArt` and
  // `WordprocessingML`. ワード alone is too common in Japanese (パスワード,
  // キーワード), so only product-shaped uses count there.
  /\bExcel\b|\bEXCEL\b|エクセル/,
  /\bWord\b|(?:MS|マイクロソフト) ?ワード|ワード(?:文書|ファイル|形式)/,
  /office kit/i,
  />[^<>{}()=;'"`]*office-kit[^<>{}()=;'"`]*</i,
];

const visibleText = (text) =>
  text
    .replace(/\/\*[\s\S]*?\*\//g, '')
    .replace(/(^|[^:\\])\/\/[^\n]*/g, '$1')
    .replace(/[a-z][\w+.-]*:\/\/[^\s"'`<>)]*/gi, '')
    .replace(/application\/vnd\.ms-powerpoint[\w.+-]*/g, '')
    .replace(/typeface="[^"]*"/g, '')
    .replace(/Microsoft_Excel_Worksheet/g, '')
    .replace(/\b(?:Each|Whole) Words?\b/g, '')
    .replace(/<(dc:creator|cp:lastModifiedBy|Application)>[^<]*<\/\1>/g, '');

function stringLiterals(file) {
  const source = ts.createSourceFile(
    file,
    readFileSync(file, 'utf8'),
    ts.ScriptTarget.Latest,
    false,
    ts.ScriptKind.JS,
  );
  const found = [];
  const visit = (node) => {
    if (ts.isImportDeclaration(node) || ts.isExportDeclaration(node)) return;
    if (
      ts.isStringLiteral(node) ||
      ts.isNoSubstitutionTemplateLiteral(node) ||
      ts.isTemplateHead(node) ||
      ts.isTemplateMiddle(node) ||
      ts.isTemplateTail(node)
    ) {
      found.push(node.text);
    }
    ts.forEachChild(node, visit);
  };
  visit(source);
  return found;
}

test('shipped UI strings name no third-party product and no Office Kit brand', () => {
  const files = globSync(BUNDLES, { cwd: root }).map((file) => resolve(root, file));
  // Guards against the exclusion or the glob silently matching nothing.
  assert.ok(files.some(isToolSchema), 'the tools() schema chunk was not found');
  const scanned = files.filter((file) => !isToolSchema(file));
  assert.ok(
    scanned.some((file) => /["']Editor Help["']/.test(readFileSync(file, 'utf8'))),
    'the editor menu bar was not in the scanned bundles; run the build first',
  );

  const violations = [];
  for (const file of scanned) {
    for (const literal of stringLiterals(file)) {
      const text = visibleText(literal);
      const pattern = FORBIDDEN.find((re) => re.test(text));
      if (!pattern) continue;
      const at = text.search(pattern);
      violations.push(
        `${relative(root, file)}: ${JSON.stringify(text.slice(Math.max(0, at - 60), at + 60))}`,
      );
    }
  }
  assert.deepEqual(violations, []);
});

// Feedback is opt-in: Help ▸ Feedback exists only when `mountDevEditor`
// (`@office-kit/pptx-editor/internal`) is given a URL, and the dev tool is the
// one caller that passes ours. Keeping the URL out of the editor's bundle
// entirely means no public host can reach our tracker from the UI, whatever
// path through the code the user takes; a source grep for menu items alone
// could not promise that.
test('the editor bundle carries no issue-tracker URL; only the dev tool passes one', () => {
  const TRACKER = /github\.com\/office-kit/i;
  const read = (pattern) =>
    globSync(pattern, { cwd: root }).map((file) => [
      file,
      readFileSync(resolve(root, file), 'utf8'),
    ]);
  const editor = read('packages/editor/dist/**/*.js');
  assert.ok(editor.length > 0, 'the editor bundle was not found; run the build first');
  assert.deepEqual(
    editor.filter(([, text]) => TRACKER.test(text)).map(([file]) => file),
    [],
  );
  assert.ok(
    read('packages/dev/dist/**/*.{js,mjs}').some(([, text]) => TRACKER.test(text)),
    'the dev tool no longer passes the feedback URL to the editor',
  );
});

// Help links to a vendor's support site would send users of any host there.
// schemas.microsoft.com namespace URIs are part of the file format and stay.
test('shipped bundles link to no vendor support or product site', () => {
  const VENDOR_SITE = /https?:\/\/(?!schemas\.)[a-z0-9.-]*(?:microsoft|office|live)\.com\b/i;
  const files = globSync('packages/{editor,dev}/dist/**/*.{js,mjs}', { cwd: root });
  assert.ok(files.length > 0, 'no bundles found; run the build first');
  assert.deepEqual(
    files.filter((file) => VENDOR_SITE.test(readFileSync(resolve(root, file), 'utf8'))),
    [],
  );
});

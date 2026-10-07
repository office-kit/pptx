import assert from 'node:assert/strict';
import { globSync, readFileSync } from 'node:fs';
import { basename, relative, resolve } from 'node:path';
import test from 'node:test';
import ts from 'typescript';

// The editor and the dev tool must not show a third-party trademark
// (PowerPoint, Microsoft, OneDrive, SharePoint) or our own brand (Office
// Kit) as UI text. Templates, i18n tables, menu data, tooltips and thrown
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

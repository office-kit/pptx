// "PowerPoint", "Microsoft" and "Office" are Microsoft trademarks, and this
// project is not affiliated with Microsoft. Source, tests and scripts describe
// compatibility neutrally: "presentation apps" in general, or "the reference
// desktop app" for a verified behaviour that deviates from ECMA-376. This
// guard fails on any new use outside the allowlist below.
//
// Markdown files are documentation and are maintained separately.
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const ROOT = fileURLToPath(new URL('..', import.meta.url));
const SCOPE = ['src', 'packages/preview', 'packages/dsl', 'scripts', 'test'];
const SELF = 'test/neutral-wording.test.ts';
const BINARY =
  /\.(pptx|potx|xlsx|docx|png|jpe?g|gif|bmp|tiff?|emf|wmf|svgz|ttf|otf|woff2?|mp3|mp4|m4v|wav|mov|bin|zip)$/i;

const BANNED = [
  /powerpoint|パワーポイント|microsoft|ms[ -]?office/gi,
  /\bOffice\b|\bOFFICE(?=\b|_)/g,
];

// Spans that may contain a banned word, removed before the scan. Each one is a
// specification identifier or a value written into or read from .pptx files,
// so rewording it would change the format, not the prose.
const ALLOWED: readonly RegExp[] = [
  // Extension namespace and relationship-type URIs ([MS-PPTX], [MS-ODRAWXML]),
  // also as escaped inside regular-expression literals.
  /schemas\\?\.microsoft\\?\.com(?:\\?\/[\w.-]*)*/g,
  // Content types of the modern-comments parts ([MS-PPTX] §2.16).
  /application\/vnd\.ms-powerpoint\.[\w.+-]*/g,
  // Part name of the embedded chart workbook the library writes; readers and
  // existing decks expect this exact name.
  /Microsoft_Excel_Worksheet/g,
  // Font family names in the default theme's per-script font lists.
  /Microsoft (?:Yi Baiti|Himalaya|Uighur|JhengHei|Tai Le|New Tai Lue)/g,
  // The ECMA-376 standard's title and the reference schema directory names.
  /Office ?Open ?XML/g,
  // The relationship-type base for the `officeDocument` URIs.
  /\bOFFICE_DOC\b/g,
  // The `Application` value of a fixture's docProps/app.xml, asserted on read.
  /'Microsoft Macintosh PowerPoint'/g,
];

// Trademark disclaimers: a sentence saying the project is not affiliated with
// or endorsed by the trademark owner may name it.
const DISCLAIMER = /\b(?:not affiliated with|not endorsed by|unrelated to|trademarks? of)\b/i;

const listFiles = (): string[] =>
  execFileSync('git', ['ls-files', '--', ...SCOPE], { cwd: ROOT, encoding: 'utf8' })
    .split('\n')
    .filter((f) => f !== '' && f !== SELF && !f.endsWith('.md') && !BINARY.test(f));

const hasBanned = (text: string): boolean => BANNED.some((re) => text.match(re) !== null);

const findings = (path: string, text: string): string[] =>
  text.split('\n').flatMap((line, i) => {
    if (!hasBanned(line)) return [];
    const scrubbed = ALLOWED.reduce((s, re) => s.replace(re, ''), line)
      .split('.')
      .filter((sentence) => !DISCLAIMER.test(sentence))
      .join('.');
    return hasBanned(scrubbed) ? [`${path}:${i + 1}: ${line.trim()}`] : [];
  });

describe('neutral wording', () => {
  it('flags a banned word and lets an allowlisted identifier through', () => {
    expect(findings('a.ts', '// matches PowerPoint here')).toHaveLength(1);
    expect(findings('a.ts', '// the Office theme')).toHaveLength(1);
    expect(findings('a.ts', 'const OFFICE_THEME = 1;')).toHaveLength(1);
    expect(findings('a.ts', `'http://schemas.microsoft.com/office/powerpoint/2010/main'`)).toEqual(
      [],
    );
    expect(findings('a.ts', '// LibreOffice and the reference desktop app')).toEqual([]);
    expect(findings('a.ts', '// Not affiliated with Microsoft. Matches PowerPoint.')).toHaveLength(
      1,
    );
    expect(findings('a.ts', '// This project is not affiliated with Microsoft.')).toEqual([]);
  });

  it('keeps Microsoft product names out of source, tests and scripts', () => {
    const hits = listFiles().flatMap((f) => {
      const text = readFileSync(`${ROOT}${f}`, 'utf8');
      return hasBanned(text) ? findings(f, text) : [];
    });
    expect(hits).toEqual([]);
  });
});

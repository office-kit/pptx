// "PowerPoint", "Microsoft", "Office", "Excel", "Word" and "WordArt" are
// Microsoft trademarks, and this project is not affiliated with Microsoft.
// Source, tests and scripts describe compatibility neutrally: "presentation
// apps" in general, or "the reference desktop app" for a verified behaviour
// that deviates from ECMA-376; "spreadsheet app" / "word processor" for the
// other products and "text art" for styled text. This
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
  // Case-sensitive so the `vnd.ms-excel` content types and the verb "excel"
  // stay out; `\b` already skips `Microsoft_Excel_Worksheet` and `QtExcel`.
  /\bExcel\b|\bEXCEL\b|エクセル/g,
  // The word processor. Case-sensitive so the English word ("word wrap") is
  // not it, and `\b` skips `WordprocessingML` (and `WordArt`, banned below).
  // In Japanese, ワード alone is too common (パスワード, キーワード), so only
  // the product-shaped uses are banned.
  /\bWord\b|(?:Microsoft|MS|マイクロソフト) ?ワード|ワード(?:文書|ファイル|形式)/g,
  // "WordArt" is that app's name for styled text; ours is "text art". The
  // case-insensitive match also catches `wordart` in identifiers and paths.
  /word[ -]?art|ワードアート/gi,
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
  // Title-case UI labels where "Word" is the English word, e.g. the Change
  // Case command.
  /\b(?:Each|Whole) Words?\b/g,
  // "Word" at the start of a sentence, as the English word.
  /\bWord (?=wrap|spacing|break|count|boundar)/g,
  // ECMA-376 values: `ST_TextVerticalType` `wordArtVert` / `wordArtVertRtl`
  // and the `<a:bodyPr fromWordArt>` attribute.
  /\bwordArtVert(?:Rtl)?\b|\bfromWordArt\b/g,
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
    expect(findings('a.ts', '// headroom, as in Excel')).toHaveLength(1);
    expect(findings('a.ts', '// エクセルで開く')).toHaveLength(1);
    expect(findings('a.ts', '// open it in Word')).toHaveLength(1);
    expect(findings('a.ts', '// ワード文書として保存')).toHaveLength(1);
    expect(findings('a.ts', "'/ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx'")).toEqual([]);
    expect(findings('a.ts', "'application/vnd.ms-excel'")).toEqual([]);
    expect(findings('a.ts', '// WordprocessingML; word wrap')).toEqual([]);
    expect(findings('a.ts', "'Capitalize Each Word'; // Word wrap mode")).toEqual([]);
    expect(findings('a.ts', "'パスワード', 'キーワード'")).toEqual([]);
    expect(findings('a.ts', "// the app's WordArt gallery")).toHaveLength(1);
    expect(findings('a.ts', "import './wordart-presets.ts';")).toHaveLength(1);
    expect(findings('a.ts', '// ワードアートのスタイル')).toHaveLength(1);
    expect(findings('a.ts', "setShapeTextDirection(box, 'wordArtVertRtl');")).toEqual([]);
    expect(findings('a.ts', "'wordArtVert' | 'eaVert'; // fromWordArt")).toEqual([]);
  });

  it('keeps Microsoft product names out of source, tests and scripts', () => {
    const hits = listFiles().flatMap((f) => {
      const text = readFileSync(`${ROOT}${f}`, 'utf8');
      return hasBanned(text) ? findings(f, text) : [];
    });
    expect(hits).toEqual([]);
  });

  // File names are not scanned as text, and fixtures, tests and components
  // were named after the feature.
  it('names no tracked file after WordArt', () => {
    const paths = execFileSync('git', ['ls-files'], { cwd: ROOT, encoding: 'utf8' }).split('\n');
    expect(paths.filter((p) => /word[ -]?art/i.test(p))).toEqual([]);
  });
});

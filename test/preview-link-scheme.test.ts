// Untrusted decks must not turn into script in the host page: the preview
// emits `<a href>` only for web, mail and in-document links, and quotes font
// names so they cannot add declarations to an inline `style`.

import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setPresentationFonts,
  setShapeClickAction,
  setShapeRunHyperlink,
  setShapeTextFormat,
  setSlideTitle,
  type PresentationData,
  type SlideData,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import { safeLinkHref } from '../packages/preview/src/link-href.ts';
import { attrsOf } from './lib/svg-query.ts';

const LAYOUTS = ['foreignObject', 'svg'] as const;
const PLACEHOLDER_URL = 'https://placeholder.invalid/';

const ALLOWED = [
  'https://example.com/?a=1&b=2',
  'http://example.com',
  'HTTPS://EXAMPLE.COM',
  ' https://example.com',
  'mailto:someone@example.com',
  '#slide-2',
];

const BLOCKED = [
  'javascript:alert(1)',
  'JaVaScRiPt:alert(1)',
  '  javascript:alert(1)',
  '\u0001javascript:alert(1)',
  '\u0000javascript:alert(1)',
  'java\tscript:alert(1)',
  'java\nscript:alert(1)',
  'java\rscript:alert(1)',
  'javascript\t:alert(1)',
  'data:text/html,<script>alert(1)</script>',
  'DATA:image/svg+xml;base64,PHN2Zy8+',
  'vbscript:msgbox(1)',
  'VBScript:msgbox(1)',
  '&#106;avascript:alert(1)',
  'javascript&colon;alert(1)',
  'file:///etc/passwd',
  'ftp://example.com/',
  'relative/path.html',
  '//evil.example/x',
  '',
];

// C0 controls other than tab/LF/CR cannot be written to XML at all, so only
// the direct `safeLinkHref` check sees them.
const TAB_CODE = 0x09;
const BLOCKED_IN_XML = BLOCKED.filter((href) => ![...href].some((c) => c.charCodeAt(0) < TAB_CODE));

const decode = (value: string): string =>
  value
    .replaceAll('&quot;', '"')
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&amp;', '&');

const hrefsOf = (svg: string): string[] =>
  attrsOf(svg, 'a').map((a) => decode(a['href'] ?? a['xlink:href'] ?? ''));

const blankSlide = async () => {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/blank.pptx', import.meta.url)),
  );
  const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
  const shape = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(1),
    text: 'click me',
  });
  return { pres, slide, shape };
};

const render = (
  pres: PresentationData,
  slide: SlideData,
  textLayout: (typeof LAYOUTS)[number],
): string => renderSlideToSvg(pres, slide, { textLayout });

describe('safeLinkHref', () => {
  it.each(ALLOWED)('keeps %j', (href) => {
    expect(safeLinkHref(href)).toBe(href);
  });
  it.each(BLOCKED)('drops %j', (href) => {
    expect(safeLinkHref(href)).toBeUndefined();
  });
});

describe('preview links', () => {
  describe.each(LAYOUTS)('textLayout %s', (textLayout) => {
    it.each(BLOCKED_IN_XML)('does not link a run hyperlink to %j', async (url) => {
      const { pres, slide, shape } = await blankSlide();
      setShapeRunHyperlink(shape, 0, 0, url);
      expect(hrefsOf(render(pres, slide, textLayout))).toEqual([]);
    });

    it.each(BLOCKED_IN_XML)('does not link a run click action to %j', async (url) => {
      const { pres, slide, shape } = await blankSlide();
      setShapeClickAction(shape, { kind: 'url', url }, { range: { start: 0, end: 5 } });
      expect(hrefsOf(render(pres, slide, textLayout))).toEqual([]);
    });

    it.each(BLOCKED_IN_XML)('does not link a shape click action to %j', async (url) => {
      const { pres, slide, shape } = await blankSlide();
      setShapeClickAction(shape, { kind: 'url', url });
      expect(hrefsOf(render(pres, slide, textLayout))).toEqual([]);
    });

    it.each(ALLOWED.filter((href) => !href.startsWith('#')))(
      'links %j on every path',
      async (url) => {
        const run = await blankSlide();
        setShapeRunHyperlink(run.shape, 0, 0, url);
        expect(hrefsOf(render(run.pres, run.slide, textLayout))).toContain(url);

        const shape = await blankSlide();
        setShapeClickAction(shape.shape, { kind: 'url', url });
        expect(hrefsOf(render(shape.pres, shape.slide, textLayout))).toContain(url);
      },
    );
  });

  // The rels parser decodes character references, so an encoded scheme
  // reaches the preview as the plain one.
  it.each([
    '&#106;avascript:alert(1)',
    '&#x6A;ava&#x09;script:alert(1)',
    'java&#10;script:alert(1)',
    '&#x20;javascript:alert(1)',
    'd&#97;ta:text/html,x',
  ])('does not link a saved rels target written as %j', async (encoded) => {
    const { pres, shape } = await blankSlide();
    setShapeRunHyperlink(shape, 0, 0, PLACEHOLDER_URL);
    const { entries } = readZip(await savePresentation(pres));
    const patched = entries.map((entry) => {
      const text = new TextDecoder().decode(entry.data);
      return entry.name.endsWith('.rels') && text.includes(PLACEHOLDER_URL)
        ? { ...entry, data: new TextEncoder().encode(text.replace(PLACEHOLDER_URL, encoded)) }
        : entry;
    });
    expect(patched).not.toEqual(entries);
    const loaded = await loadPresentation(writeZip(patched));
    const slide = getSlides(loaded).at(-1)!;
    for (const textLayout of LAYOUTS) {
      expect(hrefsOf(render(loaded, slide, textLayout))).toEqual([]);
    }
  });
});

describe('preview font names in style', () => {
  const HOSTILE = 'X;background-image:url(https://evil.example/x)';
  const QUOTE_BREAK = 'Y"; background:url(https://evil.example/y);"';

  // Declarations an injected font name would add, with quoted strings removed.
  const injectedDeclarations = (svg: string): string[] =>
    attrsOf(svg, '[a-zA-Z]+')
      .flatMap((attrs) => (attrs['style'] === undefined ? [] : [decode(attrs['style'])]))
      .map((style) => style.replace(/"(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'/g, '""'))
      .filter((style) => /url\(|background/.test(style));

  it('quotes a run font', async () => {
    const { pres, slide, shape } = await blankSlide();
    setShapeTextFormat(shape, { font: HOSTILE });
    const svg = render(pres, slide, 'foreignObject');
    expect(svg).toContain(`font-family:&quot;${HOSTILE}&quot;, Calibri`);
    expect(injectedDeclarations(svg)).toEqual([]);
  });

  it('escapes a quote inside a run font', async () => {
    const { pres, slide, shape } = await blankSlide();
    setShapeTextFormat(shape, { font: QUOTE_BREAK });
    const svg = render(pres, slide, 'foreignObject');
    expect(injectedDeclarations(svg)).toEqual([]);
  });

  it('quotes the theme major and minor faces', async () => {
    const pres = await loadPresentation(
      await readFile(new URL('./fixtures/minimal/blank.pptx', import.meta.url)),
    );
    setPresentationFonts(pres, { majorLatin: HOSTILE, minorLatin: QUOTE_BREAK });
    const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Title Only')! });
    setSlideTitle(slide, 'Title');
    addSlideTextBox(slide, {
      x: inches(1),
      y: inches(3),
      w: inches(4),
      h: inches(1),
      text: 'Body',
    });
    const svg = render(pres, slide, 'foreignObject');
    expect(svg).toContain(`font-family:&quot;${HOSTILE}&quot;, Calibri`);
    expect(svg).toContain('font-family:&quot;Y\\&quot;; background');
    expect(injectedDeclarations(svg)).toEqual([]);
  });
});

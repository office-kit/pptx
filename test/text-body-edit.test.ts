import fc from 'fast-check';
import { describe, expect, it } from 'vitest';
import { editTextBody, formatTextBodyRange } from '../src/internal/drawingml/text-body-edit.ts';
import { textBodyText } from '../src/internal/drawingml/text-body.ts';
import { parseXml, serializeXml } from '../src/internal/xml/index.ts';

const body = (content: string) =>
  parseXml(
    `<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:lstStyle/>${content}</a:txBody>`,
  );
const rich =
  '<a:p><a:pPr algn="ctr"/><a:r><a:rPr b="1"/><a:t>Hello </a:t></a:r><a:r><a:rPr i="1"/><a:t>世界🌎</a:t></a:r><a:endParaRPr sz="1800"/></a:p><a:p><a:pPr lvl="2"/><a:r><a:rPr u="sng"/><a:t>Tail</a:t></a:r></a:p>';

describe('incremental text replacement', () => {
  it('round-trips arbitrary insertions and paragraph changes', () => {
    fc.assert(
      fc.property(
        fc.array(fc.constantFrom('a', '日', '🌎', '\n', ' ', '<', '&'), { maxLength: 80 }),
        (chars) => {
          const doc = body(rich);
          const value = chars.join('');
          editTextBody(doc.root, value);
          expect(textBodyText(parseXml(serializeXml(doc)).root)).toBe(value);
        },
      ),
      { numRuns: 200 },
    );
  });
  it('keeps insertion formatting after deleting all text', () => {
    const doc = body(rich);
    editTextBody(doc.root, '');
    editTextBody(doc.root, 'Again');
    expect(serializeXml(doc)).toContain('<a:rPr b="1"/><a:t>Again</a:t>');
  });
  it('retains mixed runs, paragraph properties and end properties', () => {
    const doc = body(rich);
    editTextBody(doc.root, 'Hello 日本🌎\nTail');
    expect(textBodyText(doc.root)).toBe('Hello 日本🌎\nTail');
    const xml = serializeXml(doc);
    expect(xml).toContain('<a:rPr b="1"/><a:t>Hello </a:t>');
    expect(xml).toContain('<a:rPr i="1"/><a:t>🌎</a:t>');
    expect(xml).toContain('<a:rPr i="1"/><a:t>日本</a:t>');
    expect(xml).toContain('<a:endParaRPr sz="1800"/>');
    expect(xml).toContain('<a:pPr lvl="2"/>');
  });
  it.each([
    '',
    'new',
    'Hello 世界🌍\nTail',
    'Hello\nnew\n世界🌎\nTail',
    'Hello 世界🌎Tail',
    '\nHello 世界🌎\nTail\n',
  ])('preserves exact visible result: %j', (value) => {
    const doc = body(rich);
    editTextBody(doc.root, value);
    expect(textBodyText(doc.root)).toBe(value);
  });
  it('leaves XML untouched on no-op and retains untouched fields and soft breaks', () => {
    const doc = body(
      '<a:p><a:fld id="{id}" type="slidenum"><a:rPr b="1"/><a:t>1</a:t></a:fld><a:br/><a:r><a:t>suffix</a:t></a:r></a:p>',
    );
    const original = serializeXml(doc);
    editTextBody(doc.root, textBodyText(doc.root));
    expect(serializeXml(doc)).toBe(original);
    editTextBody(doc.root, '1\nsuffix!');
    expect(serializeXml(doc)).toContain('<a:fld id="{id}" type="slidenum">');
    expect(serializeXml(doc)).toContain('<a:br/>');
  });
});

describe('text range formatting', () => {
  it('splits boundary runs and retains mixed formats, paragraph properties and end marks', () => {
    const doc = body(rich);
    formatTextBodyRange(doc.root, { color: '#FF0000' }, { start: 4, end: 12 });
    const xml = serializeXml(doc);
    expect(textBodyText(doc.root)).toBe('Hello 世界🌎\nTail');
    expect(xml).toContain('<a:rPr b="1"/><a:t>Hell</a:t>');
    expect(xml).toContain('<a:rPr u="sng"/><a:t>ail</a:t>');
    expect(xml).toContain('<a:pPr algn="ctr"/>');
    expect(xml).toContain('<a:pPr lvl="2"/>');
    expect(xml).toContain('<a:endParaRPr sz="1800"/>');
    expect(xml).toMatch(/b="1"[^]*?val="FF0000"[^]*?<a:t>o <\/a:t>/);
    expect(xml).toMatch(/i="1"[^]*?val="FF0000"[^]*?<a:t>世界🌎<\/a:t>/);
    expect(xml).toMatch(/u="sng"[^]*?val="FF0000"[^]*?<a:t>T<\/a:t>/);
  });
  it('retains fields and breaks when fully selected, and materializes a partially selected field', () => {
    const content =
      '<a:p><a:fld id="{id}" type="datetime"><a:rPr i="1"/><a:t>2026</a:t></a:fld><a:br/><a:r><a:t>tail</a:t></a:r></a:p>';
    const doc = body(content);
    formatTextBodyRange(doc.root, { bold: true }, { start: 0, end: 5 });
    expect(serializeXml(doc)).toContain('<a:fld id="{id}" type="datetime">');
    expect(serializeXml(doc)).toContain('<a:br><a:rPr b="1"/></a:br>');
    expect(serializeXml(doc)).toContain('<a:r><a:t>tail</a:t></a:r>');
    const partial = body(content);
    formatTextBodyRange(partial.root, { bold: true }, { start: 1, end: 3 });
    expect(serializeXml(partial)).not.toContain('<a:fld');
    expect(serializeXml(partial)).toContain('<a:rPr i="1" b="1"/><a:t>02</a:t>');
    expect(textBodyText(partial.root)).toBe('2026\ntail');
  });
  it.each([
    { start: -1, end: 2 },
    { start: 2, end: 1 },
    { start: 0, end: 99 },
    { start: 0.5, end: 2 },
    { start: 0, end: 9 },
    { start: NaN, end: 2 },
  ])('rejects invalid boundaries without mutation: %j', (range) => {
    const doc = body(rich);
    const original = serializeXml(doc);
    expect(() => formatTextBodyRange(doc.root, { bold: true }, range)).toThrow(RangeError);
    expect(serializeXml(doc)).toBe(original);
  });
  it('preserves the document for empty ranges and rejected formats', () => {
    const doc = body(rich);
    const original = serializeXml(doc);
    formatTextBodyRange(doc.root, { bold: true }, { start: 2, end: 2 });
    expect(serializeXml(doc)).toBe(original);
    expect(() =>
      formatTextBodyRange(doc.root, { bold: true, color: 'invalid' }, { start: 2, end: 5 }),
    ).toThrow();
    expect(serializeXml(doc)).toBe(original);
  });
});

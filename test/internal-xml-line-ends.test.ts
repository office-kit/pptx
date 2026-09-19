// XML 1.0 §2.11 end-of-line handling in the parser: raw CR LF / CR read
// as LF, while a `&#13;` character reference still yields a CR — which is
// how the serializer keeps an authored CR verbatim across a round trip.

import { describe, expect, it } from 'vitest';
import { parseXml, serializeXml } from '../src/internal/xml/index.ts';

const textOf = (xml: string): string => {
  const root = parseXml(xml).root;
  let out = '';
  for (const child of root.children) {
    if (child.kind === 'text' || child.kind === 'cdata') out += child.data;
  }
  return out;
};

describe('xml parser: line-end normalization', () => {
  it('reads a raw CR LF as LF', () => {
    expect(textOf('<t>a\r\nb</t>')).toBe('a\nb');
  });

  it('reads a lone CR as LF', () => {
    expect(textOf('<t>a\rb\r</t>')).toBe('a\nb\n');
  });

  it('keeps a CR written as a character reference', () => {
    expect(textOf('<t>a&#13;\nb</t>')).toBe('a\r\nb');
    expect(textOf('<t>a&#xD;b</t>')).toBe('a\rb');
  });

  it('normalizes CR LF inside CDATA too', () => {
    expect(textOf('<t><![CDATA[a\r\nb]]></t>')).toBe('a\nb');
  });

  it('round-trips a CR LF through the serializer', () => {
    const doc = parseXml('<t>a&#13;\nb</t>');
    const xml = serializeXml(doc);
    expect(xml).toContain('a&#13;\nb');
    expect(textOf(xml)).toBe('a\r\nb');
  });
});

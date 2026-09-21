import { describe, expect, it } from 'vitest';
import { replaceTextInTree } from '../src/internal/drawingml/text-body-mutation.ts';
import { NS, elem, qname, text, textContent } from '../src/internal/xml/index.ts';

const run = (value: string) =>
  elem(qname('a', 'r', NS.dml), {
    children: [elem(qname('a', 't', NS.dml), { children: [text(value)] })],
  });
const paragraph = (values: string[]) =>
  elem(qname('a', 'p', NS.dml), { children: values.map(run) });

describe('text replacement across runs', () => {
  it.each([
    ['日本語 日本語', /日本語/g, 'English'],
    ['Hello hello HELLO', /hello/gi, 'こんにちは'],
    ['A😀B😀C', /😀/gu, '🌸'],
    ['abc abc', /(a)(b)(c)/g, '$3$2$1 / $12 / $01 / $99 / $0'],
    ['abc axc', /a(?<middle>b)?[xc]?/g, '$<middle>/$<missing>'],
    ['abc abc', /abc/g, "$$/$&/$`/$'/$<missing>"],
    ['a b', /(?=.)/g, '-'],
    ['😀B', /(?:)/gu, '.'],
    ['abc', /$/g, '!'],
    ['abc', /^/g, '!'],
    ['abcabc', /abc/g, ''],
    ['abcabc', /abc/g, '$&'],
    ['abc', /z/g, 'unused'],
  ] as const)('matches native replacement for %s with %s', (source, pattern, replacement) => {
    // Split at every UTF-16 offset, including inside an emoji surrogate pair.
    const root = paragraph(source.split(''));
    replaceTextInTree(root, pattern, replacement);
    expect(textContent(root)).toBe(source.replace(pattern, replacement));
  });

  it('does not match across paragraphs, soft breaks, or separate text bodies', () => {
    const first = paragraph(['Hello']);
    first.children.push(elem(qname('a', 'br', NS.dml)), run('world'));
    const root = elem(qname('p', 'spTree', NS.pml), { children: [first, paragraph(['again'])] });
    expect(replaceTextInTree(root, 'Helloworld', 'wrong')).toBe(0);
    expect(replaceTextInTree(root, 'worldagain', 'wrong')).toBe(0);
    expect(textContent(root)).toBe('Helloworldagain');
  });

  it('counts each changed text node once, including with empty adjacent runs', () => {
    const root = paragraph(['', 'ab ab', '', 'cd', '']);
    expect(replaceTextInTree(root, /ab/g, 'X')).toBe(1);
    expect(textContent(root)).toBe('X Xcd');
    expect(replaceTextInTree(root, 'Xcd', 'Y')).toBe(2);
    expect(textContent(root)).toBe('X Y');
    expect(replaceTextInTree(root, 'X Y', '$&')).toBe(0);
  });

  it('leaves caller regular expression state unchanged and replaces every occurrence', () => {
    const expression = /ab/g;
    expression.lastIndex = 3;
    const root = paragraph(['a', 'b ab']);
    replaceTextInTree(root, expression, 'X');
    expect(textContent(root)).toBe('X X');
    expect(expression.lastIndex).toBe(3);
  });

  it('treats a string search as literal and retains replacement token semantics', () => {
    const root = paragraph(['[a', '] $', '1']);
    expect(replaceTextInTree(root, '[a] $1', '$$&')).toBe(3);
    expect(textContent(root)).toBe('$&');
  });
});

import {
  NS,
  elem,
  firstChildElement,
  qname,
  text,
  textContent,
  type XmlElement,
} from '../xml/index.ts';
import { applyRunFormat, type TextFormat } from './text-format.ts';
import { paragraphText, paragraphsOf, textBodyText } from './text-body.ts';

const name = (local: string) => qname('a', local, NS.dml);
const is = (node: XmlElement, local: string) =>
  node.name.namespaceURI === NS.dml && node.name.localName === local;
const copy = (node: XmlElement): XmlElement => structuredClone(node);
const run = (value: string, properties: XmlElement | null): XmlElement =>
  elem(name('r'), {
    children: [
      ...(properties ? [copy(properties)] : []),
      elem(name('t'), { children: [text(value)] }),
    ],
  });

/** Slice visible characters while retaining untouched fields, breaks and run XML. */
function slice(paragraph: XmlElement, start: number, end: number): XmlElement[] {
  const result: XmlElement[] = [];
  let offset = 0;
  for (const child of paragraph.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.dml) continue;
    if (!['r', 'fld', 'br'].includes(child.name.localName)) continue;
    const content = is(child, 'br')
      ? '\n'
      : textContent(firstChildElement(child, name('t')) ?? elem(name('t')));
    const next = offset + content.length;
    const from = Math.max(start, offset);
    const to = Math.min(end, next);
    if (from < to) {
      result.push(
        from === offset && to === next
          ? copy(child)
          : run(content.slice(from - offset, to - offset), firstChildElement(child, name('rPr'))),
      );
    }
    offset = next;
  }
  return result;
}

function propertiesAt(paragraph: XmlElement, at: number, insertion: boolean): XmlElement | null {
  let offset = 0;
  let previous: XmlElement | null = null;
  for (const child of paragraph.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.dml) continue;
    if (!['r', 'fld', 'br'].includes(child.name.localName)) continue;
    const length = is(child, 'br')
      ? 1
      : textContent(firstChildElement(child, name('t')) ?? elem(name('t'))).length;
    const properties = firstChildElement(child, name('rPr'));
    if (at < offset + length || (insertion && at === offset + length)) return properties;
    offset += length;
    previous = properties;
  }
  const end = firstChildElement(paragraph, name('endParaRPr'));
  if (previous || !end) return previous;
  const result = copy(end);
  result.name = name('rPr');
  return result;
}

/** Preserve the unchanged prefix/suffix, including their original paragraph XML. */
export function editTextBody(txBody: XmlElement, value: string): void {
  const before = textBodyText(txBody);
  if (before === value) return;
  // Code points prevent splitting surrogate pairs when two emoji share a high surrogate.
  const oldChars = Array.from(before);
  const newChars = Array.from(value);
  let prefix = 0;
  while (
    prefix < oldChars.length &&
    prefix < newChars.length &&
    oldChars[prefix] === newChars[prefix]
  )
    prefix++;
  let suffix = 0;
  while (
    suffix < oldChars.length - prefix &&
    suffix < newChars.length - prefix &&
    oldChars[oldChars.length - 1 - suffix] === newChars[newChars.length - 1 - suffix]
  )
    suffix++;
  const start = oldChars.slice(0, prefix).join('').length;
  const end = before.length - oldChars.slice(oldChars.length - suffix).join('').length;
  const replacement = newChars.slice(prefix, newChars.length - suffix).join('');
  const paragraphs = paragraphsOf(txBody);
  const spans = paragraphs.map((paragraph) => ({
    paragraph,
    length: paragraphText(paragraph).length,
  }));
  let offset = 0;
  let first = 0;
  let last = 0;
  let firstOffset = 0;
  let lastOffset = 0;
  for (let i = 0; i < spans.length; i++) {
    const span = spans[i]!;
    if (start > offset + span.length) {
      first = i + 1;
      firstOffset = offset + span.length + 1;
    }
    if (end > offset + span.length) {
      last = i + 1;
      lastOffset = offset + span.length + 1;
    }
    offset += span.length + 1;
  }
  const firstParagraph = spans[first]?.paragraph ?? elem(name('p'));
  const lastParagraph = spans[last]?.paragraph ?? firstParagraph;
  const left = slice(firstParagraph, 0, start - firstOffset);
  const right = slice(lastParagraph, end - lastOffset, paragraphText(lastParagraph).length);
  const properties = propertiesAt(firstParagraph, start - firstOffset, start === end);
  const lines = replacement.split('\n');
  const inserted = lines.map((line, index) => {
    const p = copy(firstParagraph);
    const pPr = firstChildElement(p, name('pPr'));
    const endPr = firstChildElement(
      index === lines.length - 1 ? lastParagraph : firstParagraph,
      name('endParaRPr'),
    );
    p.children = [
      ...(pPr ? [pPr] : []),
      ...(index === 0 ? left : []),
      ...(line ||
      ((index !== 0 || left.length === 0) && (index !== lines.length - 1 || right.length === 0))
        ? [run(line, properties)]
        : []),
      ...(index === lines.length - 1 ? right : []),
      ...(endPr ? [copy(endPr)] : []),
    ];
    return p;
  });
  const firstIndex = txBody.children.indexOf(firstParagraph);
  const lastIndex = txBody.children.indexOf(lastParagraph);
  if (firstIndex === -1) txBody.children.push(...inserted);
  else txBody.children.splice(firstIndex, lastIndex - firstIndex + 1, ...inserted);
}

/** Format a UTF-16 text range without replacing unaffected paragraph/run XML. */
export function formatTextBodyRange(
  txBody: XmlElement,
  format: TextFormat,
  range: { start: number; end: number },
): void {
  const value = textBodyText(txBody);
  const { start, end } = range;
  const splitsSurrogate = (at: number) =>
    at > 0 &&
    at < value.length &&
    /[\uD800-\uDBFF]/.test(value[at - 1]!) &&
    /[\uDC00-\uDFFF]/.test(value[at]!);
  if (
    !Number.isInteger(start) ||
    !Number.isInteger(end) ||
    start < 0 ||
    end < start ||
    end > value.length ||
    splitsSurrogate(start) ||
    splitsSurrogate(end)
  ) {
    throw new RangeError(
      'setShapeTextFormat: range must contain valid UTF-16 boundaries within the text',
    );
  }
  // Validate all format values before touching the document, including empty ranges.
  applyRunFormat(elem(name('rPr')), format);
  if (start === end) return;
  const updated = copy(txBody);
  let offset = 0;
  for (const paragraph of paragraphsOf(updated)) {
    const children: typeof paragraph.children = [];
    for (const child of paragraph.children) {
      if (
        child.kind !== 'element' ||
        child.name.namespaceURI !== NS.dml ||
        !['r', 'fld', 'br'].includes(child.name.localName)
      ) {
        children.push(child);
        continue;
      }
      const content = is(child, 'br')
        ? '\n'
        : textContent(firstChildElement(child, name('t')) ?? elem(name('t')));
      const from = Math.max(0, start - offset);
      const to = Math.min(content.length, end - offset);
      offset += content.length;
      if (from >= to) {
        children.push(child);
        continue;
      }
      const fragment = (a: number, b: number) => {
        if (a === 0 && b === content.length) return copy(child);
        // Editing part of a generated field makes those characters literal text.
        if (is(child, 'fld'))
          return run(content.slice(a, b), firstChildElement(child, name('rPr')));
        const result = copy(child);
        const t = firstChildElement(result, name('t'));
        if (t) t.children = [text(content.slice(a, b))];
        return result;
      };
      if (from > 0) children.push(fragment(0, from));
      const selected = fragment(from, to);
      let properties = firstChildElement(selected, name('rPr'));
      if (!properties) {
        properties = elem(name('rPr'));
        selected.children.unshift(properties);
      }
      applyRunFormat(properties, format);
      children.push(selected);
      if (to < content.length) children.push(fragment(to, content.length));
    }
    paragraph.children = children;
    offset++; // The visible paragraph separator occupies one UTF-16 position.
  }
  txBody.children = updated.children;
}

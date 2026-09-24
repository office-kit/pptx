import { applyRunFormat, type TextFormat } from './text-format.ts';
import { NS, type XmlElement, elem, firstChildElement, qname } from '../xml/index.ts';

const NAME_A_P = qname('a', 'p', NS.dml);
const NAME_A_R = qname('a', 'r', NS.dml);
const NAME_A_RPR = qname('a', 'rPr', NS.dml);
const NAME_A_T = qname('a', 't', NS.dml);

const paragraphsOf = (txBody: XmlElement): XmlElement[] =>
  txBody.children.filter(
    (c): c is XmlElement =>
      c.kind === 'element' &&
      c.name.namespaceURI === NAME_A_P.namespaceURI &&
      c.name.localName === 'p',
  );

const readRunText = (run: XmlElement): string => {
  const tEl = firstChildElement(run, NAME_A_T);
  if (tEl === null) return '';
  let out = '';
  for (const child of tEl.children) {
    if (child.kind === 'text' || child.kind === 'cdata') out += child.data;
  }
  return out;
};

const writeRunText = (run: XmlElement, value: string): void => {
  let tEl = firstChildElement(run, NAME_A_T);
  if (tEl === null) {
    tEl = elem(NAME_A_T);
    run.children.push(tEl);
  }
  // Empty XML text nodes disappear when parsed; keep the serialized form stable.
  tEl.children = value ? [{ kind: 'text', data: value }] : [];
};

export const validateTextRange = (value: string, start: number, end: number): void => {
  if (
    !Number.isInteger(start) ||
    !Number.isInteger(end) ||
    start < 0 ||
    end < start ||
    end > value.length
  )
    throw new RangeError('Invalid text range.');
  for (const offset of [start, end]) {
    if (
      offset > 0 &&
      offset < value.length &&
      /[\uD800-\uDBFF]/.test(value[offset - 1]!) &&
      /[\uDC00-\uDFFF]/.test(value[offset]!)
    )
      throw new RangeError('Text range splits a surrogate pair.');
  }
};

const isTextInline = (node: XmlElement['children'][number]): node is XmlElement =>
  node.kind === 'element' &&
  node.name.namespaceURI === NS.dml &&
  ['r', 'fld', 'br'].includes(node.name.localName);
const inlineText = (node: XmlElement): string =>
  node.name.localName === 'br' ? '\n' : readRunText(node);

/** Replace a text-body range without losing unaffected XML, including links and paragraph properties. */
export function replaceTextBodyRange(
  original: XmlElement,
  value: string,
  start: number,
  end: number,
  replacement: string,
): void {
  validateTextRange(value, start, end);
  if (typeof replacement !== 'string') throw new TypeError('Replacement must be text.');
  if (start === end && !replacement) return;
  const body = structuredClone(original);
  const paragraphs = paragraphsOf(body);
  if (!paragraphs.length) paragraphs.push(elem(NAME_A_P));
  const locate = (position: number) => {
    let offset = 0;
    for (let index = 0; index < paragraphs.length; index++) {
      const length = paragraphs[index]!.children.filter(isTextInline).reduce(
        (n, r) => n + inlineText(r).length,
        0,
      );
      if (position <= offset + length) return { index, offset: position - offset };
      offset += length + 1;
    }
    throw new RangeError('Invalid text position.');
  };
  const first = locate(start),
    last = locate(end);
  const slice = (paragraph: XmlElement, from: number, to: number): XmlElement[] => {
    let offset = 0;
    const result: XmlElement[] = [];
    for (const node of paragraph.children.filter(isTextInline)) {
      const value = inlineText(node);
      const a = Math.max(0, from - offset),
        b = Math.min(value.length, to - offset);
      offset += value.length;
      if (a >= b) continue;
      if (a === 0 && b === value.length) result.push(structuredClone(node));
      else {
        // A partially edited field becomes literal text, retaining its run properties.
        const run = node.name.localName === 'r' ? structuredClone(node) : elem(NAME_A_R);
        const properties = firstChildElement(node, NAME_A_RPR);
        if (node.name.localName !== 'r' && properties)
          run.children.push(structuredClone(properties));
        writeRunText(run, value.slice(a, b));
        result.push(run);
      }
    }
    return result;
  };
  const prefix = slice(paragraphs[first.index]!, 0, first.offset);
  const suffix = slice(paragraphs[last.index]!, last.offset, Infinity);
  const firstInline = paragraphs[first.index]!.children.find(isTextInline);
  const selectedInline =
    end > start ? slice(paragraphs[first.index]!, first.offset, first.offset + 1)[0] : undefined;
  const inherited = selectedInline ?? prefix.at(-1) ?? firstInline;
  const properties = inherited ? firstChildElement(inherited, NAME_A_RPR) : null;
  const pieces = replacement.replace(/\r\n?/g, '\n').split('\n');
  const changed = pieces.map((piece, index) => {
    const paragraph = structuredClone(paragraphs[first.index]!);
    const children: XmlElement['children'] = [];
    // pPr precedes runs; endParaRPr follows them.
    const before = paragraph.children.filter(
      (node) => node.kind === 'element' && node.name.localName === 'pPr',
    );
    const after = paragraph.children.filter(
      (node) => !isTextInline(node) && !before.includes(node),
    );
    children.push(...before);
    if (index === 0) children.push(...prefix);
    if (piece) {
      const run = elem(NAME_A_R);
      if (properties) run.children.push(structuredClone(properties));
      writeRunText(run, piece);
      children.push(run);
    }
    if (index === pieces.length - 1) children.push(...suffix);
    if (!children.some(isTextInline) && properties) {
      const run = elem(NAME_A_R);
      run.children.push(structuredClone(properties));
      writeRunText(run, '');
      children.push(run);
    }
    children.push(...after);
    paragraph.children = children;
    return paragraph;
  });
  const result = [
    ...paragraphs.slice(0, first.index),
    ...changed,
    ...paragraphs.slice(last.index + 1),
  ];
  body.children = [
    ...body.children.filter(
      (node) =>
        !(
          node.kind === 'element' &&
          node.name.namespaceURI === NS.dml &&
          node.name.localName === 'p'
        ),
    ),
    ...result,
  ];
  original.children = body.children;
}

const ensureRPr = (run: XmlElement): XmlElement => {
  const existing = firstChildElement(run, NAME_A_RPR);
  if (existing) return existing;
  const fresh = elem(NAME_A_RPR);
  run.children.unshift(fresh);
  return fresh;
};

/** Format a range, preserving unselected text and existing run XML. */
export function formatTextBodyRange(
  original: XmlElement,
  value: string,
  start: number,
  end: number,
  format: TextFormat,
): void {
  validateTextRange(value, start, end);
  // Validate before touching the original tree, including empty selections.
  applyRunFormat(elem(NAME_A_RPR), format);
  mutateTextBodyRange(original, value, start, end, (properties) =>
    applyRunFormat(properties, format),
  );
}

/** Apply a property mutation to selected runs, preserving surrounding XML. Fields remain atomic. */
export function mutateTextBodyRange(
  original: XmlElement,
  value: string,
  start: number,
  end: number,
  mutate: (properties: XmlElement) => void,
): void {
  validateTextRange(value, start, end);
  if (start === end) return;
  const body = structuredClone(original);
  let offset = 0;
  for (const paragraph of paragraphsOf(body)) {
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
      const runText = child.name.localName === 'br' ? '\n' : readRunText(child);
      const from = Math.max(0, start - offset),
        to = Math.min(runText.length, end - offset);
      offset += runText.length;
      if (from >= to) {
        children.push(child);
        continue;
      }
      if (child.name.localName !== 'r') {
        mutate(ensureRPr(child));
        children.push(child);
        continue;
      }
      for (const [a, b, selected] of [
        [0, from, false],
        [from, to, true],
        [to, runText.length, false],
      ] as const) {
        if (a === b) continue;
        const fragment = structuredClone(child);
        writeRunText(fragment, runText.slice(a, b));
        if (selected) mutate(ensureRPr(fragment));
        children.push(fragment);
      }
    }
    paragraph.children = children;
    offset++;
  }
  original.children = body.children;
}

/** Paragraphs touched by a UTF-16 selection; an empty selection targets its caret paragraph. */
export function textBodyParagraphsInRange(
  body: XmlElement,
  value: string,
  start: number,
  end: number,
): XmlElement[] {
  validateTextRange(value, start, end);
  let offset = 0;
  return paragraphsOf(body).filter((paragraph) => {
    const length = paragraph.children
      .filter(isTextInline)
      .reduce((n, child) => n + inlineText(child).length, 0);
    const selected =
      start === end
        ? start >= offset && start <= offset + length
        : start <= offset + length && end > offset;
    offset += length + 1;
    return selected;
  });
}

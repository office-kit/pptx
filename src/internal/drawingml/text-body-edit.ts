import {
  NS,
  elem,
  firstChildElement,
  qname,
  serializeFragment,
  text,
  textContent,
  type XmlElement,
} from '../xml/index.ts';
import { applyRunFormat, resetRunFormat, type TextFormat } from './text-format.ts';
import { paragraphText, paragraphsOf, textBodyText } from './text-body.ts';

const name = (local: string) => qname('a', local, NS.dml);
export type TextCase = 'sentence' | 'lower' | 'upper' | 'title' | 'toggle';
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

/** Replace a run's text without discarding extension or hyperlink children. */
const textFragment = (source: XmlElement, value: string, materializeField = false): XmlElement => {
  const result = copy(source);
  if (materializeField) {
    result.name = name('r');
    // Field identity attributes are only valid on a:fld, not a:r.
    result.attrs = result.attrs.filter(
      (attr) => attr.name.namespaceURI !== '' || !['id', 'type'].includes(attr.name.localName),
    );
    // a:fld permits paragraph properties that are not valid children of a:r.
    result.children = result.children.filter(
      (child) => child.kind !== 'element' || !is(child, 'pPr'),
    );
  }
  const t = firstChildElement(result, name('t'));
  if (t) t.children = [text(value)];
  return result;
};

const isCased = (value: string): boolean => value.toLowerCase() !== value.toUpperCase();

const casePointModes = (source: string, mode: TextCase): TextCase[] => {
  const points = Array.from(source);
  const pointModes: TextCase[] = points.map(() => mode);
  if (mode === 'sentence') {
    let capitalize = true;
    points.forEach((point, index) => {
      if (isCased(point)) {
        pointModes[index] = capitalize ? 'upper' : 'lower';
        capitalize = false;
      }
      if (point === '.' || point === '?' || point === '!' || point === '\n') capitalize = true;
    });
  } else if (mode === 'title') {
    const segmenter = new Intl.Segmenter(undefined, { granularity: 'word' });
    let pointIndex = 0;
    for (const { segment, isWordLike } of segmenter.segment(source)) {
      const segmentPoints = Array.from(segment);
      if (isWordLike) {
        const first = segmentPoints.findIndex(isCased);
        segmentPoints.forEach((point, index) => {
          if (isCased(point)) pointModes[pointIndex + index] = index === first ? 'upper' : 'lower';
        });
      }
      pointIndex += segmentPoints.length;
    }
  }
  return pointModes;
};

const contextualCasePoints = (source: string, mode: TextCase): string[] => {
  const points = Array.from(source);
  const lowered = source.toLowerCase();
  const lowerLengths = points.map((point) => point.toLowerCase().length);
  const lowerTotal = lowerLengths.reduce((sum, length) => sum + length, 0);
  if (lowerTotal !== lowered.length)
    throw new Error('Unicode case mapping cannot be partitioned safely');
  const modes = casePointModes(source, mode);
  const result: string[] = [];
  let lowerOffset = 0;
  points.forEach((point, index) => {
    const length = lowerLengths[index]!;
    result.push(
      modes[index] === 'upper'
        ? point.toUpperCase()
        : lowered.slice(lowerOffset, lowerOffset + length),
    );
    lowerOffset += length;
  });
  return result;
};

const toggleCasePoints = (source: string): string[] => {
  const points = Array.from(source);
  const lowered = source.toLowerCase();
  const lowerLengths = points.map((point) => point.toLowerCase().length);
  if (lowerLengths.reduce((sum, length) => sum + length, 0) !== lowered.length)
    throw new Error('Unicode case mapping cannot be partitioned safely');
  const result: string[] = [];
  let lowerOffset = 0;
  points.forEach((point, index) => {
    const length = lowerLengths[index]!;
    result.push(
      isCased(point) && point === point.toUpperCase()
        ? lowered.slice(lowerOffset, lowerOffset + length)
        : isCased(point)
          ? point.toUpperCase()
          : lowered.slice(lowerOffset, lowerOffset + length),
    );
    lowerOffset += length;
  });
  return result;
};

const transformCaseText = (value: string, mode: TextCase): string => {
  if (mode === 'lower') return value.toLowerCase();
  if (mode === 'upper') return value.toUpperCase();
  if (mode === 'toggle') return toggleCasePoints(value).join('');
  return contextualCasePoints(value, mode).join('');
};

const caseBoundaryMap = (source: string, transformed: string, mode: TextCase): number[] => {
  const points = Array.from(source);
  const lengths =
    mode === 'lower' || mode === 'sentence' || mode === 'title'
      ? contextualCasePoints(source, mode).map((point) => point.length)
      : mode === 'toggle'
        ? toggleCasePoints(source).map((point) => point.length)
        : points.map((point) => transformCaseText(point, mode).length);
  if (lengths.reduce((sum, length) => sum + length, 0) !== transformed.length)
    throw new Error('Unicode case mapping cannot be partitioned safely');
  const boundaries = [0];
  for (const length of lengths) boundaries.push(boundaries.at(-1)! + length);
  return boundaries;
};

/** Changes cached text while retaining the original run/field XML structure. */
export function transformTextBodyCase(
  txBody: XmlElement,
  mode: TextCase,
  range?: { start: number; end: number },
): void {
  if (!['sentence', 'lower', 'upper', 'title', 'toggle'].includes(mode))
    throw new RangeError(`text case must be one of sentence, lower, upper, title, toggle`);
  const before = textBodyText(txBody);
  const target = range ?? { start: 0, end: before.length };
  validateTextRange(before, target, 'setText');
  const selected = before.slice(target.start, target.end);
  const transformed = transformCaseText(selected, mode);
  if (selected === transformed) return;
  const boundaries = caseBoundaryMap(selected, transformed, mode);
  const utf16ToPoint = Array.from<number>({ length: selected.length + 1 });
  let point = 0;
  for (let index = 0; index < selected.length; ) {
    utf16ToPoint[index] = point;
    const width = selected.codePointAt(index)! > 0xffff ? 2 : 1;
    if (width === 2) utf16ToPoint[index + 1] = point;
    index += width;
    point++;
  }
  utf16ToPoint[selected.length] = point;
  const toOutputOffset = (offset: number): number =>
    boundaries[utf16ToPoint[offset]!] ?? transformed.length;
  let offset = 0;
  for (const paragraph of paragraphsOf(txBody)) {
    for (const child of paragraph.children) {
      if (
        child.kind !== 'element' ||
        child.name.namespaceURI !== NS.dml ||
        !['r', 'fld'].includes(child.name.localName)
      ) {
        if (child.kind === 'element' && is(child, 'br')) offset++;
        continue;
      }
      const t = firstChildElement(child, name('t'));
      if (!t) continue;
      const content = textContent(t);
      const nodeStart = offset;
      const nodeEnd = nodeStart + content.length;
      const from = Math.max(target.start, nodeStart);
      const to = Math.min(target.end, nodeEnd);
      if (from < to) {
        const transformedFrom = toOutputOffset(from - target.start);
        const transformedTo = toOutputOffset(to - target.start);
        const prefix = content.slice(0, from - nodeStart);
        const suffix = content.slice(to - nodeStart);
        const replacement = transformed.slice(transformedFrom, transformedTo);
        t.children = [text(prefix + replacement + suffix)];
      }
      offset = nodeEnd;
    }
    offset++;
  }
}

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
          : is(child, 'r') || is(child, 'fld')
            ? textFragment(child, content.slice(from - offset, to - offset), is(child, 'fld'))
            : run(content.slice(from - offset, to - offset), firstChildElement(child, name('rPr'))),
      );
    }
    offset = next;
  }
  return result;
}

/** Copy a UTF-16 range without rebuilding paragraph properties or untouched runs. */
export function copyTextBodyRange(
  txBody: XmlElement,
  range: { start: number; end: number },
): XmlElement[] {
  return copyTextBodyRanges(txBody, [range])[0]!;
}

/** Index once when distributing a body's paragraphs across multiple shapes. */
export function copyTextBodyRanges(
  txBody: XmlElement,
  ranges: ReadonlyArray<{ start: number; end: number }>,
): XmlElement[][] {
  const value = textBodyText(txBody);
  for (const range of ranges) validateTextRange(value, range, 'setShapeParagraphs');
  let offset = 0;
  const indexed = paragraphsOf(txBody).map((paragraph) => {
    const start = offset;
    const end = start + paragraphText(paragraph).length;
    offset = end + 1;
    return { paragraph, start, end };
  });
  return ranges.map((range) => {
    const result: XmlElement[] = [];
    let low = 0;
    let high = indexed.length;
    while (low < high) {
      const mid = Math.floor((low + high) / 2);
      if (indexed[mid]!.end < range.start) low = mid + 1;
      else high = mid;
    }
    for (let index = low; index < indexed.length; index++) {
      const { paragraph, start, end } = indexed[index]!;
      if (start > range.end) break;
      const from = Math.max(0, range.start - start);
      const to = Math.min(end - start, range.end - start);
      const cloned = copy(paragraph);
      if (from !== 0 || to !== end - start) {
        const contents = slice(paragraph, from, to);
        let inserted = false;
        cloned.children = cloned.children.flatMap((child) => {
          if (child.kind !== 'element' || !['r', 'fld', 'br'].some((local) => is(child, local)))
            return [child];
          if (inserted) return [];
          inserted = true;
          return contents;
        });
      }
      result.push(cloned);
    }
    return result;
  });
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

// The properties of a run that holds only text, or `null` for anything else
// (fields, breaks, runs carrying extensions). Two such neighbours with equal
// properties read as one run.
const plainRunProperties = (node: XmlElement): string | null => {
  if (!is(node, 'r') || node.attrs.length > 0) return null;
  let properties = '';
  for (const child of node.children) {
    if (child.kind !== 'element') continue;
    if (is(child, 'rPr')) properties = serializeFragment(child);
    else if (!is(child, 't')) return null;
  }
  return properties;
};

// PowerPoint extends the run the caret is in, so typing never leaves one
// `<a:r>` per keystroke, and deleting the text between two halves of a run
// leaves one run again. Only the seam between `before` and `after` is
// joined; runs elsewhere keep their own structure.
const joinAtSeam = (before: XmlElement[], after: XmlElement[]): XmlElement[] => {
  const left = before.at(-1);
  const right = after[0];
  if (!left || !right) return [...before, ...after];
  const properties = plainRunProperties(left);
  if (properties === null || properties !== plainRunProperties(right)) return [...before, ...after];
  const joined = textFragment(
    left,
    textContent(firstChildElement(left, name('t')) ?? elem(name('t'))) +
      textContent(firstChildElement(right, name('t')) ?? elem(name('t'))),
  );
  return [...before.slice(0, -1), joined, ...after.slice(1)];
};

/** Preserve the unchanged prefix/suffix, including their original paragraph XML. */
export function editTextBody(
  txBody: XmlElement,
  value: string,
  range?: { start: number; end: number },
  newlines: 'paragraph' | 'break' = 'paragraph',
): void {
  const before = textBodyText(txBody);
  if (range) validateTextRange(before, range, 'setText');
  // Explicit newline replacements can change a paragraph separator into a soft
  // break (or vice versa) without changing the plain-text representation.
  if (
    range
      ? before.slice(range.start, range.end) === value && !value.includes('\n')
      : before === value
  )
    return;
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
  const start = range?.start ?? oldChars.slice(0, prefix).join('').length;
  const end =
    range?.end ?? before.length - oldChars.slice(oldChars.length - suffix).join('').length;
  const replacement = range ? value : newChars.slice(prefix, newChars.length - suffix).join('');
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
  const lines = newlines === 'break' ? [replacement] : replacement.split('\n');
  const inserted = lines.map((line, index) => {
    const p = copy(firstParagraph);
    const pPr = firstChildElement(p, name('pPr'));
    const endPr = firstChildElement(
      index === lines.length - 1 ? lastParagraph : firstParagraph,
      name('endParaRPr'),
    );
    const content =
      line ||
      ((index !== 0 || left.length === 0) && (index !== lines.length - 1 || right.length === 0))
        ? line
            .split('\n')
            .flatMap((part, partIndex) => [
              ...(partIndex
                ? [elem(name('br'), { children: properties ? [copy(properties)] : [] })]
                : []),
              ...(part || !line ? [run(part, properties)] : []),
            ])
        : [];
    p.children = [
      ...(pPr ? [pPr] : []),
      ...joinAtSeam(
        joinAtSeam(index === 0 ? left : [], content),
        index === lines.length - 1 ? right : [],
      ),
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
  reset = false,
): void {
  applyRunFormat(elem(name('rPr')), format);
  mutateTextBodyRangeProperties(txBody, range, (properties) => {
    if (reset) resetRunFormat(properties);
    applyRunFormat(properties, format);
  });
}

/** Format one paragraph's end mark without rebuilding its runs or paragraph XML. */
export function formatTextBodyParagraphEnd(
  txBody: XmlElement,
  paragraphIndex: number,
  format: TextFormat,
  reset = false,
): void {
  const paragraph = paragraphsOf(txBody)[paragraphIndex];
  if (!paragraph) throw new RangeError(`paragraph index out of range: ${paragraphIndex}`);
  const existing = firstChildElement(paragraph, name('endParaRPr'));
  // Apply to a copy so invalid formatting cannot partially change the document.
  const properties = existing ? copy(existing) : elem(name('endParaRPr'));
  if (reset) resetRunFormat(properties);
  applyRunFormat(properties, format);
  if (existing) paragraph.children[paragraph.children.indexOf(existing)] = properties;
  else paragraph.children.push(properties);
}

/** Split boundary runs and mutate only the selected characters' properties. */
export function mutateTextBodyRangeProperties(
  txBody: XmlElement,
  range: { start: number; end: number },
  mutate: (properties: XmlElement) => void,
): void {
  validateTextRange(textBodyText(txBody), range, 'text range');
  const { start, end } = range;
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
        if (is(child, 'fld')) return textFragment(child, content.slice(a, b), true);
        return textFragment(child, content.slice(a, b));
      };
      if (from > 0) children.push(fragment(0, from));
      const selected = fragment(from, to);
      let properties = firstChildElement(selected, name('rPr'));
      if (!properties) {
        properties = elem(name('rPr'));
        selected.children.unshift(properties);
      }
      mutate(properties);
      children.push(selected);
      if (to < content.length) children.push(fragment(to, content.length));
    }
    paragraph.children = children;
    offset++; // The visible paragraph separator occupies one UTF-16 position.
  }
  txBody.children = updated.children;
}

export function validateTextRange(
  value: string,
  range: { start: number; end: number },
  caller: string,
): void {
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
    throw new RangeError(`${caller}: range must contain valid UTF-16 boundaries within the text`);
  }
}

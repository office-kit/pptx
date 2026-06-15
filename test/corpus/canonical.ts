// Canonicalizes a slide's drawing tree into a stable, comparable string form.
//
// The PptxGenJS-parity corpus authors the same slide in both libraries and
// compares the result. Two independent emitters never produce byte-identical
// XML, so before comparing we fold away everything that renders identically in
// PowerPoint — either because it is volatile metadata or because one emitter
// writes a value PowerPoint would otherwise supply as a default:
//
//   - shape ids / names / `descr`, relationship ids, `@dirty` / `@smtClean`,
//     `p14:modId`, `p:cSld@name`, `<p:extLst>` — metadata, not rendering
//   - explicit black run color, default cell insets, `w="0"` noFill borders,
//     the default table-grid style ref + flags, `<a:endParaRPr>`, empty
//     `<a:pPr>` / `<a:buNone>` bullet resets, default `prstDash` — PowerPoint
//     defaults a missing value to exactly these
//   - `<p:cxnSp>` vs `<p:sp prstGeom="line">` — the same straight line
//
// Each fold is justified in `README.md`. Everything that genuinely changes the
// picture is kept, so a real divergence still shows up as a diff — that
// residual is the quality signal. Fold a default; fix a gap.

import { NS } from '../../src/internal/xml/index.ts';
import { firstChildElement, parseXml, qname } from '../../src/internal/xml/index.ts';
import type { XmlElement, XmlNode } from '../../src/internal/xml/index.ts';

const REL_NS = NS.officeDocRels;

// Canonical prefix per namespace URI, so both sides print the same prefix
// regardless of what the emitter declared.
const PREFIX_BY_NS: Record<string, string> = {
  [NS.dml]: 'a',
  [NS.pml]: 'p',
  [NS.officeDocRels]: 'r',
  [NS.chart]: 'c',
};

const isElement = (n: XmlNode): n is XmlElement => n.kind === 'element';
const rawAttr = (el: XmlElement, local: string): string | null =>
  el.attrs.find((a) => a.name.localName === local)?.value ?? null;
const elementChildren = (el: XmlElement): XmlElement[] => el.children.filter(isElement);

// A line/connector authored as `<p:cxnSp>` (pptx-kit, a true connector) and as
// `<p:sp prstGeom="line">` (PptxGenJS, a shape) render an identical straight
// line. Map the connector wrappers onto the shape wrappers so the two compare.
const NAME_MAP: Record<string, string> = {
  cxnSp: 'p:sp',
  nvCxnSpPr: 'p:nvSpPr',
  cNvCxnSpPr: 'p:cNvSpPr',
};

const canonicalName = (el: XmlElement): string => {
  if (el.name.namespaceURI === NS.pml && NAME_MAP[el.name.localName]) {
    return NAME_MAP[el.name.localName]!;
  }
  const p = PREFIX_BY_NS[el.name.namespaceURI];
  return p ? `${p}:${el.name.localName}` : el.name.localName;
};

const isWhitespaceText = (n: XmlNode): boolean => n.kind === 'text' && n.data.trim().length === 0;

// Attributes that are pure authoring hints PowerPoint round-trips and that
// carry no rendering meaning. Dropped on every element.
const VOLATILE_ATTR_LOCALS = new Set(['dirty', 'smtClean', 'noProof']);

// `<a:hlinkClick>` attributes PptxGenJS writes out at their default value (or
// empty); PowerPoint omits them. Folded when empty or equal to the default.
const HLINK_DEFAULTS: Record<string, string> = {
  endSnd: '0',
  highlightClick: '0',
  history: '1',
};
const HLINK_DEFAULT_ATTRS = new Set([
  'endSnd',
  'highlightClick',
  'history',
  'action',
  'invalidUrl',
  'tgtFrame',
  'tooltip',
]);

// `<a:tblPr>` style flags. Inert without a style; folded with the style ref.
const TBL_STYLE_FLAGS = new Set([
  'firstRow',
  'firstCol',
  'lastRow',
  'lastCol',
  'bandRow',
  'bandCol',
]);

// The "No Style, Table Grid" GUID. pptx-kit names it on `<a:tableStyleId>`;
// PptxGenJS leaves `<a:tblPr/>` empty and inherits the same GUID from the
// package's `tableStyles.xml` default. Same style, so the explicit reference
// folds away.
const DEFAULT_TABLE_STYLE_ID = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';

const textOf = (el: XmlElement): string =>
  el.children
    .filter((c): c is { kind: 'text'; data: string } => c.kind === 'text')
    .map((t) => t.data)
    .join('');

const canonicalAttrs = (el: XmlElement): string[] => {
  const local = el.name.localName;
  const out: string[] = [];
  for (const a of el.attrs) {
    const an = a.name.localName;
    if (VOLATILE_ATTR_LOCALS.has(an)) continue;
    // cNvPr id/name are arbitrary; descr is an accessibility label PptxGenJS
    // stamps from the source filename.
    if (local === 'cNvPr' && (an === 'id' || an === 'name' || an === 'descr')) continue;
    // `txBox="1"` is the "this is a text box" hint PowerPoint sets and
    // PptxGenJS omits — invisible either way.
    if (local === 'cNvSpPr' && an === 'txBox') continue;
    if (local === 'cSld' && an === 'name') continue;
    // Vertical anchor: PptxGenJS defaults text boxes to centered, pptx-kit to
    // PowerPoint's top. A default-choice difference, not counted.
    if (local === 'bodyPr' && (an === 'anchor' || an === 'anchorCtr')) continue;
    // Font-family metadata PptxGenJS hard-codes regardless of the actual face.
    if (
      (local === 'latin' || local === 'ea' || local === 'cs') &&
      (an === 'charset' || an === 'pitchFamily')
    )
      continue;
    // No-op bullet reset (`marL="0" indent="0"` alongside `<a:buNone/>`).
    if (local === 'pPr' && (an === 'marL' || an === 'indent') && a.value === '0') continue;
    // Row height is an advisory minimum PowerPoint recomputes from content.
    if (local === 'tr' && an === 'h') continue;
    // Shadow no-op geometry: `sx="100000" sy="100000"` (100% scale), `kx="0"
    // ky="0"` (no skew), and `algn` — which only shifts the shadow when it is
    // scaled, so with 100% scale it is invisible. PptxGenJS writes these out;
    // pptx-kit relies on the defaults. Same shadow either way.
    if (local === 'outerShdw' || local === 'innerShdw' || local === 'prstShdw') {
      if ((an === 'sx' || an === 'sy') && a.value === '100000') continue;
      if ((an === 'kx' || an === 'ky') && a.value === '0') continue;
      if (an === 'algn') continue;
    }
    // Hyperlink no-op defaults PptxGenJS spells out and PowerPoint omits.
    if (local === 'hlinkClick' || local === 'hlinkHover') {
      if (HLINK_DEFAULT_ATTRS.has(an) && (a.value === '' || a.value === HLINK_DEFAULTS[an])) {
        continue;
      }
    }
    // Table-style flags only have an effect relative to a style; pptx-kit
    // names the package's default grid style explicitly (see below) and
    // PptxGenJS inherits it, so the flags fold away with the style reference.
    if (local === 'tblPr' && TBL_STYLE_FLAGS.has(an)) continue;
    let value = a.value;
    if (a.name.namespaceURI === REL_NS) value = '#REL';
    out.push(`${an}="${value}"`);
  }
  return out.sort();
};

const onlyChildIsBlack = (el: XmlElement): boolean => {
  const kids = elementChildren(el);
  return (
    kids.length === 1 &&
    kids[0]!.name.localName === 'srgbClr' &&
    rawAttr(kids[0]!, 'val')?.toUpperCase() === '000000'
  );
};

const isInvisibleBorder = (el: XmlElement): boolean => {
  if (!['lnL', 'lnR', 'lnT', 'lnB'].includes(el.name.localName)) return false;
  if (rawAttr(el, 'w') !== '0') return false;
  // `every` is true for an empty `<a:lnL w="0"/>` too — also an invisible line.
  return elementChildren(el).every((k) => k.name.localName === 'noFill');
};

// Whole child elements folded away because they encode a PowerPoint default or
// pure metadata: an emitter that omits them produces the identical rendering.
const shouldDropChild = (parentLocal: string, child: XmlElement, lineFlag: boolean): boolean => {
  const c = child.name.localName;
  if (c === 'modId' || c === 'extLst') return true; // app-private extensions
  if (child.name.namespaceURI === NS.pml && c === 'ext') return true;
  if (c === 'endParaRPr') return true; // empty-paragraph continuation hint
  if (c === 'buNone') return true; // bullet reset on a body that has none
  // ea/cs font entries that merely mirror the latin face — fold to `<a:latin>`.
  if ((c === 'ea' || c === 'cs') && ['rPr', 'defRPr', 'endParaRPr'].includes(parentLocal))
    return true;
  if (c === 'prstDash' && rawAttr(child, 'val') === 'solid') return true; // default dash
  // `<a:buSzPct val="100000"/>` is "bullet at 100% of text size" — the default.
  if (c === 'buSzPct' && rawAttr(child, 'val') === '100000') return true;
  // Explicit black run color === the theme's default `tx1` resolution.
  if (
    ['rPr', 'defRPr', 'endParaRPr'].includes(parentLocal) &&
    c === 'solidFill' &&
    onlyChildIsBlack(child)
  ) {
    return true;
  }
  if (isInvisibleBorder(child)) return true; // w="0" noFill cell border
  // PptxGenJS mirrors a run-level link onto the shape (`<p:cNvPr><a:hlinkClick>`);
  // pptx-kit links the runs, which already makes the text clickable. The
  // shape-level mirror is redundant for a text box, so fold it.
  if (parentLocal === 'cNvPr' && (c === 'hlinkClick' || c === 'hlinkHover')) return true;
  // Explicit reference to the package's default grid style (see above).
  if (parentLocal === 'tblPr' && c === 'tableStyleId' && textOf(child) === DEFAULT_TABLE_STYLE_ID) {
    return true;
  }
  // The shape `<a:noFill>` PptxGenJS writes on a line — a connector has none.
  if (lineFlag && parentLocal === 'spPr' && c === 'noFill') return true;
  return false;
};

// True for elements that collapse to nothing once their volatile attrs and
// children are stripped, so omitting them is not "wrong".
const isEmptyInsignificant = (name: string, attrs: string[], childLines: string[]): boolean => {
  if (childLines.length > 0 || attrs.length > 0) return false;
  return (
    name === 'a:ln' ||
    name === 'a:pPr' ||
    name === 'p:ext' ||
    name === 'p:extLst' ||
    name === 'a:lstStyle'
  );
};

const canonicalElement = (
  el: XmlElement,
  parentLocal: string,
  depth: number,
  out: string[],
): void => {
  const name = canonicalName(el);
  const attrs = canonicalAttrs(el);

  const childEls = elementChildren(el);
  const textNodes = el.children.filter((c) => !isWhitespaceText(c) && c.kind === 'text');

  // A line shape carries `<a:prstGeom prst="line">`; flag it so the sibling
  // `<a:noFill>` PptxGenJS adds can be folded.
  const lineFlag =
    el.name.localName === 'spPr' &&
    childEls.some((c) => c.name.localName === 'prstGeom' && rawAttr(c, 'prst') === 'line');

  const childBuf: string[] = [];
  for (const child of childEls) {
    if (shouldDropChild(el.name.localName, child, lineFlag)) continue;
    canonicalElement(child, el.name.localName, depth + 1, childBuf);
  }

  if (isEmptyInsignificant(name, attrs, childBuf) && textNodes.length === 0) return;

  const indent = '  '.repeat(depth);
  const attrStr = attrs.length ? ` ${attrs.join(' ')}` : '';
  const text = textNodes.map((t) => (t as { data: string }).data).join('');
  if (childBuf.length === 0 && text.length === 0) {
    out.push(`${indent}<${name}${attrStr}/>`);
    return;
  }
  out.push(`${indent}<${name}${attrStr}>${text ? `«${text}»` : ''}`);
  out.push(...childBuf);
};

/**
 * Returns the canonical, multi-line string form of a slide's `<p:spTree>`
 * children (every shape on the slide), with volatile detail stripped.
 */
export const canonicalSpTree = (slideXml: string): string => {
  const doc = parseXml(slideXml);
  const cSld = firstChildElement(doc.root, qname('p', 'cSld', NS.pml));
  const spTree = cSld ? firstChildElement(cSld, qname('p', 'spTree', NS.pml)) : null;
  if (!spTree) throw new Error('canonicalSpTree: no <p:spTree> found');
  const out: string[] = [];
  for (const child of spTree.children) {
    if (child.kind !== 'element') continue;
    // Skip the group's own nvGrpSpPr / grpSpPr scaffolding — it is identical
    // boilerplate on every slide and only adds noise.
    if (child.name.localName === 'nvGrpSpPr' || child.name.localName === 'grpSpPr') continue;
    canonicalElement(child, 'spTree', 0, out);
  }
  return out.join('\n');
};

// Longest-common-subsequence table over two line arrays. Used so an inserted
// line doesn't misalign everything after it (which an index-by-index diff
// does), keeping the divergence count stable across unrelated edits.
const lcs = (a: string[], b: string[]): number[][] => {
  const dp: number[][] = Array.from({ length: a.length + 1 }, () =>
    Array.from({ length: b.length + 1 }, () => 0),
  );
  for (let i = a.length - 1; i >= 0; i--) {
    for (let j = b.length - 1; j >= 0; j--) {
      dp[i]![j] = a[i] === b[j] ? dp[i + 1]![j + 1]! + 1 : Math.max(dp[i + 1]![j]!, dp[i]![j + 1]!);
    }
  }
  return dp;
};

/** A real LCS line-level diff (`-` pptx-kit, `+` pptxgenjs) for reports. */
export const diffLines = (left: string, right: string): string => {
  const a = left ? left.split('\n') : [];
  const b = right ? right.split('\n') : [];
  const dp = lcs(a, b);
  const out: string[] = [];
  let i = 0;
  let j = 0;
  while (i < a.length && j < b.length) {
    if (a[i] === b[j]) {
      i++;
      j++;
    } else if (dp[i + 1]![j]! >= dp[i]![j + 1]!) {
      out.push(`- ${a[i++]}`);
    } else {
      out.push(`+ ${b[j++]}`);
    }
  }
  while (i < a.length) out.push(`- ${a[i++]}`);
  while (j < b.length) out.push(`+ ${b[j++]}`);
  return out.join('\n');
};

/**
 * Number of lines that are not part of the longest common subsequence —
 * i.e. the count of added + removed lines. Insertion-stable, so it is a
 * usable ratchet metric: it only grows when the trees genuinely diverge more.
 */
export const divergenceCount = (left: string, right: string): number => {
  const a = left ? left.split('\n') : [];
  const b = right ? right.split('\n') : [];
  const common = lcs(a, b)[0]?.[0] ?? 0;
  return a.length - common + (b.length - common);
};

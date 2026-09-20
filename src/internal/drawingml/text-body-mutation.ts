// Mutating helpers for `a:txBody`. Companion to `text-body.ts`, which is
// read-only.
//
// Why a separate file: the read-only path is hot (called on every slide
// load), the mutation path runs only when the user calls a `set*` method.
// Keeping them separate keeps `text-body.ts` small and clearly side-effect-
// free.

import { oneOf } from '../bounds.ts';
import { AUTO_NUMBER_SCHEMES } from '../enum-values.ts';

import {
  NS,
  type XmlAttr,
  type XmlElement,
  type XmlNode,
  attr,
  elem,
  firstChildElement,
  qname,
  text,
  walkElements,
} from '../xml/index.ts';
import { type TextFormat, applyRunFormat } from './text-format.ts';

const NAME_BU_CHAR = qname('a', 'buChar', NS.dml);
const NAME_BU_AUTO_NUM = qname('a', 'buAutoNum', NS.dml);
const NAME_BU_NONE = qname('a', 'buNone', NS.dml);
const NAME_BU_FONT = qname('a', 'buFont', NS.dml);
const ATTR_CHAR = qname('', 'char', '');
const ATTR_BU_TYPE = qname('', 'type', '');
const ATTR_START_AT = qname('', 'startAt', '');
const ATTR_TYPEFACE = qname('', 'typeface', '');

const NAME_P = qname('a', 'p', NS.dml);
const NAME_R = qname('a', 'r', NS.dml);
const NAME_T = qname('a', 't', NS.dml);
const NAME_RPR = qname('a', 'rPr', NS.dml);
const NAME_PPR = qname('a', 'pPr', NS.dml);
const NAME_END_PARA_RPR = qname('a', 'endParaRPr', NS.dml);
const ATTR_XML_SPACE = qname('xml', 'space', NS.xml);

const cloneAttrs = (attrs: ReadonlyArray<XmlAttr>): XmlAttr[] =>
  attrs.map((a) => attr(a.name, a.value));

const cloneElement = (e: XmlElement): XmlElement =>
  elem(e.name, {
    attrs: cloneAttrs(e.attrs),
    prefixDecls: new Map(e.prefixDecls),
    children: e.children.map(cloneNode),
  });

const cloneNode = (n: XmlNode): XmlNode => {
  switch (n.kind) {
    case 'element':
      return cloneElement(n);
    case 'text':
      return { kind: 'text', data: n.data };
    case 'cdata':
      return { kind: 'cdata', data: n.data };
    case 'comment':
      return { kind: 'comment', data: n.data };
    case 'pi':
      return { kind: 'pi', target: n.target, data: n.data };
  }
};

/**
 * Returns the first `a:rPr` reachable on the first run of the first paragraph
 * inside `txBody`. Used to preserve formatting on text replacement: we clone
 * the original rPr into the new run so font, color, size, etc. survive.
 */
const findFirstRunProperties = (txBody: XmlElement): XmlElement | null => {
  const firstP = firstChildElement(txBody, NAME_P);
  if (firstP === null) return null;
  const firstR = firstChildElement(firstP, NAME_R);
  if (firstR === null) return null;
  return firstChildElement(firstR, NAME_RPR);
};

/**
 * Returns the first `a:pPr` from the first paragraph. Used to preserve
 * paragraph-level formatting (alignment, bullet, indent) when replacing.
 */
const findFirstParagraphProperties = (txBody: XmlElement): XmlElement | null => {
  const firstP = firstChildElement(txBody, NAME_P);
  if (firstP === null) return null;
  return firstChildElement(firstP, NAME_PPR);
};

const removeAllParagraphs = (txBody: XmlElement): void => {
  txBody.children = txBody.children.filter(
    (c) => !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'p'),
  );
};

// Token pattern: `{{key}}` where `key` matches any non-brace characters.
// This is the same syntax docxtemplater / Handlebars use, and it survives
// XML escaping cleanly since `{` and `}` are not XML metacharacters.
const TOKEN_PATTERN = /\{\{([^{}]+)\}\}/g;

/**
 * Replaces `{{key}}` tokens in every `a:t` element under `root` with values
 * from `tokens`. Tokens whose key is not in `tokens` are left untouched.
 *
 * Returns the number of substitutions performed. Useful for callers that
 * want to know whether anything matched.
 *
 * Limitation: a token must fit entirely within one `<a:t>` element to
 * match. PowerPoint normally serializes contiguous user text as a single
 * `<a:t>`, so the limitation only bites when a placeholder was edited
 * character-by-character (causing PowerPoint to split runs). For those
 * cases, fall back to `setText()`.
 */
export const replaceTokensInTree = (root: XmlElement, tokens: Record<string, string>): number => {
  let count = 0;
  walkElements(root, (el) => {
    if (el.name.namespaceURI !== NAME_T.namespaceURI) return;
    if (el.name.localName !== 't') return;
    const child = el.children[0];
    if (!child || child.kind !== 'text') return;
    const before = child.data;
    let didMatch = false;
    const after = before.replace(TOKEN_PATTERN, (match, key: string) => {
      if (Object.hasOwn(tokens, key)) {
        didMatch = true;
        return tokens[key] ?? '';
      }
      return match;
    });
    if (didMatch) {
      child.data = after;
      count++;
    }
  });
  return count;
};

/**
 * Replaces every occurrence of `from` in every `<a:t>` element under
 * `root` with `to`. `from` may be a string (literal) or a `RegExp`
 * (matched per-run). Returns the count of `<a:t>` elements that were
 * mutated — not the count of substitutions inside them, since callers
 * usually want "did anything change" rather than "how many letters
 * moved."
 *
 * Same single-run constraint as `replaceTokensInTree`: matches must
 * fit inside one `<a:t>`. Use this for the broader "find/replace
 * across the deck" use case where {{token}} syntax isn't a fit.
 */
export const replaceTextInTree = (root: XmlElement, from: string | RegExp, to: string): number => {
  // Build a global RegExp so .replace() touches every occurrence per run.
  const pattern =
    typeof from === 'string'
      ? new RegExp(from.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g')
      : from.global
        ? from
        : new RegExp(from.source, `${from.flags}g`);
  let count = 0;
  walkElements(root, (el) => {
    if (el.name.namespaceURI !== NAME_T.namespaceURI) return;
    if (el.name.localName !== 't') return;
    const child = el.children[0];
    if (!child || child.kind !== 'text') return;
    const before = child.data;
    const after = before.replace(pattern, to);
    if (after !== before) {
      child.data = after;
      count++;
    }
  });
  return count;
};

/**
 * Bullet style descriptor for a paragraph.
 *
 *   - `'bullet'` — shortcut for `{ char: '•' }`.
 *   - `'number'` — shortcut for `{ autoNum: 'arabicPeriod' }` (1., 2., 3.).
 *   - `'none'` — emits `<a:buNone/>`, forcing the paragraph to be bullet-free
 *     even if the layout has a bullet default.
 *   - `{ char }` — custom bullet character (any single grapheme).
 *   - `{ autoNum }` — any ECMA-376 `ST_TextAutonumberScheme` token
 *     (`arabicPeriod`, `romanLcParenR`, `alphaUcPeriod`, ...).
 */
export type BulletStyle = 'bullet' | 'number' | 'none' | { char: string } | { autoNum: string };

const normalizeBulletStyle = (
  s: BulletStyle,
  caller: string,
): { kind: 'char'; char: string } | { kind: 'autoNum'; type: string } | { kind: 'none' } => {
  if (typeof s === 'string') oneOf(s, ['bullet', 'number', 'none'], `${caller}: bullets`);
  if (s === 'bullet') return { kind: 'char', char: '•' };
  if (s === 'number') return { kind: 'autoNum', type: 'arabicPeriod' };
  if (s === 'none') return { kind: 'none' };
  if ('char' in s) return { kind: 'char', char: s.char };
  return {
    kind: 'autoNum',
    type: oneOf(s.autoNum, AUTO_NUMBER_SCHEMES, `${caller}: bullets.autoNum`),
  };
};

type NormalizedBullet = ReturnType<typeof normalizeBulletStyle>;

const buildBulletElement = (n: NormalizedBullet): XmlElement => {
  switch (n.kind) {
    case 'char':
      return elem(NAME_BU_CHAR, { attrs: [attr(ATTR_CHAR, n.char)] });
    case 'autoNum':
      // `startAt="1"` is the default, but PowerPoint and PptxGenJS write it
      // explicitly on an authored numbered list.
      return elem(NAME_BU_AUTO_NUM, {
        attrs: [attr(ATTR_START_AT, '1'), attr(ATTR_BU_TYPE, n.type)],
      });
    case 'none':
      return elem(NAME_BU_NONE);
  }
};

const NAME_PPR_FOR_BULLET = qname('a', 'pPr', NS.dml);
const NAME_P_FOR_BULLET = qname('a', 'p', NS.dml);

/**
 * Paragraph horizontal alignment per `ST_TextAlignType` (ECMA-376
 * §17.18.59). The library accepts plain-English names and translates to
 * the spec tokens; advanced callers can pass the raw token directly.
 */
export type ParagraphAlignment =
  | 'left'
  | 'center'
  | 'right'
  | 'justify'
  | 'distribute'
  | 'l'
  | 'ctr'
  | 'r'
  | 'just'
  | 'dist'
  | 'justLow'
  | 'thaiDist';

export const alignToken = (a: ParagraphAlignment, caller: string): string => {
  oneOf(
    a,
    [
      'left',
      'center',
      'right',
      'justify',
      'distribute',
      'l',
      'ctr',
      'r',
      'just',
      'dist',
      'justLow',
      'thaiDist',
    ],
    `${caller}: align`,
  );
  switch (a) {
    case 'left':
    case 'l':
      return 'l';
    case 'center':
    case 'ctr':
      return 'ctr';
    case 'right':
    case 'r':
      return 'r';
    case 'justify':
    case 'just':
      return 'just';
    case 'distribute':
    case 'dist':
      return 'dist';
    default:
      return a;
  }
};

const ATTR_ALGN = qname('', 'algn', '');

/**
 * Sets `<a:pPr algn="...">` on every paragraph in `txBody`. Existing
 * pPr attributes other than `algn` are preserved.
 */
export const applyAlignmentToAllParagraphs = (
  txBody: XmlElement,
  align: ParagraphAlignment,
  caller = 'setShapeAlignment',
): void => {
  applyAlignmentTokenToAllParagraphs(txBody, alignToken(align, caller));
};

export const applyAlignmentTokenToAllParagraphs = (txBody: XmlElement, token: string): void => {
  for (const p of txBody.children) {
    if (
      p.kind !== 'element' ||
      p.name.namespaceURI !== NAME_P_FOR_BULLET.namespaceURI ||
      p.name.localName !== 'p'
    ) {
      continue;
    }
    let pPr = firstChildElement(p, NAME_PPR_FOR_BULLET);
    if (pPr === null) {
      pPr = elem(NAME_PPR_FOR_BULLET);
      p.children.unshift(pPr);
    }
    pPr.attrs = pPr.attrs.filter((a) => a.name.localName !== 'algn');
    pPr.attrs.push(attr(ATTR_ALGN, token));
  }
};

/**
 * Sets the bullet style on a single `<a:p>` paragraph. Drops any
 * existing bullet child element (`a:buChar`, `a:buAutoNum`, `a:buNone`)
 * before inserting the new one. Creates `<a:pPr>` if absent.
 */
// PowerPoint's default hanging indent for a bulleted paragraph, by list level
// (0-based). These mirror the master `bodyStyle` lvlNpPr defaults, so a bullet
// authored on a *text box* — which inherits the `otherStyle` (marL=0), not the
// body style — still gets the standard bullet/text gap instead of the glyph
// jammed against the text. PptxGenJS and PowerPoint both write these explicitly.
const BULLET_INDENT_BY_LEVEL: ReadonlyArray<{ marL: number; indent: number }> = [
  { marL: 342900, indent: -342900 },
  { marL: 742950, indent: -285750 },
  { marL: 1143000, indent: -228600 },
  { marL: 1600200, indent: -228600 },
  { marL: 2057400, indent: -228600 },
];
const bulletIndentForLevel = (lvl: number): { marL: number; indent: number } =>
  BULLET_INDENT_BY_LEVEL[Math.min(lvl, BULLET_INDENT_BY_LEVEL.length - 1)]!;

const ATTR_MAR_L = qname('', 'marL', '');
const ATTR_INDENT = qname('', 'indent', '');

export const updateBulletIndentForLevel = (
  pPr: XmlElement,
  previousLevel: number,
  level: number,
): void => {
  const previous = bulletIndentForLevel(previousLevel);
  const marL = pPr.attrs.find((a) => a.name.namespaceURI === '' && a.name.localName === 'marL');
  const indent = pPr.attrs.find((a) => a.name.namespaceURI === '' && a.name.localName === 'indent');
  if (marL?.value !== String(previous.marL) || indent?.value !== String(previous.indent)) return;
  const next = bulletIndentForLevel(level);
  pPr.attrs = pPr.attrs.map((a) =>
    a === marL
      ? attr(a.name, String(next.marL))
      : a === indent
        ? attr(a.name, String(next.indent))
        : a,
  );
};

const hasAttr = (el: XmlElement, local: string): boolean =>
  el.attrs.some((a) => a.name.namespaceURI === '' && a.name.localName === local);

const applyNormalizedBullet = (paragraph: XmlElement, style: NormalizedBullet): void => {
  let pPr = firstChildElement(paragraph, NAME_PPR_FOR_BULLET);
  if (pPr === null) {
    pPr = elem(NAME_PPR_FOR_BULLET);
    // <a:pPr> must be the first child of <a:p>.
    paragraph.children.unshift(pPr);
  }
  // Remove any existing bullet child (and the number font we may have added).
  pPr.children = pPr.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        (c.name.localName === 'buChar' ||
          c.name.localName === 'buAutoNum' ||
          c.name.localName === 'buNone' ||
          c.name.localName === 'buFont')
      ),
  );

  // A visible bullet needs a hanging indent or the glyph overlaps the text.
  // Only fill it in when the caller hasn't set their own marL / indent, and
  // never for `none` (which removes the bullet). marL / indent are pPr
  // ATTRIBUTES, so they must come before the bullet child element.
  if (style.kind !== 'none') {
    const lvl = Number.parseInt(
      pPr.attrs.find((a) => a.name.localName === 'lvl')?.value ?? '0',
      10,
    );
    const { marL, indent } = bulletIndentForLevel(Number.isFinite(lvl) ? lvl : 0);
    if (!hasAttr(pPr, 'marL')) pPr.attrs.push(attr(ATTR_MAR_L, String(marL)));
    if (!hasAttr(pPr, 'indent')) pPr.attrs.push(attr(ATTR_INDENT, String(indent)));
  }

  // A numbered list needs a bullet font so the number glyph has a face —
  // PowerPoint and PptxGenJS emit `<a:buFont typeface="+mj-lt"/>` (the theme's
  // major font) ahead of `<a:buAutoNum>`. A character bullet carries its glyph
  // directly and needs none. `<a:buFont>` precedes the bullet child per the
  // CT_TextParagraphProperties element order.
  if (style.kind === 'autoNum') {
    pPr.children.push(elem(NAME_BU_FONT, { attrs: [attr(ATTR_TYPEFACE, '+mj-lt')] }));
  }
  pPr.children.push(buildBulletElement(style));
};

export const applyBulletToParagraph = (paragraph: XmlElement, style: BulletStyle): void => {
  applyNormalizedBullet(paragraph, normalizeBulletStyle(style, 'setParagraphBullet'));
};

export const applyBulletToAllParagraphs = (txBody: XmlElement, style: BulletStyle): void => {
  const normalized = normalizeBulletStyle(style, 'setShapeBullets');
  for (const p of txBody.children) {
    if (
      p.kind !== 'element' ||
      p.name.namespaceURI !== NAME_P_FOR_BULLET.namespaceURI ||
      p.name.localName !== 'p'
    ) {
      continue;
    }
    applyNormalizedBullet(p, normalized);
  }
};

/**
 * Replaces the entire visible text of a `txBody` element with `value`.
 * Newlines in `value` start a new paragraph.
 *
 * The first existing `a:rPr` (run properties) and `a:pPr` (paragraph
 * properties) are cloned into every new paragraph so that font, color, size,
 * alignment, and bullet style survive the replacement. If the source had
 * mixed formatting per run, that gets collapsed into the first-run formatting
 * across the entire new text — matching what PowerPoint does when you select
 * all and type.
 *
 * The `bodyPr` and `lstStyle` children (if any) are preserved untouched.
 */
export const setTextBody = (txBody: XmlElement, value: string, bullets?: BulletStyle): void => {
  const normalized =
    bullets === undefined ? undefined : normalizeBulletStyle(bullets, 'setShapeText');
  const rPrTemplate = findFirstRunProperties(txBody);
  const pPrTemplate = findFirstParagraphProperties(txBody);

  removeAllParagraphs(txBody);

  const lines = value.split('\n');
  for (const line of lines) {
    // Per the strict ECMA schema, `<a:t>` does NOT accept `xml:space`. We
    // split on `\n` so each `<a:t>` holds a single line and leading /
    // trailing whitespace is handled by the body / lst style, not by an
    // illegal attribute on the text element.
    const t = elem(NAME_T, {
      children: line.length > 0 ? [text(line)] : [],
    });
    const r = elem(NAME_R, {
      children: rPrTemplate !== null ? [cloneElement(rPrTemplate), t] : [t],
    });
    const p = elem(NAME_P, {
      children: pPrTemplate !== null ? [cloneElement(pPrTemplate), r] : [r],
    });
    if (normalized !== undefined) applyNormalizedBullet(p, normalized);
    txBody.children.push(p);
  }
  void ATTR_XML_SPACE;
};

/** One run of a `ParagraphSpec`: text plus the format applied to its `<a:rPr>`. */
export interface RunSpec {
  /**
   * Run text, written verbatim — unlike `setShapeText`, newlines are not
   * split into paragraphs (PowerPoint shows a CR LF inside a run as a line
   * break, which some exporters rely on).
   */
  readonly text: string;
  readonly format?: TextFormat;
}

/** One paragraph for `setTextBodyParagraphs`: optional alignment plus its runs. */
export interface ParagraphSpec {
  readonly align?: ParagraphAlignment;
  readonly runs: ReadonlyArray<RunSpec>;
  /**
   * Format of the paragraph-end mark (`<a:endParaRPr>`), the place to give a
   * paragraph with no runs a font size. Omit to write no `<a:endParaRPr>`.
   */
  readonly endFormat?: TextFormat;
}

/**
 * Builds the `<a:p>` elements for `paragraphs`. Every format is validated
 * here, so a caller that builds before it touches the tree leaves the
 * existing text intact when a value is rejected.
 */
export const buildTextBodyParagraphs = (
  paragraphs: ReadonlyArray<ParagraphSpec>,
  caller = 'setShapeParagraphs',
): ReadonlyArray<XmlElement> => {
  // CT_TextBody requires at least one <a:p>; `[{ runs: [] }]` is the empty body.
  if (paragraphs.length === 0) {
    throw new Error('setTextBodyParagraphs: at least one paragraph is required');
  }
  const built: XmlElement[] = [];
  // for...of, not map: map skips the holes of a sparse array and would hand
  // them on as undefined children.
  for (const para of paragraphs) {
    const children: XmlElement[] = [];
    if (para.align !== undefined) {
      children.push(elem(NAME_PPR, { attrs: [attr(ATTR_ALGN, alignToken(para.align, caller))] }));
    }
    for (const run of para.runs) {
      const rPr = elem(NAME_RPR);
      if (run.format !== undefined) applyRunFormat(rPr, run.format, caller);
      const t = elem(NAME_T, { children: run.text.length > 0 ? [text(run.text)] : [] });
      children.push(elem(NAME_R, { children: [rPr, t] }));
    }
    // CT_TextParagraph is a sequence: <a:endParaRPr> comes after every run.
    if (para.endFormat !== undefined) {
      const endParaRPr = elem(NAME_END_PARA_RPR);
      applyRunFormat(endParaRPr, para.endFormat, caller);
      children.push(endParaRPr);
    }
    built.push(elem(NAME_P, { children }));
  }
  return built;
};

export const replaceTextBodyParagraphs = (
  txBody: XmlElement,
  built: ReadonlyArray<XmlElement>,
): void => {
  removeAllParagraphs(txBody);
  txBody.children.push(...built);
};

/**
 * Replaces every paragraph of `txBody` with explicitly structured ones.
 * Unlike `setTextBody`, nothing is inherited from the existing runs: each
 * run gets a fresh `<a:rPr>` carrying only its own `format`. A rejected
 * format throws before `txBody` changes.
 */
export const setTextBodyParagraphs = (
  txBody: XmlElement,
  paragraphs: ReadonlyArray<ParagraphSpec>,
): void => {
  replaceTextBodyParagraphs(txBody, buildTextBodyParagraphs(paragraphs));
};

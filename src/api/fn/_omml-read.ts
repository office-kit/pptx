// Reads equations (OMML, ECMA-376 Part 1 §22.1) out of a DrawingML
// paragraph. The desktop app writes an equation as `<a14:m>` ([MS-ODRAWXML]
// §2.3.1.21) directly inside `<a:p>`, holding either one inline `<m:oMath>` or
// an `<m:oMathPara>` of display equations.
//
// Only the structures listed in `MathNode` are modeled. Anything else becomes
// an `unsupported` node that keeps its element name and its plain text, so a
// renderer can flag it instead of leaving a blank.

import { NS, firstChildElement, getAttrValue, qname } from '../../internal/xml/index.ts';
import type { XmlElement } from '../../internal/xml/index.ts';
import type { ReadTextFormat } from '../../internal/drawingml/index.ts';
import { parseRPrLikeElement } from './shape-color.ts';

/** A sequence of math nodes laid out left to right (an OMML argument). */
export type MathNodes = ReadonlyArray<MathNode>;

/**
 * Character style of a math run (`<m:sty>`). Omitted means the math default:
 * letters italic, digits and operators upright.
 */
export type MathTextStyle = 'plain' | 'bold' | 'italic' | 'boldItalic';

/** Where an n-ary operator puts its limits (`<m:limLoc>`). */
export type MathLimitLocation = 'underOver' | 'subSup';

/** Fraction layout (`<m:type>` of `<m:fPr>`). */
export type MathFractionType = 'bar' | 'skewed' | 'linear' | 'noBar';

/**
 * One node of an equation, modeled on the OMML elements of
 * ECMA-376 Part 1 §22.1.2. Character defaults the schema leaves implicit are
 * resolved here (a delimiter's `(` `)` `|`, an n-ary operator's `∫`, an
 * accent's combining circumflex), so readers never see an empty character.
 */
export type MathNode =
  /** `<m:r>`: literal text. `format` is the run's own `<a:rPr>`. */
  | {
      readonly kind: 'text';
      readonly text: string;
      readonly style?: MathTextStyle;
      readonly format: ReadTextFormat | null;
    }
  /** `<m:f>` */
  | {
      readonly kind: 'fraction';
      readonly type: MathFractionType;
      readonly numerator: MathNodes;
      readonly denominator: MathNodes;
    }
  /** `<m:sSub>` */
  | { readonly kind: 'subscript'; readonly base: MathNodes; readonly sub: MathNodes }
  /** `<m:sSup>` */
  | { readonly kind: 'superscript'; readonly base: MathNodes; readonly sup: MathNodes }
  /** `<m:sSubSup>` */
  | {
      readonly kind: 'subSuperscript';
      readonly base: MathNodes;
      readonly sub: MathNodes;
      readonly sup: MathNodes;
    }
  /** `<m:rad>`. `degree` is `null` for a square root (`<m:degHide>`). */
  | { readonly kind: 'radical'; readonly degree: MathNodes | null; readonly body: MathNodes }
  /**
   * `<m:nary>`: a big operator such as `∑` or `∫`. A hidden limit
   * (`<m:subHide>` / `<m:supHide>`) is `null`. `limitLocation` is omitted
   * when the equation leaves it to the application default.
   */
  | {
      readonly kind: 'nary';
      readonly operator: string;
      readonly limitLocation?: MathLimitLocation;
      readonly sub: MathNodes | null;
      readonly sup: MathNodes | null;
      readonly body: MathNodes;
    }
  /**
   * `<m:d>`: items between `open` and `close`, joined by `separator`. An
   * empty string is an omitted delimiter.
   */
  | {
      readonly kind: 'delimiter';
      readonly open: string;
      readonly close: string;
      readonly separator: string;
      readonly items: ReadonlyArray<MathNodes>;
    }
  /** `<m:m>`: rows of cells. */
  | { readonly kind: 'matrix'; readonly rows: ReadonlyArray<ReadonlyArray<MathNodes>> }
  /** `<m:acc>`: `accent` is the combining character placed over `body`. */
  | { readonly kind: 'accent'; readonly accent: string; readonly body: MathNodes }
  /** `<m:func>`: a function name such as `sin` applied to `body`. */
  | { readonly kind: 'function'; readonly name: MathNodes; readonly body: MathNodes }
  /** `<m:limLow>` / `<m:limUpp>`: `limit` under or over `base` (as in `lim`). */
  | {
      readonly kind: 'lowerLimit' | 'upperLimit';
      readonly base: MathNodes;
      readonly limit: MathNodes;
    }
  /**
   * An OMML element outside the modeled subset. `element` is its local name
   * (`eqArr`, `groupChr`, …) and `text` its run text in document order.
   */
  | { readonly kind: 'unsupported'; readonly element: string; readonly text: string };

const m = (localName: string) => qname('m', localName, NS.math);
const NAME_VAL = m('val');
const NAME_A_RPR = qname('a', 'rPr', NS.dml);

const STYLE_BY_STY: Readonly<Record<string, MathTextStyle>> = {
  p: 'plain',
  b: 'bold',
  i: 'italic',
  bi: 'boldItalic',
};

const FRACTION_TYPES: Readonly<Record<string, MathFractionType>> = {
  bar: 'bar',
  skw: 'skewed',
  lin: 'linear',
  noBar: 'noBar',
};

// §22.1.2: the characters a property element falls back to when it, or its
// `m:val`, is absent.
const DEFAULT_DELIMITER_OPEN = '(';
const DEFAULT_DELIMITER_CLOSE = ')';
const DEFAULT_DELIMITER_SEPARATOR = '|';
const DEFAULT_NARY_OPERATOR = '∫';
const DEFAULT_ACCENT = '\u0302';

const isMath = (node: XmlElement['children'][number]): node is XmlElement =>
  node.kind === 'element' && node.name.namespaceURI === NS.math;

const childOf = (parent: XmlElement | null, localName: string): XmlElement | null =>
  parent === null ? null : firstChildElement(parent, m(localName));

const childrenOf = (parent: XmlElement, localName: string): XmlElement[] =>
  parent.children.filter(
    (child): child is XmlElement => isMath(child) && child.name.localName === localName,
  );

const valOf = (parent: XmlElement | null, localName: string): string | null => {
  const el = childOf(parent, localName);
  return el === null ? null : getAttrValue(el, NAME_VAL);
};

// ST_OnOff (§22.9.2.7): a present element without `m:val` is on.
const flagOf = (parent: XmlElement | null, localName: string): boolean => {
  const el = childOf(parent, localName);
  if (el === null) return false;
  const value = getAttrValue(el, NAME_VAL);
  return value === null || value === 'on' || value === '1' || value === 'true';
};

// A character property (`m:chr`, `m:begChr`, …) present without `m:val`
// means "no character" (§22.1.2.20), unlike an absent one.
const charOf = (props: XmlElement | null, localName: string, fallback: string): string => {
  const el = childOf(props, localName);
  if (el === null) return fallback;
  return getAttrValue(el, NAME_VAL) ?? '';
};

const runText = (run: XmlElement): string => {
  let out = '';
  for (const t of childrenOf(run, 't')) {
    for (const c of t.children) if (c.kind === 'text' || c.kind === 'cdata') out += c.data;
  }
  return out;
};

/** Plain text of every `<m:t>` under `el`, in document order. */
export const mathPlainText = (el: XmlElement): string => {
  let out = '';
  for (const child of el.children) {
    if (child.kind !== 'element') continue;
    if (child.name.namespaceURI === NS.math && child.name.localName === 't') {
      for (const c of child.children) if (c.kind === 'text' || c.kind === 'cdata') out += c.data;
    } else {
      out += mathPlainText(child);
    }
  }
  return out;
};

const readArg = (parent: XmlElement, localName: string): MathNodes => {
  const el = childOf(parent, localName);
  return el === null ? [] : readMathNodes(el);
};

const readRun = (run: XmlElement): MathNode => {
  const sty = valOf(childOf(run, 'rPr'), 'sty');
  const aRPr = firstChildElement(run, NAME_A_RPR);
  const style = sty === null ? undefined : STYLE_BY_STY[sty];
  return {
    kind: 'text',
    text: runText(run),
    ...(style !== undefined ? { style } : {}),
    format: aRPr === null ? null : parseRPrLikeElement(aRPr),
  };
};

const readNode = (el: XmlElement): MathNode | null => {
  switch (el.name.localName) {
    case 'r':
      return readRun(el);
    case 'f':
      return {
        kind: 'fraction',
        type: FRACTION_TYPES[valOf(childOf(el, 'fPr'), 'type') ?? 'bar'] ?? 'bar',
        numerator: readArg(el, 'num'),
        denominator: readArg(el, 'den'),
      };
    case 'sSub':
      return { kind: 'subscript', base: readArg(el, 'e'), sub: readArg(el, 'sub') };
    case 'sSup':
      return { kind: 'superscript', base: readArg(el, 'e'), sup: readArg(el, 'sup') };
    case 'sSubSup':
      return {
        kind: 'subSuperscript',
        base: readArg(el, 'e'),
        sub: readArg(el, 'sub'),
        sup: readArg(el, 'sup'),
      };
    case 'rad':
      return {
        kind: 'radical',
        degree: flagOf(childOf(el, 'radPr'), 'degHide') ? null : readArg(el, 'deg'),
        body: readArg(el, 'e'),
      };
    case 'nary': {
      const props = childOf(el, 'naryPr');
      const limLoc = valOf(props, 'limLoc');
      return {
        kind: 'nary',
        operator: charOf(props, 'chr', DEFAULT_NARY_OPERATOR),
        ...(limLoc === 'undOvr'
          ? { limitLocation: 'underOver' as const }
          : limLoc === 'subSup'
            ? { limitLocation: 'subSup' as const }
            : {}),
        sub: flagOf(props, 'subHide') ? null : readArg(el, 'sub'),
        sup: flagOf(props, 'supHide') ? null : readArg(el, 'sup'),
        body: readArg(el, 'e'),
      };
    }
    case 'd': {
      const props = childOf(el, 'dPr');
      return {
        kind: 'delimiter',
        open: charOf(props, 'begChr', DEFAULT_DELIMITER_OPEN),
        close: charOf(props, 'endChr', DEFAULT_DELIMITER_CLOSE),
        separator: charOf(props, 'sepChr', DEFAULT_DELIMITER_SEPARATOR),
        items: childrenOf(el, 'e').map(readMathNodes),
      };
    }
    case 'm':
      return {
        kind: 'matrix',
        rows: childrenOf(el, 'mr').map((row) => childrenOf(row, 'e').map(readMathNodes)),
      };
    case 'acc':
      return {
        kind: 'accent',
        accent: charOf(childOf(el, 'accPr'), 'chr', DEFAULT_ACCENT),
        body: readArg(el, 'e'),
      };
    case 'func':
      return { kind: 'function', name: readArg(el, 'fName'), body: readArg(el, 'e') };
    case 'limLow':
    case 'limUpp':
      return {
        kind: el.name.localName === 'limLow' ? 'lowerLimit' : 'upperLimit',
        base: readArg(el, 'e'),
        limit: readArg(el, 'lim'),
      };
    // Property and control elements carry no content.
    case 'ctrlPr':
    case 'argPr':
    case 'oMathParaPr':
      return null;
    default:
      return { kind: 'unsupported', element: el.name.localName, text: mathPlainText(el) };
  }
};

/** Reads the content children of an OMML argument or `<m:oMath>`. */
export const readMathNodes = (parent: XmlElement): MathNodes => {
  const out: MathNode[] = [];
  for (const child of parent.children) {
    if (!isMath(child)) continue;
    const node = readNode(child);
    if (node !== null) out.push(node);
  }
  return out;
};

/** One `<m:oMath>` of an `<a14:m>`, with whether it sits in an `<m:oMathPara>`. */
export interface MathZoneEquation {
  readonly display: boolean;
  readonly oMath: XmlElement;
}

/**
 * The equations of one `<a14:m>`: its inline `<m:oMath>`, or each
 * `<m:oMath>` of its `<m:oMathPara>`.
 */
export const mathZoneEquations = (zone: XmlElement): MathZoneEquation[] => {
  const out: MathZoneEquation[] = [];
  for (const child of zone.children) {
    if (!isMath(child)) continue;
    if (child.name.localName === 'oMath') out.push({ display: false, oMath: child });
    else if (child.name.localName === 'oMathPara') {
      for (const oMath of childrenOf(child, 'oMath')) out.push({ display: true, oMath });
    }
  }
  return out;
};

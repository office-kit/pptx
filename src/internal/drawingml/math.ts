// Writes equations into DrawingML paragraphs.
//
// The desktop app reads an equation as `<a14:m>` ([MS-ODRAWXML] §2.3.1.21)
// holding OMML (ECMA-376 Part 1 §22.1). Apps without the a14 extension skip
// that element and show nothing, so the zone is written inside an
// `mc:AlternateContent` whose fallback is a run of the equation's linear text:
//
//   <mc:AlternateContent>
//     <mc:Choice Requires="a14"><a14:m><m:oMath>…</m:oMath></a14:m></mc:Choice>
//     <mc:Fallback><a:r><a:t>x=(−b±√(b^2−4ac))/(2a)</a:t></a:r></mc:Fallback>
//   </mc:AlternateContent>
//
// The desktop app itself wraps the whole shape and falls back to a picture,
// which only a renderer can draw. It opens this paragraph-level form as a
// native equation and rewrites it to its own form on save; LibreOffice shows
// the fallback text.

import { NS, attr, elem, qname, text, type XmlElement } from '../xml/index.ts';
import { applyRunFormat, type ImageEmbedder, type TextFormat } from './text-format.ts';

/**
 * Character style of a math run (`<m:sty>`). Omitted means the math default:
 * letters italic, digits and operators upright.
 */
export type MathTextStyle = 'plain' | 'bold' | 'italic' | 'boldItalic';

/** Where an n-ary operator puts its limits (`<m:limLoc>`). */
export type MathLimitLocation = 'underOver' | 'subSup';

/** Fraction layout (`<m:type>` of `<m:fPr>`). */
export type MathFractionType = 'bar' | 'skewed' | 'linear' | 'noBar';

/** A sequence of math nodes laid out left to right. */
export type MathNodeSpecs = ReadonlyArray<MathNodeSpec>;

/**
 * One node of an equation to write; the input counterpart of `MathNode`.
 * An omitted character takes its schema default, as in OMML: `∫` for an
 * n-ary operator, `(` `)` `|` for a delimiter, a circumflex for an accent.
 */
export type MathNodeSpec =
  | {
      readonly kind: 'text';
      readonly text: string;
      readonly style?: MathTextStyle;
      readonly format?: TextFormat;
    }
  | {
      readonly kind: 'fraction';
      readonly type?: MathFractionType;
      readonly numerator: MathNodeSpecs;
      readonly denominator: MathNodeSpecs;
    }
  | { readonly kind: 'subscript'; readonly base: MathNodeSpecs; readonly sub: MathNodeSpecs }
  | { readonly kind: 'superscript'; readonly base: MathNodeSpecs; readonly sup: MathNodeSpecs }
  | {
      readonly kind: 'subSuperscript';
      readonly base: MathNodeSpecs;
      readonly sub: MathNodeSpecs;
      readonly sup: MathNodeSpecs;
    }
  /** A `null` or omitted degree is a square root. */
  | {
      readonly kind: 'radical';
      readonly degree?: MathNodeSpecs | null;
      readonly body: MathNodeSpecs;
    }
  /** A `null` or omitted limit is hidden. */
  | {
      readonly kind: 'nary';
      readonly operator?: string;
      readonly limitLocation?: MathLimitLocation;
      readonly sub?: MathNodeSpecs | null;
      readonly sup?: MathNodeSpecs | null;
      readonly body: MathNodeSpecs;
    }
  /** An empty string omits that delimiter. */
  | {
      readonly kind: 'delimiter';
      readonly open?: string;
      readonly close?: string;
      readonly separator?: string;
      readonly items: ReadonlyArray<MathNodeSpecs>;
    }
  /** Every row has the same number of cells. */
  | { readonly kind: 'matrix'; readonly rows: ReadonlyArray<ReadonlyArray<MathNodeSpecs>> }
  /** `accent` is a combining character, such as U+0302 (hat) or U+0305 (bar). */
  | { readonly kind: 'accent'; readonly accent?: string; readonly body: MathNodeSpecs }
  | { readonly kind: 'function'; readonly name: MathNodeSpecs; readonly body: MathNodeSpecs }
  | {
      readonly kind: 'lowerLimit' | 'upperLimit';
      readonly base: MathNodeSpecs;
      readonly limit: MathNodeSpecs;
    };

/** An equation in a paragraph, beside its text runs. */
export interface EquationSpec {
  readonly equation: MathNodeSpecs;
  /** A display equation (`<m:oMathPara>`) instead of an inline one. */
  readonly display?: boolean;
  /** Format of every math run (size, color, …), under each text node's own. */
  readonly format?: TextFormat;
}

// The desktop app sets every math run in Cambria Math and writes the face
// into each run's `<a:rPr>`.
const MATH_FONT = 'Cambria Math';
const DEFAULT_ACCENT = '\u0302';
const DEFAULT_NARY_OPERATOR = '∫';

const m = (localName: string) => qname('m', localName, NS.math);
const NAME_VAL = m('val');
const NAME_A_RPR = qname('a', 'rPr', NS.dml);
const NAME_A_R = qname('a', 'r', NS.dml);
const NAME_A_T = qname('a', 't', NS.dml);
const NAME_A14_M = qname('a14', 'm', NS.a14);
const NAME_ALTERNATE_CONTENT = qname('mc', 'AlternateContent', NS.mc);
const NAME_CHOICE = qname('mc', 'Choice', NS.mc);
const NAME_FALLBACK = qname('mc', 'Fallback', NS.mc);
const ATTR_REQUIRES = qname('', 'Requires', '');
const ATTR_XML_SPACE = qname('xml', 'space', NS.xml);

const STY_BY_STYLE: Readonly<Record<MathTextStyle, string>> = {
  plain: 'p',
  bold: 'b',
  italic: 'i',
  boldItalic: 'bi',
};
const FRACTION_TOKEN: Readonly<Record<Exclude<MathFractionType, 'bar'>, string>> = {
  skewed: 'skw',
  linear: 'lin',
  noBar: 'noBar',
};

const valElement = (localName: string, value: string): XmlElement =>
  elem(m(localName), { attrs: [attr(NAME_VAL, value)] });

interface WriteContext {
  readonly format: TextFormat | undefined;
  readonly caller: string;
  readonly images: ImageEmbedder | undefined;
}

const textElement = (name: typeof NAME_A_T, value: string): XmlElement =>
  elem(name, {
    attrs: /^\s|\s$/.test(value) ? [attr(ATTR_XML_SPACE, 'preserve')] : [],
    children: [text(value)],
  });

const runElement = (node: Extract<MathNodeSpec, { kind: 'text' }>, ctx: WriteContext) => {
  const rPr = elem(NAME_A_RPR);
  applyRunFormat(rPr, { font: MATH_FONT, ...ctx.format, ...node.format }, ctx.caller, ctx.images);
  return elem(m('r'), {
    children: [
      ...(node.style === undefined
        ? []
        : [elem(m('rPr'), { children: [valElement('sty', STY_BY_STYLE[node.style])] })]),
      rPr,
      textElement(m('t'), node.text),
    ],
  });
};

const arg = (localName: string, nodes: MathNodeSpecs, ctx: WriteContext): XmlElement =>
  elem(m(localName), { children: nodes.map((node) => nodeElement(node, ctx)) });

const props = (localName: string, children: XmlElement[]): XmlElement[] =>
  children.length === 0 ? [] : [elem(m(localName), { children })];

const nodeElement = (node: MathNodeSpec, ctx: WriteContext): XmlElement => {
  switch (node.kind) {
    case 'text':
      return runElement(node, ctx);
    case 'fraction':
      return elem(m('f'), {
        children: [
          ...props(
            'fPr',
            node.type === undefined || node.type === 'bar'
              ? []
              : [valElement('type', FRACTION_TOKEN[node.type])],
          ),
          arg('num', node.numerator, ctx),
          arg('den', node.denominator, ctx),
        ],
      });
    case 'subscript':
      return elem(m('sSub'), { children: [arg('e', node.base, ctx), arg('sub', node.sub, ctx)] });
    case 'superscript':
      return elem(m('sSup'), { children: [arg('e', node.base, ctx), arg('sup', node.sup, ctx)] });
    case 'subSuperscript':
      return elem(m('sSubSup'), {
        children: [arg('e', node.base, ctx), arg('sub', node.sub, ctx), arg('sup', node.sup, ctx)],
      });
    case 'radical': {
      const degree = node.degree ?? [];
      // CT_Rad requires `<m:deg>` even when the degree is hidden.
      return elem(m('rad'), {
        children: [
          ...props('radPr', degree.length === 0 ? [valElement('degHide', '1')] : []),
          arg('deg', degree, ctx),
          arg('e', node.body, ctx),
        ],
      });
    }
    case 'nary': {
      const hiddenSub = node.sub === undefined || node.sub === null;
      const hiddenSup = node.sup === undefined || node.sup === null;
      // CT_NaryPr is a sequence: chr, limLoc, grow, subHide, supHide.
      return elem(m('nary'), {
        children: [
          ...props('naryPr', [
            ...(node.operator === undefined || node.operator === DEFAULT_NARY_OPERATOR
              ? []
              : [valElement('chr', node.operator)]),
            ...(node.limitLocation === undefined
              ? []
              : [valElement('limLoc', node.limitLocation === 'underOver' ? 'undOvr' : 'subSup')]),
            ...(hiddenSub ? [valElement('subHide', '1')] : []),
            ...(hiddenSup ? [valElement('supHide', '1')] : []),
          ]),
          arg('sub', node.sub ?? [], ctx),
          arg('sup', node.sup ?? [], ctx),
          arg('e', node.body, ctx),
        ],
      });
    }
    case 'delimiter':
      // CT_DPr is a sequence: begChr, sepChr, endChr.
      return elem(m('d'), {
        children: [
          ...props('dPr', [
            ...(node.open === undefined ? [] : [valElement('begChr', node.open)]),
            ...(node.separator === undefined ? [] : [valElement('sepChr', node.separator)]),
            ...(node.close === undefined ? [] : [valElement('endChr', node.close)]),
          ]),
          ...node.items.map((item) => arg('e', item, ctx)),
        ],
      });
    case 'matrix': {
      const columns = node.rows[0]?.length ?? 0;
      return elem(m('m'), {
        children: [
          // The column count the desktop app writes; it sizes the grid from it.
          elem(m('mPr'), {
            children: [
              elem(m('mcs'), {
                children: [
                  elem(m('mc'), {
                    children: [
                      elem(m('mcPr'), {
                        children: [
                          valElement('count', String(columns)),
                          valElement('mcJc', 'center'),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
          ...node.rows.map((row) =>
            elem(m('mr'), { children: row.map((cell) => arg('e', cell, ctx)) }),
          ),
        ],
      });
    }
    case 'accent':
      return elem(m('acc'), {
        children: [
          ...props(
            'accPr',
            node.accent === undefined || node.accent === DEFAULT_ACCENT
              ? []
              : [valElement('chr', node.accent)],
          ),
          arg('e', node.body, ctx),
        ],
      });
    case 'function':
      return elem(m('func'), {
        children: [arg('fName', node.name, ctx), arg('e', node.body, ctx)],
      });
    case 'lowerLimit':
    case 'upperLimit':
      return elem(m(node.kind === 'lowerLimit' ? 'limLow' : 'limUpp'), {
        children: [arg('e', node.base, ctx), arg('lim', node.limit, ctx)],
      });
  }
};

// ---------------------------------------------------------------------------
// Linear text for the fallback run, close to UnicodeMath so it reads as math.

// A single word (accents included) or symbol needs no parentheses as an operand.
const ATOM = /^(?:[\p{L}\p{M}\p{N}.]+|.)$/u;
const group = (s: string): string => (ATOM.test(s) ? s : `(${s})`);

const linear = (nodes: MathNodeSpecs): string => nodes.map(linearNode).join('');

const linearNode = (node: MathNodeSpec): string => {
  switch (node.kind) {
    case 'text':
      return node.text;
    case 'fraction':
      return `${group(linear(node.numerator))}/${group(linear(node.denominator))}`;
    case 'subscript':
      return `${group(linear(node.base))}_${group(linear(node.sub))}`;
    case 'superscript':
      return `${group(linear(node.base))}^${group(linear(node.sup))}`;
    case 'subSuperscript':
      return `${group(linear(node.base))}_${group(linear(node.sub))}^${group(linear(node.sup))}`;
    case 'radical': {
      const degree = node.degree ? linear(node.degree) : '';
      return degree === '' ? `√${group(linear(node.body))}` : `√(${degree}&${linear(node.body)})`;
    }
    case 'nary': {
      const sub = node.sub ? `_${group(linear(node.sub))}` : '';
      const sup = node.sup ? `^${group(linear(node.sup))}` : '';
      return `${node.operator ?? DEFAULT_NARY_OPERATOR}${sub}${sup} ${linear(node.body)}`;
    }
    case 'delimiter':
      return `${node.open ?? '('}${node.items.map(linear).join(node.separator ?? '|')}${node.close ?? ')'}`;
    case 'matrix':
      return `[${node.rows.map((row) => row.map(linear).join(', ')).join('; ')}]`;
    case 'accent':
      return `${group(linear(node.body))}${node.accent ?? DEFAULT_ACCENT}`;
    case 'function':
      return `${linear(node.name)} ${linear(node.body)}`;
    case 'lowerLimit':
      return `${linear(node.base)}_${group(linear(node.limit))}`;
    case 'upperLimit':
      return `${linear(node.base)}^${group(linear(node.limit))}`;
  }
};

const validate = (nodes: MathNodeSpecs, caller: string): void => {
  for (const node of nodes) {
    switch (node.kind) {
      case 'text':
        break;
      case 'fraction':
        validate(node.numerator, caller);
        validate(node.denominator, caller);
        break;
      case 'subscript':
        validate(node.base, caller);
        validate(node.sub, caller);
        break;
      case 'superscript':
        validate(node.base, caller);
        validate(node.sup, caller);
        break;
      case 'subSuperscript':
        validate(node.base, caller);
        validate(node.sub, caller);
        validate(node.sup, caller);
        break;
      case 'radical':
        validate(node.degree ?? [], caller);
        validate(node.body, caller);
        break;
      case 'nary':
        validate(node.sub ?? [], caller);
        validate(node.sup ?? [], caller);
        validate(node.body, caller);
        break;
      case 'delimiter':
        if (node.items.length === 0) {
          throw new RangeError(`${caller}: a delimiter needs at least one item`);
        }
        for (const item of node.items) validate(item, caller);
        break;
      case 'matrix': {
        const columns = node.rows[0]?.length ?? 0;
        if (columns === 0) throw new RangeError(`${caller}: a matrix needs at least one cell`);
        for (const row of node.rows) {
          if (row.length !== columns) {
            throw new RangeError(`${caller}: every matrix row needs ${columns} cells`);
          }
          for (const cell of row) validate(cell, caller);
        }
        break;
      }
      case 'accent':
        validate(node.body, caller);
        break;
      case 'function':
        validate(node.name, caller);
        validate(node.body, caller);
        break;
      case 'lowerLimit':
      case 'upperLimit':
        validate(node.base, caller);
        validate(node.limit, caller);
        break;
      default:
        throw new TypeError(
          `${caller}: unknown math node kind ${JSON.stringify((node as { kind?: unknown }).kind)}`,
        );
    }
  }
};

/** Builds the paragraph child that holds `spec`: an equation with its text fallback. */
export const buildEquation = (
  spec: EquationSpec,
  caller: string,
  images?: ImageEmbedder,
): XmlElement => {
  validate(spec.equation, caller);
  const ctx: WriteContext = { format: spec.format, caller, images };
  const oMath = elem(m('oMath'), { children: spec.equation.map((node) => nodeElement(node, ctx)) });
  const zone = elem(NAME_A14_M, {
    children: [
      spec.display === true
        ? elem(m('oMathPara'), { prefixDecls: new Map([['m', NS.math]]), children: [oMath] })
        : { ...oMath, prefixDecls: new Map([['m', NS.math]]) },
    ],
  });
  const fallbackPr = elem(NAME_A_RPR);
  if (spec.format !== undefined) applyRunFormat(fallbackPr, spec.format, caller, images);
  return elem(NAME_ALTERNATE_CONTENT, {
    prefixDecls: new Map([['mc', NS.mc]]),
    children: [
      elem(NAME_CHOICE, {
        attrs: [attr(ATTR_REQUIRES, 'a14')],
        prefixDecls: new Map([['a14', NS.a14]]),
        children: [zone],
      }),
      elem(NAME_FALLBACK, {
        children: [
          elem(NAME_A_R, { children: [fallbackPr, textElement(NAME_A_T, linear(spec.equation))] }),
        ],
      }),
    ],
  });
};

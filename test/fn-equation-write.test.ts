// `setShapeParagraphs` / `setTableCellParagraphs` write an `EquationSpec` as an
// `<a14:m>` equation inside an `mc:AlternateContent` whose fallback is the
// equation's linear text, so apps without the a14 extension still show it.

import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTable,
  addSlideTextBox,
  createPresentation,
  getShapeParagraphElements,
  getShapeText,
  getShapeXmlString,
  getSlides,
  getSlideShapes,
  getTableCellParagraphs,
  getTableCells,
  inches,
  loadPresentation,
  savePresentation,
  setShapeParagraphs,
  setShapeTextFormat,
  setTableCellParagraphs,
  validatePresentation,
  type MathNodeSpec,
  type MathNodeSpecs,
} from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const t = (text: string): MathNodeSpec => ({ kind: 'text', text });

// x = (−b ± √(b² − 4ac)) / 2a
const QUADRATIC: MathNodeSpecs = [
  t('x='),
  {
    kind: 'fraction',
    numerator: [
      t('−b±'),
      {
        kind: 'radical',
        body: [{ kind: 'superscript', base: [t('b')], sup: [t('2')] }, t('−4ac')],
      },
    ],
    denominator: [t('2a')],
  },
];

const textBox = () => {
  const pres = createPresentation();
  const shape = addSlideTextBox(addBlankSlide(pres), {
    x: inches(1),
    y: inches(1),
    w: inches(6),
    h: inches(2),
    text: '',
  });
  return { pres, shape };
};

const reload = async (pres: ReturnType<typeof createPresentation>) => {
  const loaded = await loadPresentation(await savePresentation(pres));
  return { loaded, shape: getSlideShapes(getSlides(loaded)[0]!)[0]! };
};

describe('writing equations', () => {
  it('writes an inline equation between runs and reads it back', async () => {
    const { pres, shape } = textBox();
    setShapeParagraphs(shape, [
      { runs: [{ text: 'Roots: ' }, { equation: QUADRATIC }, { text: ' (real)' }] },
    ]);
    const { loaded, shape: read } = await reload(pres);
    expect(validatePresentation(loaded)).toEqual([]);
    const elements = getShapeParagraphElements(read, 0);
    expect(elements.map((e) => e.kind)).toEqual(['r', 'math', 'r']);
    const math = elements[1]!;
    if (math.kind !== 'math') throw new Error('expected an equation');
    expect(math.display).toBe(false);
    expect(math.text).toBe('x=−b±b2−4ac2a');
    const font = { font: 'Cambria Math' };
    const r = (text: string) => ({ kind: 'text', text, format: font });
    expect(math.math).toEqual([
      r('x='),
      {
        kind: 'fraction',
        type: 'bar',
        numerator: [
          r('−b±'),
          {
            kind: 'radical',
            degree: null,
            body: [{ kind: 'superscript', base: [r('b')], sup: [r('2')] }, r('−4ac')],
          },
        ],
        denominator: [r('2a')],
      },
    ]);
    // Plain text keeps ignoring equations.
    expect(getShapeText(read)).toBe('Roots:  (real)');
  });

  it('wraps the zone in AlternateContent with a linear-text fallback', () => {
    const { shape } = textBox();
    setShapeParagraphs(shape, [{ runs: [{ equation: QUADRATIC, format: { size: 20 } }] }]);
    const xml = getShapeXmlString(shape);
    expect(xml).toMatch(
      /<mc:AlternateContent xmlns:mc="[^"]+"><mc:Choice xmlns:a14="[^"]+" Requires="a14"><a14:m><m:oMath xmlns:m="http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/math"><m:r><a:rPr sz="2000"><a:latin typeface="Cambria Math"\/><\/a:rPr><m:t>x=<\/m:t><\/m:r>/,
    );
    expect(xml).toContain(
      '<mc:Fallback><a:r><a:rPr sz="2000"/><a:t>x=(−b±√(b^2−4ac))/2a</a:t></a:r></mc:Fallback></mc:AlternateContent>',
    );
  });

  it.each<[string, MathNodeSpecs, string]>([
    [
      'a sum with limits',
      [
        {
          kind: 'nary',
          operator: '∑',
          sub: [t('i=1')],
          sup: [t('n')],
          body: [{ kind: 'subscript', base: [{ kind: 'accent', body: [t('x')] }], sub: [t('i')] }],
        },
      ],
      '∑_(i=1)^n x\u0302_i',
    ],
    [
      'a matrix in parentheses',
      [
        {
          kind: 'delimiter',
          items: [
            [
              {
                kind: 'matrix',
                rows: [
                  [[t('a')], [t('b')]],
                  [[t('c')], [t('d')]],
                ],
              },
            ],
          ],
        },
      ],
      '([a, b; c, d])',
    ],
    ['a cube root', [{ kind: 'radical', degree: [t('3')], body: [t('x')] }], '√(3&amp;x)'],
    ['a hat accent', [{ kind: 'accent', body: [t('y')] }], 'y\u0302'],
    [
      'a limit',
      [
        {
          kind: 'function',
          name: [{ kind: 'lowerLimit', base: [t('lim')], limit: [t('n→∞')] }],
          body: [t('a')],
        },
      ],
      'lim_(n→∞) a',
    ],
  ])('falls back to linear text for %s', (_, equation, fallback) => {
    const { shape } = textBox();
    setShapeParagraphs(shape, [{ runs: [{ equation }] }]);
    expect(getShapeXmlString(shape)).toContain(`<mc:Fallback><a:r><a:rPr/><a:t>${fallback}</a:t>`);
  });

  it('writes the defaults the schema leaves implicit as omissions, and reads them back resolved', async () => {
    const { pres, shape } = textBox();
    const equation: MathNodeSpecs = [
      { kind: 'nary', body: [t('f')] },
      { kind: 'delimiter', open: '[', close: '', items: [[t('z')]] },
      { kind: 'accent', accent: '\u0305', body: [t('x')] },
      { kind: 'fraction', type: 'linear', numerator: [t('1')], denominator: [t('2')] },
      { kind: 'text', text: 'd', style: 'plain' },
      {
        kind: 'nary',
        operator: '∑',
        limitLocation: 'subSup',
        sub: [t('i')],
        body: [t('g')],
      },
    ];
    setShapeParagraphs(shape, [{ runs: [{ equation, display: true }] }]);
    const xml = getShapeXmlString(shape);
    expect(xml).toContain(
      '<m:nary><m:naryPr><m:subHide m:val="1"/><m:supHide m:val="1"/></m:naryPr>',
    );
    expect(xml).toContain('<m:dPr><m:begChr m:val="["/><m:endChr m:val=""/></m:dPr>');
    expect(xml).toContain('<m:oMathPara xmlns:m=');
    const { shape: read } = await reload(pres);
    const [math] = getShapeParagraphElements(read, 0);
    if (math?.kind !== 'math') throw new Error('expected an equation');
    expect(math.display).toBe(true);
    const r = (text: string, style?: 'plain') => ({
      kind: 'text',
      text,
      ...(style ? { style } : {}),
      format: { font: 'Cambria Math' },
    });
    expect(math.math).toEqual([
      { kind: 'nary', operator: '∫', sub: null, sup: null, body: [r('f')] },
      { kind: 'delimiter', open: '[', close: '', separator: '|', items: [[r('z')]] },
      { kind: 'accent', accent: '\u0305', body: [r('x')] },
      { kind: 'fraction', type: 'linear', numerator: [r('1')], denominator: [r('2')] },
      r('d', 'plain'),
      {
        kind: 'nary',
        operator: '∑',
        limitLocation: 'subSup',
        sub: [r('i')],
        sup: null,
        body: [r('g')],
      },
    ]);
  });

  it('lets a text node format override the equation format', async () => {
    const { pres, shape } = textBox();
    setShapeParagraphs(shape, [
      {
        runs: [
          {
            equation: [t('a'), { kind: 'text', text: 'b', format: { color: '#FF0000' } }],
            format: { size: 24 },
          },
        ],
      },
    ]);
    const { shape: read } = await reload(pres);
    const [math] = getShapeParagraphElements(read, 0);
    if (math?.kind !== 'math') throw new Error('expected an equation');
    expect(math.math[0]).toMatchObject({ format: { size: 24, font: 'Cambria Math' } });
    expect(math.math[1]).toMatchObject({ format: { size: 24, color: '#FF0000' } });
  });

  it('gives equations the shape-wide format, keeping their math face', async () => {
    const { pres, shape } = textBox();
    setShapeParagraphs(shape, [{ runs: [{ text: 'f: ' }, { equation: [t('x')] }] }]);
    setShapeTextFormat(shape, { size: 30, color: '#1F4E79', bold: true, font: 'Arial' });
    expect(getShapeXmlString(shape)).toContain(
      '<mc:Fallback><a:r><a:rPr sz="3000" b="1"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill><a:latin typeface="Arial"/></a:rPr>',
    );
    const { shape: read } = await reload(pres);
    const [, math] = getShapeParagraphElements(read, 0);
    if (math?.kind !== 'math') throw new Error('expected an equation');
    expect(math.math[0]).toMatchObject({
      format: { size: 30, color: '#1F4E79', bold: true, font: 'Cambria Math' },
    });
  });

  it('writes equations into table cells', async () => {
    const pres = createPresentation();
    const table = addSlideTable(addBlankSlide(pres), {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      rows: [['']],
    });
    setTableCellParagraphs(getTableCells(table)[0]![0]!, [{ runs: [{ equation: QUADRATIC }] }]);
    const loaded = await loadPresentation(await savePresentation(pres));
    const cell = getTableCells(getSlideShapes(getSlides(loaded)[0]!)[0]!)[0]![0]!;
    expect(getTableCellParagraphs(cell)[0]!.elements.map((e) => e.kind)).toEqual(['math']);
  });

  it.each<[string, MathNodeSpecs, RegExp]>([
    ['an empty matrix', [{ kind: 'matrix', rows: [] }], /a matrix needs at least one cell/],
    [
      'ragged matrix rows',
      [{ kind: 'matrix', rows: [[[t('a')], [t('b')]], [[t('c')]]] }],
      /every matrix row needs 2 cells/,
    ],
    [
      'an empty delimiter',
      [{ kind: 'delimiter', items: [] }],
      /a delimiter needs at least one item/,
    ],
    [
      'an unknown node',
      [{ kind: 'box' } as unknown as MathNodeSpec],
      /unknown math node kind "box"/,
    ],
  ])('rejects %s and leaves the text unchanged', (_, equation, message) => {
    const { shape } = textBox();
    setShapeParagraphs(shape, [{ runs: [{ text: 'before' }] }]);
    expect(() => setShapeParagraphs(shape, [{ runs: [{ equation }] }])).toThrow(message);
    expect(getShapeText(shape)).toBe('before');
  });

  it.skipIf(!isSchemaValidationAvailable())('writes schema-valid XML', () => {
    const { shape } = textBox();
    const equation: MathNodeSpecs = [
      ...QUADRATIC,
      {
        kind: 'nary',
        operator: '∑',
        limitLocation: 'underOver',
        sub: [t('i')],
        sup: [t('n')],
        body: [t('x')],
      },
      { kind: 'delimiter', open: '{', close: '', separator: ',', items: [[t('a')], [t('b')]] },
      { kind: 'matrix', rows: [[[t('a')], [t('b')]]] },
      { kind: 'accent', accent: '\u20D7', body: [t('v')] },
      { kind: 'subSuperscript', base: [t('x')], sub: [t('0')], sup: [t('2')] },
      { kind: 'function', name: [{ kind: 'text', text: 'sin', style: 'plain' }], body: [t('θ')] },
      { kind: 'upperLimit', base: [t('x')], limit: [t('k')] },
      { kind: 'fraction', type: 'skewed', numerator: [t('1')], denominator: [t('2')] },
    ];
    setShapeParagraphs(shape, [{ runs: [{ equation }, { equation, display: true }] }]);
    const slide = getShapeXmlString(shape).replace(/^<\?xml[^>]*\?>/, '');
    // What an app without a14 sees: the fallback runs in place of each zone.
    const fallback = slide.replace(
      /<mc:AlternateContent\b[^>]*>[\s\S]*?<mc:Fallback>([\s\S]*?)<\/mc:Fallback><\/mc:AlternateContent>/g,
      '$1',
    );
    expectSchemaValid(
      `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${fallback}</p:spTree></p:cSld></p:sld>`,
      'pml',
    );
    // The OMML itself, without the [MS-ODRAWXML] `<a:rPr>` that each run adds
    // to ECMA-376's CT_R.
    const zones = [...slide.matchAll(/<a14:m>([\s\S]*?)<\/a14:m>/g)].map((m) =>
      m[1]!.replace(/<a:rPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:rPr>)/g, ''),
    );
    expect(zones).toHaveLength(2);
    for (const zone of zones) expectSchemaValid(zone, 'math');
  });
});

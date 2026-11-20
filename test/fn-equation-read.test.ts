// Equations (`<a14:m>`) inside a paragraph reads as `math` elements of
// `getShapeParagraphElements`, modeled on the OMML elements of ECMA-376 §22.1.

import { describe, expect, it } from 'vitest';
import {
  getParagraphElementLanguages,
  getShapeParagraphElements,
  getSlides,
  getSlideShapes,
  getTableCellParagraphs,
  getTableCells,
  loadPresentation,
  savePresentation,
  type MathNode,
  type PresentationData,
} from '../src/api/index.ts';
import {
  displayMathZone,
  mathRun as run,
  mathZone as zone,
  presentationWithParagraph,
} from './lib/equation-deck.ts';

const firstShape = (pres: PresentationData) => getSlideShapes(getSlides(pres)[0]!)[0]!;
const deckWith = async (xml: string, table = false) =>
  firstShape(await presentationWithParagraph(xml, { table }));

// What the desktop app writes for σ(z) = 1/(1+e^{−z}): math italic code
// points and a DrawingML `<a:rPr>` in every run (trimmed to two runs' worth).
const DESKTOP_RUN_PR =
  '<a:rPr><a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr>';
const desktopRun = (text: string) => `<m:r>${DESKTOP_RUN_PR}<m:t>${text}</m:t></m:r>`;

describe('getShapeParagraphElements: equations', () => {
  it('reads an inline equation between runs, in document order', async () => {
    const shape = await deckWith(
      `<a:r><a:rPr lang="en-US"/><a:t>Inline: </a:t></a:r>` +
        zone(
          desktopRun('𝜎') +
            `<m:f><m:fPr><m:ctrlPr>${DESKTOP_RUN_PR}</m:ctrlPr></m:fPr><m:num>${desktopRun('1')}</m:num>` +
            `<m:den>${desktopRun('1+')}<m:sSup><m:e>${desktopRun('𝑒')}</m:e><m:sup>${desktopRun('−𝑧')}</m:sup></m:sSup></m:den></m:f>`,
        ) +
        `<a:r><a:rPr lang="ja-JP"/><a:t> after</a:t></a:r>`,
    );
    const elements = getShapeParagraphElements(shape, 0);
    expect(elements.map((e) => e.kind)).toEqual(['r', 'math', 'r']);
    const math = elements[1]!;
    expect(math).toMatchObject({ kind: 'math', display: false, text: '𝜎11+𝑒−𝑧', format: null });
    if (math.kind !== 'math') throw new Error('unreachable');
    expect(math.math).toEqual([
      { kind: 'text', text: '𝜎', format: { font: 'Cambria Math' } },
      {
        kind: 'fraction',
        type: 'bar',
        numerator: [{ kind: 'text', text: '1', format: { font: 'Cambria Math' } }],
        denominator: [
          { kind: 'text', text: '1+', format: { font: 'Cambria Math' } },
          {
            kind: 'superscript',
            base: [{ kind: 'text', text: '𝑒', format: { font: 'Cambria Math' } }],
            sup: [{ kind: 'text', text: '−𝑧', format: { font: 'Cambria Math' } }],
          },
        ],
      },
    ]);
    // Languages stay aligned with the element list: the equation has none.
    expect(getParagraphElementLanguages(shape, 0)).toEqual(['en-US', null, 'ja-JP']);
  });

  it('reads each equation of a display zone as its own element', async () => {
    const shape = await deckWith(displayMathZone(run('a'), run('b')));
    expect(getShapeParagraphElements(shape, 0)).toEqual([
      {
        kind: 'math',
        text: 'a',
        display: true,
        format: null,
        math: [{ kind: 'text', text: 'a', format: null }],
      },
      {
        kind: 'math',
        text: 'b',
        display: true,
        format: null,
        math: [{ kind: 'text', text: 'b', format: null }],
      },
    ]);
  });

  it('resolves the characters and flags the schema leaves implicit', async () => {
    const shape = await deckWith(
      zone(
        `<m:d><m:e>${run('x')}</m:e><m:e>${run('y')}</m:e></m:d>` +
          `<m:d><m:dPr><m:begChr m:val="["/><m:endChr/></m:dPr><m:e>${run('z')}</m:e></m:d>` +
          `<m:nary><m:naryPr><m:subHide/></m:naryPr><m:sub/><m:sup>${run('1')}</m:sup><m:e>${run('f')}</m:e></m:nary>` +
          `<m:nary><m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="undOvr"/><m:supHide m:val="on"/></m:naryPr><m:sub>${run('i')}</m:sub><m:sup/><m:e>${run('g')}</m:e></m:nary>` +
          `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${run('2')}</m:e></m:rad>` +
          `<m:rad><m:radPr><m:degHide m:val="off"/></m:radPr><m:deg>${run('3')}</m:deg><m:e>${run('8')}</m:e></m:rad>` +
          `<m:acc><m:e>${run('y')}</m:e></m:acc>` +
          `<m:f><m:fPr><m:type m:val="lin"/></m:fPr><m:num>${run('1')}</m:num><m:den>${run('2')}</m:den></m:f>` +
          `<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>d</m:t></m:r>`,
      ),
    );
    const [math] = getShapeParagraphElements(shape, 0);
    if (math?.kind !== 'math') throw new Error('expected an equation');
    const t = (text: string): MathNode => ({ kind: 'text', text, format: null });
    expect(math.math).toEqual([
      { kind: 'delimiter', open: '(', close: ')', separator: '|', items: [[t('x')], [t('y')]] },
      { kind: 'delimiter', open: '[', close: '', separator: '|', items: [[t('z')]] },
      { kind: 'nary', operator: '∫', sub: null, sup: [t('1')], body: [t('f')] },
      {
        kind: 'nary',
        operator: '∑',
        limitLocation: 'underOver',
        sub: [t('i')],
        sup: null,
        body: [t('g')],
      },
      { kind: 'radical', degree: null, body: [t('2')] },
      { kind: 'radical', degree: [t('3')], body: [t('8')] },
      { kind: 'accent', accent: '\u0302', body: [t('y')] },
      { kind: 'fraction', type: 'linear', numerator: [t('1')], denominator: [t('2')] },
      { kind: 'text', text: 'd', style: 'plain', format: null },
    ]);
  });

  it('reads matrices, scripts, functions and limits', async () => {
    const shape = await deckWith(
      zone(
        `<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="2"/></m:mcPr></m:mc></m:mcs></m:mPr>` +
          `<m:mr><m:e>${run('a')}</m:e><m:e>${run('b')}</m:e></m:mr><m:mr><m:e>${run('c')}</m:e><m:e>${run('d')}</m:e></m:mr></m:m>` +
          `<m:sSub><m:e>${run('y')}</m:e><m:sub>${run('i')}</m:sub></m:sSub>` +
          `<m:sSubSup><m:e>${run('x')}</m:e><m:sub>${run('0')}</m:sub><m:sup>${run('2')}</m:sup></m:sSubSup>` +
          `<m:func><m:fName><m:limLow><m:e>${run('lim')}</m:e><m:lim>${run('n→∞')}</m:lim></m:limLow></m:fName><m:e>${run('a')}</m:e></m:func>` +
          `<m:limUpp><m:e>${run('x')}</m:e><m:lim>${run('k')}</m:lim></m:limUpp>`,
      ),
    );
    const [math] = getShapeParagraphElements(shape, 0);
    if (math?.kind !== 'math') throw new Error('expected an equation');
    const t = (text: string): MathNode => ({ kind: 'text', text, format: null });
    expect(math.math).toEqual([
      {
        kind: 'matrix',
        rows: [
          [[t('a')], [t('b')]],
          [[t('c')], [t('d')]],
        ],
      },
      { kind: 'subscript', base: [t('y')], sub: [t('i')] },
      { kind: 'subSuperscript', base: [t('x')], sub: [t('0')], sup: [t('2')] },
      {
        kind: 'function',
        name: [{ kind: 'lowerLimit', base: [t('lim')], limit: [t('n→∞')] }],
        body: [t('a')],
      },
      { kind: 'upperLimit', base: [t('x')], limit: [t('k')] },
    ]);
  });

  it('keeps unmodeled elements as unsupported nodes with their text', async () => {
    const shape = await deckWith(
      zone(`<m:eqArr><m:e>${run('x=1')}</m:e><m:e>${run('y=2')}</m:e></m:eqArr>`),
    );
    const [math] = getShapeParagraphElements(shape, 0);
    if (math?.kind !== 'math') throw new Error('expected an equation');
    expect(math.math).toEqual([{ kind: 'unsupported', element: 'eqArr', text: 'x=1y=2' }]);
  });

  it('keeps table-cell click actions on the runs they belong to', async () => {
    const shape = await deckWith(
      zone(run('x')) +
        `<a:r><a:rPr lang="en-US"><a:hlinkClick r:id="" action="ppaction://hlinkshowjump?jump=nextslide"/></a:rPr><a:t>next</a:t></a:r>`,
      true,
    );
    const cell = getTableCells(shape)[0]![0]!;
    const [paragraph] = getTableCellParagraphs(cell);
    expect(paragraph!.elements.map((e) => e.kind)).toEqual(['math', 'r']);
    expect(paragraph!.elements[0]).not.toHaveProperty('clickAction');
    expect(paragraph!.elements[1]).toHaveProperty('clickAction');
  });

  it('survives a load / save round trip unchanged', async () => {
    const xml = zone(`<m:f><m:num>${run('1')}</m:num><m:den>${run('2')}</m:den></m:f>`);
    const pres = await presentationWithParagraph(xml);
    const before = getShapeParagraphElements(firstShape(pres), 0);
    expect(before[0]?.kind).toBe('math');
    const reloaded = await loadPresentation(await savePresentation(pres));
    const after = getShapeParagraphElements(firstShape(reloaded), 0);
    expect(after).toEqual(before);
  });
});

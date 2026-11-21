// The preview lays out equations (`<a14:m>`) as an inline SVG box in the
// line, instead of leaving the equation out.

import { describe, expect, it } from 'vitest';
import { getSlides } from '../src/api/index.ts';
import { auditTextLayout, renderSlideToSvg } from '../packages/preview/src/index.ts';
import {
  displayMathZone,
  mathRun as run,
  mathZone as zone,
  presentationWithParagraph,
} from './lib/equation-deck.ts';

const FRACTION = `<m:f><m:num>${run('1')}</m:num><m:den>${run('2')}</m:den></m:f>`;
const BEFORE = '<a:r><a:rPr lang="en-US"/><a:t>Before </a:t></a:r>';

const mathGroups = (svg: string): string[] => svg.match(/<g data-pptx-math="">.*?<\/g>/g) ?? [];

const textAt = (group: string): Array<{ x: number; y: number; size: number; text: string }> =>
  [
    ...group.matchAll(
      /<text x="([^"]+)" y="([^"]+)"[^>]*font-size="([^"]+)"[^>]*>([^<]*)<\/text>/g,
    ),
  ].map((m) => ({ x: Number(m[1]), y: Number(m[2]), size: Number(m[3]), text: m[4]! }));

describe('equations in the preview', () => {
  it.each(['svg', 'foreignObject'] as const)(
    'draws an inline equation (%s)',
    async (textLayout) => {
      const pres = await presentationWithParagraph(BEFORE + zone(run('𝑥=') + FRACTION));
      const svg = renderSlideToSvg(pres, getSlides(pres)[0]!, { textLayout });
      const [group] = mathGroups(svg);
      expect(group).toBeDefined();
      const texts = textAt(group!);
      // Math italic code points fold back to the letter, set in italic.
      expect(group).toMatch(/font-style="italic"[^>]*>x<\/text>/);
      // The fraction's numerator sits above its denominator, both at script size.
      const one = texts.find((t) => t.text === '1')!;
      const two = texts.find((t) => t.text === '2')!;
      const x = texts.find((t) => t.text === 'x')!;
      expect(one.y).toBeLessThan(x.y);
      expect(two.y).toBeGreaterThan(x.y);
      expect(one.size).toBeCloseTo(x.size * 0.7, 1);
      expect(group).toContain('<line ');
      // The equation follows the run before it on the same baseline.
      expect(svg).toMatch(/Before <\/tspan>/);
    },
  );

  it('sets a display equation at full size with limits above and below', async () => {
    const sum = `<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub>${run('i')}</m:sub><m:sup>${run('n')}</m:sup><m:e>${run('a')}</m:e></m:nary>`;
    const inline = await presentationWithParagraph(zone(FRACTION + sum));
    const display = await presentationWithParagraph(displayMathZone(FRACTION + sum));
    const textsOf = (pres: typeof inline) =>
      textAt(mathGroups(renderSlideToSvg(pres, getSlides(pres)[0]!, { textLayout: 'svg' }))[0]!);
    const inlineTexts = textsOf(inline);
    const displayTexts = textsOf(display);
    const find = (texts: typeof inlineTexts, text: string) => texts.find((t) => t.text === text)!;
    expect(find(displayTexts, '1').size).toBeGreaterThan(find(inlineTexts, '1').size);
    // Inline limits sit to the right of the operator; display limits stack on it.
    const op = (texts: typeof inlineTexts) => find(texts, '∑');
    expect(find(inlineTexts, 'n').x).toBeGreaterThan(op(inlineTexts).x);
    expect(Math.abs(find(displayTexts, 'n').x - find(displayTexts, 'i').x)).toBeLessThan(
      op(displayTexts).size,
    );
    expect(find(displayTexts, 'n').y).toBeLessThan(
      op(displayTexts).y - op(displayTexts).size * 0.5,
    );
  });

  it('grows the line to fit a tall equation', async () => {
    const tall = `<m:f><m:num>${FRACTION}</m:num><m:den>${FRACTION}</m:den></m:f>`;
    const plain = await presentationWithParagraph(BEFORE + '<a:br/>' + BEFORE);
    const withMath = await presentationWithParagraph(
      BEFORE + displayMathZone(tall) + '<a:br/>' + BEFORE,
    );
    const secondBaseline = (pres: typeof plain) => {
      const svg = renderSlideToSvg(pres, getSlides(pres)[0]!, { textLayout: 'svg' });
      return [...svg.matchAll(/<text x="[^"]+" y="([^"]+)" text-anchor/g)].map((m) =>
        Number(m[1]),
      )[1]!;
    };
    expect(secondBaseline(withMath)).toBeGreaterThan(secondBaseline(plain) + 10);
  });

  it('draws equations in table cells', async () => {
    const pres = await presentationWithParagraph(zone(FRACTION), { table: true });
    const svg = renderSlideToSvg(pres, getSlides(pres)[0]!);
    expect(mathGroups(svg)).toHaveLength(1);
  });

  it('marks and audits elements outside the modeled subset', async () => {
    const pres = await presentationWithParagraph(
      zone(
        `<m:eqArr><m:e>${run('x=1')}</m:e></m:eqArr><m:groupChr><m:e>${run('y')}</m:e></m:groupChr>`,
      ),
    );
    const svg = renderSlideToSvg(pres, getSlides(pres)[0]!, { textLayout: 'svg' });
    expect(svg).toContain('<title>Unsupported equation element: eqArr</title>');
    expect(svg).toMatch(/>x=1<\/text>/);
    expect(auditTextLayout(pres)).toContainEqual(
      expect.objectContaining({
        kind: 'unsupported-math',
        paragraphIndex: 0,
        elements: ['eqArr', 'groupChr'],
      }),
    );
  });
});

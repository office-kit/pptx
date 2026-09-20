import { describe, expect, it } from 'vitest';
import * as api from '../src/api/index.ts';
import { Presentation, Slide, Text, compile, type TextProps } from '../packages/dsl/src/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const BOX = { x: 0.6, y: 1, width: 8, height: 2 } as const;

const compileText = async (props: Omit<TextProps, keyof typeof BOX>) => {
  const pres = await compile(
    Presentation({ children: Slide({ children: Text({ ...BOX, ...props }) }) }),
  );
  return { pres, shape: api.getSlideShapes(api.getSlides(pres)[0]!)[0]! };
};

const leftMargins = (shape: api.SlideShapeData) =>
  [...api.getShapeXmlString(shape).matchAll(/<a:pPr\b([^>]*)>/g)].map(
    (match) => / marL="(\d+)"/.exec(match[1]!)?.[1] ?? null,
  );

describe('Text paragraph bullet and level', () => {
  it('writes a bullet and a level on single paragraphs', async () => {
    const { shape } = await compileText({
      paragraphs: [
        { runs: [{ text: 'Churn fell to 4.2%' }], bullet: 'bullet' },
        { runs: [{ text: 'Driver: simpler setup' }], bullet: { char: '–' }, level: 1 },
        { runs: [{ text: 'Plain closing line' }] },
      ],
    });
    expect([0, 1, 2].map((i) => api.getParagraphBullet(shape, i))).toEqual([
      'bullet',
      { char: '–' },
      null,
    ]);
    expect([0, 1, 2].map((i) => api.getParagraphLevel(shape, i))).toEqual([0, 1, 0]);
  });

  it("lets a paragraph's bullet win over bullets", async () => {
    const { shape } = await compileText({
      bullets: 'number',
      paragraphs: [
        { runs: [{ text: 'Heading' }], bullet: 'none' },
        { runs: [{ text: 'First' }] },
        { runs: [{ text: 'Detail' }], bullet: { char: '–' }, level: 1 },
      ],
    });
    expect([0, 1, 2].map((i) => api.getParagraphBullet(shape, i))).toEqual([
      'none',
      'number',
      { char: '–' },
    ]);
  });

  // The core derives a bullet's hanging indent from the level at the moment the
  // bullet is written, and an explicit marL overrides the level's own indent.
  it('indents a nested bullet deeper than its parent', async () => {
    const { shape } = await compileText({
      bullets: 'bullet',
      paragraphs: [
        { runs: [{ text: 'Parent' }] },
        { runs: [{ text: 'Child' }], level: 1 },
        { runs: [{ text: 'Grandchild' }], level: 2, bullet: { char: '–' } },
      ],
    });
    const [parent, child, grandchild] = leftMargins(shape).map(Number);
    expect(parent).toBeGreaterThan(0);
    expect(child).toBeGreaterThan(parent!);
    expect(grandchild).toBeGreaterThan(child!);
  });

  it('keeps paragraphSpacing, paragraph align and run formats next to them', async () => {
    const { shape } = await compileText({
      size: 16,
      bullets: 'bullet',
      paragraphSpacing: { after: 6 },
      paragraphs: [
        { align: 'right', level: 1, runs: [{ text: 'a' }, { text: 'b', format: { bold: true } }] },
      ],
    });
    expect(api.getParagraphAlignment(shape, 0)).toBe('r');
    expect(api.getParagraphLevel(shape, 0)).toBe(1);
    expect(api.getParagraphSpacing(shape, 0)).toMatchObject({ afterPts: 6 });
    expect(
      api.getShapeParagraphElements(shape, 0).map((run) => run.kind === 'r' && run.format),
    ).toEqual([{ size: 16 }, { size: 16, bold: true }]);
  });

  it('rejects an out-of-range level in the type, and at run time for untyped callers', async () => {
    await expect(
      // @ts-expect-error TextLevel is 0 to 8.
      compileText({ paragraphs: [{ runs: [{ text: 'x' }], level: 9 }] }),
    ).rejects.toThrow(/level/);
  });

  (isSchemaValidationAvailable() ? it : it.skip)('emits schema-valid slide XML', async () => {
    const { pres } = await compileText({
      bullets: 'number',
      paragraphSpacing: { before: 4, after: 4 },
      paragraphs: [
        { runs: [{ text: 'a' }] },
        { runs: [{ text: 'b' }], level: 1, bullet: { char: '–' } },
        { runs: [{ text: 'c' }], bullet: 'none' },
      ],
    });
    const reloaded = await api.loadPresentation(await api.savePresentation(pres));
    expect(api.validatePresentation(reloaded).filter((i) => i.severity === 'error')).toEqual([]);
    expectSchemaValid(
      new TextDecoder().decode(api.readPackagePart(reloaded, '/ppt/slides/slide1.xml')!),
      'pml',
    );
  });
});

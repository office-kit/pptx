// auditTextLayout — text overflow (はみ出し) and soft-wrap (段落ち) detection.
//
// Decks are built in memory with the public authoring API and measured with
// the bundled fontkit measurer (real glyph metrics), so the expectations are
// deterministic. Import pattern follows test/preview-svg-text-modes.test.ts.

import { readFile } from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  loadPresentation,
  inches,
  setShapeTextAutoFit,
  setShapeTextWrap,
} from '../src/api/index.ts';
import { auditTextLayout, type TextAuditIssue } from '../packages/preview/src/index.ts';
import { buildFontkitMeasurer, FONT_DIR } from '../packages/preview/src/node.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

const blankSlide = async () => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('Blank layout not found');
  return { pres, slide: addSlide(pres, { layout }) };
};

const measureText = buildFontkitMeasurer();

const ofShape = (issues: ReadonlyArray<TextAuditIssue>, name: string): TextAuditIssue[] =>
  issues.filter((i) => i.shapeName === name);

// One long paragraph — wraps to many lines in a 2" column at the default 18pt.
const LONG_PARAGRAPH = Array.from({ length: 30 }, (_, i) => `word${i + 1}`).join(' ');

describe('auditTextLayout — overflow', () => {
  it('reports nothing for text that fits its box', async () => {
    const { pres, slide } = await blankSlide();
    addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'Fits fine',
      name: 'fits',
    });
    expect(ofShape(auditTextLayout(pres, { measureText }), 'fits')).toEqual([]);
  });

  it('reports overflow-y when wrapped text is taller than the box', async () => {
    const { pres, slide } = await blankSlide();
    addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(0.5),
      text: LONG_PARAGRAPH,
      name: 'tall',
    });
    const issues = ofShape(auditTextLayout(pres, { measureText }), 'tall');
    const oy = issues.find((i) => i.kind === 'overflow-y');
    expect(oy).toBeDefined();
    expect(oy!.kind === 'overflow-y' && oy!.overflowPx).toBeGreaterThan(1);
    expect(oy!.approximate).toBe(false);
  });

  it('reports overflow-x for wrap="none" text wider than the box', async () => {
    const { pres, slide } = await blankSlide();
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(1),
      h: inches(1),
      text: 'This single unwrapped line is far wider than one inch',
      name: 'wide',
    });
    setShapeTextWrap(box, 'none');
    const issues = ofShape(auditTextLayout(pres, { measureText }), 'wide');
    expect(issues.some((i) => i.kind === 'overflow-x')).toBe(true);
  });

  it('honours normAutofit: the shrink-to-fit search suppresses the overflow', async () => {
    const mk = async () => {
      const { pres, slide } = await blankSlide();
      const box = addSlideTextBox(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(3),
        h: inches(0.6),
        text: LONG_PARAGRAPH,
        name: 'fit',
      });
      return { pres, box };
    };
    // Without autofit the box overflows …
    const plain = await mk();
    expect(
      ofShape(auditTextLayout(plain.pres, { measureText }), 'fit').some(
        (i) => i.kind === 'overflow-y',
      ),
    ).toBe(true);
    // … with a bare <a:normAutofit/> the audit shrinks like the renderer does.
    const fitted = await mk();
    setShapeTextAutoFit(fitted.box, 'normal');
    expect(
      ofShape(auditTextLayout(fitted.pres, { measureText }), 'fit').some(
        (i) => i.kind === 'overflow-y',
      ),
    ).toBe(false);
  });

  it('flags issues as approximate under the default heuristic measurer', async () => {
    const { pres, slide } = await blankSlide();
    addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(0.5),
      text: LONG_PARAGRAPH,
      name: 'heuristic',
    });
    const issues = ofShape(auditTextLayout(pres), 'heuristic');
    expect(issues.length).toBeGreaterThan(0);
    expect(issues.every((i) => i.approximate)).toBe(true);
  });
});

describe('auditTextLayout — soft wraps (段落ち)', () => {
  it('reports wrapped paragraphs only when opted in', async () => {
    const { pres, slide } = await blankSlide();
    addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(6), // tall enough that nothing overflows
      text: LONG_PARAGRAPH,
      name: 'wrapping',
    });
    expect(ofShape(auditTextLayout(pres, { measureText }), 'wrapping')).toEqual([]);
    const issues = ofShape(
      auditTextLayout(pres, { measureText, reportSoftWraps: true }),
      'wrapping',
    );
    expect(issues).toHaveLength(1);
    const wrap = issues[0]!;
    expect(wrap.kind).toBe('soft-wrap');
    if (wrap.kind === 'soft-wrap') {
      expect(wrap.paragraphIndex).toBe(0);
      expect(wrap.extraLines).toBeGreaterThan(0);
    }
  });

  it('does not count authored line breaks as soft wraps', async () => {
    const { pres, slide } = await blankSlide();
    addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(3),
      text: 'first\nsecond\nthird',
      name: 'authored-breaks',
    });
    expect(
      ofShape(auditTextLayout(pres, { measureText, reportSoftWraps: true }), 'authored-breaks'),
    ).toEqual([]);
  });
});

describe('buildFontkitMeasurer — registered fonts and glyph fallback', () => {
  const spec = { family: 'Carlito', sizePx: 24, bold: false, italic: false, letterSpacingPx: 0 };

  it('estimates CJK glyphs the bundled Latin fonts lack (1em, approximate)', () => {
    const r = measureText('あいう', spec);
    expect(r.approximate).toBe(true);
    expect(r.widthPx).toBeCloseTo(3 * spec.sizePx, 5);
    // Vertical metrics still come from a real font.
    expect(r.ascentPx).toBeGreaterThan(0);
  });

  it('measures fully covered text with real glyphs (no approximate flag)', () => {
    const r = measureText('Hello world', spec);
    expect(r.approximate).toBeUndefined();
    expect(r.widthPx).toBeGreaterThan(0);
  });

  it('resolves a registered font by its authored family name', () => {
    const m = buildFontkitMeasurer({
      fonts: [{ family: 'My Brand Serif', source: `${FONT_DIR}Caladea-Regular.ttf` }],
    });
    const branded = m('Measurement sample', { ...spec, family: 'My Brand Serif' });
    const serif = m('Measurement sample', { ...spec, family: 'Caladea' });
    const sans = m('Measurement sample', { ...spec, family: 'Carlito' });
    expect(branded.widthPx).toBeCloseTo(serif.widthPx, 5);
    expect(branded.widthPx).not.toBeCloseTo(sans.widthPx, 1);
  });

  it('accepts font bytes as the source', () => {
    const bytes = new Uint8Array(readFileSync(`${FONT_DIR}Caladea-Regular.ttf`));
    const m = buildFontkitMeasurer({ fonts: [{ family: 'From Bytes', source: bytes }] });
    const fromBytes = m('Measurement sample', { ...spec, family: 'From Bytes' });
    const serif = m('Measurement sample', { ...spec, family: 'Caladea' });
    expect(fromBytes.widthPx).toBeCloseTo(serif.widthPx, 5);
  });

  it('substitutes authored Office names like the render path does', () => {
    // The audit passes authored names through; unknown/Office names must land
    // on the same bundled faces the emitter substitutes.
    const arial = measureText('Measurement sample', { ...spec, family: 'Arial' });
    const liberation = measureText('Measurement sample', { ...spec, family: 'Liberation Sans' });
    expect(arial.widthPx).toBeCloseTo(liberation.widthPx, 5);
  });
});

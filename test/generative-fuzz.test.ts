// Generative schema-validity fuzzing — the exhaustive net behind the per-setter
// regression tests. It composes many random-but-valid decks from the public
// authoring API (random combinations of shapes, formatting, tables, charts,
// images, transitions, animations) and asserts that EVERY emitted part is
// schema-valid against the ECMA-376 XSDs. Hand-written examples cover known
// cases; this covers the combinatorial space where "feature A + feature B emits
// invalid child order / an attribute on the wrong type" bugs hide.
//
// Deterministic: a fixed seed drives a small PRNG, so a failure reproduces and
// CI is stable. The failing iteration's seed/index is reported. Bump ITERATIONS
// locally to fuzz harder.

import { describe, it } from 'vitest';
import {
  _internalPackageOf,
  addBlankSlide,
  addContentSlide,
  addSlideChart,
  addSlideImage,
  addSlideLine,
  addSlideShape,
  addSlideTable,
  addSlideTextBox,
  addTitleSlide,
  createPresentation,
  getTableCell,
  inches,
  loadPresentation,
  pt,
  type PresentationData,
  savePresentation,
  setShapeAnimation,
  setShapeFill,
  setShapeGradientFill,
  setShapePatternFill,
  setShapeRunFormat,
  setShapeShadow,
  setShapeStroke,
  setShapeTextColumns,
  setSlideNotes,
  setSlideTransition,
  setTableCellBorders,
  setTableCellFill,
} from '../src/api/index.ts';
import { buildPng } from './lib/build-png.ts';
import {
  expectSchemaValid,
  isSchemaValidationAvailable,
  type SchemaKind,
} from './lib/expect-schema-valid.ts';

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;
const decode = (b: Uint8Array): string => new TextDecoder().decode(b);

const kindFor = (name: string): SchemaKind | null => {
  if (/\/(slides|slideLayouts|slideMasters|notesSlides|notesMasters)\/[^/]*\.xml$/.test(name)) {
    return 'pml';
  }
  if (name === '/ppt/presentation.xml') return 'pml';
  if (/\/charts\/chart\d+\.xml$/.test(name)) return 'chart';
  if (/\/theme\/theme\d+\.xml$/.test(name)) return 'dml';
  return null;
};

// mulberry32 — tiny deterministic PRNG so failures reproduce.
const makeRng = (seed: number): (() => number) => {
  let a = seed >>> 0;
  return () => {
    a |= 0;
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
};

const HEX = '0123456789abcdef';
const SCHEME = ['accent1', 'accent2', 'accent3', 'tx1', 'bg1', 'dk1', 'lt1'];
const PATTERNS = ['pct50', 'dkUpDiag', 'wave', 'cross', 'horzBrick', 'zigZag'] as const;
const TRANSITIONS = [
  { effect: 'fade' as const },
  { effect: 'fade' as const, thruBlack: true },
  { effect: 'push' as const, direction: 'l' },
  { effect: 'wipe' as const, direction: 'd' },
  { effect: 'blinds' as const, direction: 'horz' },
  { effect: 'split' as const, orientation: 'horz' as const },
  { effect: 'cover' as const, direction: 'ru' },
  { effect: 'zoom' as const, direction: 'in' },
  { effect: 'dissolve' as const },
  { effect: 'none' as const },
];
const ANIMS = ['fadeIn', 'fadeOut', 'appear', 'disappear'] as const;
// Authorable chart kinds (scatter / radar / bubble are read-only for now).
const CHART_KINDS = ['bar', 'column', 'line', 'pie', 'doughnut', 'area'] as const;

const validateDeck = (pres: PresentationData): void => {
  const pkg = _internalPackageOf(pres);
  for (const part of pkg.parts) {
    const kind = kindFor(part.name);
    if (kind) expectSchemaValid(decode(part.data), kind);
  }
};

describe('generative fuzz: every authored part is schema-valid', () => {
  // Each iteration saves a deck and shells out to xmllint per part, so keep the
  // count modest for CI; bump locally to fuzz harder. Generous timeout covers
  // the per-part process spawns.
  const ITERATIONS = 40;

  skipIfNoXmllint(
    'random valid decks validate against the ECMA-376 XSDs',
    async () => {
      for (let iter = 0; iter < ITERATIONS; iter++) {
        const seed = 0x9e3779b9 ^ (iter * 0x85ebca77);
        const rng = makeRng(seed);
        const pick = <T>(arr: ReadonlyArray<T>): T => arr[Math.floor(rng() * arr.length)]!;
        const chance = (p: number): boolean => rng() < p;
        const color = (): string => {
          if (chance(0.3)) {
            const token = pick(SCHEME);
            // Exercise both spellings the API accepts: bare and `scheme:`-prefixed
            // (the latter is what the read-back getters emit, so it must round-trip).
            return chance(0.5) ? `scheme:${token}` : token;
          }
          let s = '#';
          const n = chance(0.5) ? 3 : 6;
          for (let i = 0; i < n; i++) s += HEX[Math.floor(rng() * 16)];
          return s;
        };
        const emu = (lo: number, hi: number) => inches(lo + rng() * (hi - lo));

        let pres: PresentationData;
        try {
          pres = createPresentation(chance(0.5) ? { size: '16:9' } : { size: '4:3' });

          const slideCount = 1 + Math.floor(rng() * 3);
          for (let s = 0; s < slideCount; s++) {
            const flavor = rng();
            if (flavor < 0.2) {
              addTitleSlide(pres, `Deck ${iter} slide ${s}`);
              continue;
            }
            if (flavor < 0.35) {
              addContentSlide(pres, { title: `T${s}` });
              continue;
            }
            const slide = addBlankSlide(pres);

            // A handful of random shapes with stacked random formatting.
            const shapeCount = 1 + Math.floor(rng() * 4);
            for (let k = 0; k < shapeCount; k++) {
              const hasText = chance(0.7);
              const sh = addSlideShape(slide, {
                preset: pick([
                  'rect',
                  'roundRect',
                  'ellipse',
                  'triangle',
                  'diamond',
                  'mathMinus',
                  'mathMultiply',
                  'mathNotEqual',
                ]),
                x: emu(0.2, 4),
                y: emu(0.2, 3),
                w: emu(0.5, 4),
                h: emu(0.5, 3),
                ...(hasText ? { text: `s${k}` } : {}),
                textAnchor: pick(['t', 'ctr', 'b']),
              });
              const fillRoll = rng();
              if (fillRoll < 0.4) {
                setShapeFill(sh, color());
              } else if (fillRoll < 0.7) {
                setShapeGradientFill(sh, {
                  stops: [
                    { offset: 0, color: color() },
                    { offset: 1, color: color() },
                  ],
                  angleDeg: Math.floor(rng() * 360),
                  ...(chance(0.4) ? { path: pick(['circle', 'rect'] as const) } : {}),
                });
              } else {
                setShapePatternFill(sh, {
                  preset: pick(PATTERNS),
                  foreground: color(),
                  background: color(),
                });
              }
              if (chance(0.5)) {
                setShapeStroke(sh, {
                  color: color(),
                  widthEmu: pt(0.5 + rng() * 4),
                  ...(chance(0.5) ? { dash: pick(['dash', 'dot', 'dashDot']) } : {}),
                });
              }
              if (chance(0.4)) {
                setShapeShadow(sh, {
                  color: '#000000',
                  blurEmu: pt(rng() * 10),
                  offsetEmu: pt(rng() * 6),
                  angleDeg: Math.floor(rng() * 360),
                  opacity: rng(),
                });
              }
              if (hasText && chance(0.6)) {
                setShapeRunFormat(sh, 0, 0, {
                  bold: chance(0.5),
                  italic: chance(0.5),
                  size: 8 + Math.floor(rng() * 40),
                  color: color(),
                  ...(chance(0.3) ? { underline: true } : {}),
                  ...(chance(0.3) ? { highlight: color() } : {}),
                });
              }
              if (hasText && chance(0.2))
                setShapeTextColumns(sh, { count: 2 + Math.floor(rng() * 3) });
              if (chance(0.3)) setShapeAnimation(sh, { effect: pick(ANIMS) });
            }

            // Random extras.
            if (chance(0.4)) {
              addSlideTextBox(slide, {
                x: emu(0.3, 3),
                y: emu(0.3, 3),
                w: emu(1, 4),
                h: emu(0.5, 2),
                text: `box ${iter}-${s}`,
              });
            }
            if (chance(0.4)) {
              addSlideLine(slide, {
                from: { x: emu(0.2, 2), y: emu(0.2, 2) },
                to: { x: emu(2, 5), y: emu(2, 4) },
                ...(chance(0.6) ? { color: color(), widthEmu: pt(0.5 + rng() * 3) } : {}),
              });
            }
            if (chance(0.35)) {
              const t = addSlideTable(slide, {
                x: emu(0.3, 2),
                y: emu(0.3, 2),
                w: emu(3, 7),
                h: emu(1, 3),
                rows: [
                  ['H1', 'H2', 'H3'],
                  ['a', 'b', 'c'],
                  ['d', 'e', 'f'],
                ],
                firstRow: chance(0.7),
                bandRow: chance(0.7),
              });
              if (chance(0.5)) {
                const cell = getTableCell(t, 1, 1);
                setTableCellFill(cell, color());
                if (chance(0.5)) {
                  setTableCellBorders(cell, {
                    top: { color: color(), widthEmu: pt(0.5 + rng() * 2) },
                    bottom: { color: color(), widthEmu: pt(0.5 + rng() * 2) },
                  });
                }
              }
            }
            if (chance(0.3)) {
              const kind = pick(CHART_KINDS);
              // pie / doughnut take exactly one series.
              const single = kind === 'pie' || kind === 'doughnut';
              const series = single
                ? [{ name: 'A', values: [1, 2, 3, 4], ...(chance(0.5) ? { color: '4472C4' } : {}) }]
                : [
                    {
                      name: 'A',
                      values: [1, 2, 3, 4],
                      ...(chance(0.5) ? { color: '4472C4' } : {}),
                    },
                    { name: 'B', values: [4, 3, 2, 1] },
                  ];
              addSlideChart(slide, {
                x: emu(0.3, 2),
                y: emu(0.3, 2),
                w: emu(4, 7),
                h: emu(2, 4),
                spec: {
                  kind,
                  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
                  series,
                  // Random-but-valid percentages for the kinds that take them, so
                  // the bounds wrappers (gap 0..500, hole 1..90, angle 0..360) are
                  // exercised at real values rather than only at the defaults.
                  ...((kind === 'bar' || kind === 'column') && chance(0.5)
                    ? { gapWidthPct: Math.floor(rng() * 501) }
                    : {}),
                  ...(kind === 'doughnut' && chance(0.5)
                    ? { holeSizePct: 1 + Math.floor(rng() * 90) }
                    : {}),
                  ...((kind === 'pie' || kind === 'doughnut') && chance(0.5)
                    ? { firstSliceAngleDeg: Math.floor(rng() * 361) }
                    : {}),
                  ...(chance(0.5)
                    ? {
                        dataLabels: {
                          showValue: true,
                          showCategory: false,
                          showSeriesName: false,
                          showPercent: false,
                        },
                      }
                    : {}),
                },
              });
            }
            if (chance(0.3)) {
              addSlideImage(slide, buildPng(32, 24, [Math.floor(rng() * 255), 10, 200]), {
                x: emu(0.3, 3),
                y: emu(0.3, 3),
                w: emu(1, 3),
                h: emu(1, 3),
                fit: pick(['contain', 'fill'] as const),
              });
            }
            if (chance(0.5)) setSlideTransition(slide, pick(TRANSITIONS));
            if (chance(0.3)) setSlideNotes(slide, `notes for ${iter}-${s} <&> "ok"`);
          }
        } catch (err) {
          // A thrown boundary error is acceptable (rejecting bad input is correct);
          // only schema-INVALID *output* is a failure. But our generators only emit
          // valid ranges, so a throw here is an unexpected bug — surface it.
          throw new Error(`iteration ${iter} (seed ${seed}) threw unexpectedly: ${String(err)}`);
        }

        try {
          const bytes = await savePresentation(pres);
          validateDeck(await loadPresentation(bytes));
        } catch (err) {
          throw new Error(
            `iteration ${iter} (seed ${seed}) produced invalid output: ${String(err)}`,
          );
        }
      }
    },
    120_000,
  );
});

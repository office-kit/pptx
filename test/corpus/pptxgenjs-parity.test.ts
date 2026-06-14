// PptxGenJS-parity corpus.
//
// PptxGenJS is the most widely-used PPTX generator in the JS ecosystem and its
// output is battle-tested to open cleanly in PowerPoint, Keynote, Google
// Slides, and LibreOffice. This suite authors the *same* slide through both
// PptxGenJS and pptx-kit and compares the result, turning "is pptx-kit's
// generated output any good?" into two concrete, falsifiable checks:
//
//   1. HARD GATE — every slide pptx-kit emits is schema-valid against the
//      ECMA-376 PresentationML XSD. A real structural defect fails the build.
//
//   2. RATCHET — the canonical drawing tree is diffed against PptxGenJS's and
//      the number of divergent lines is compared to a committed baseline. The
//      count may only shrink; a change that pushes pptx-kit further from the
//      reference fails. Record a new baseline with `CORPUS_RECORD=1`.
//
// The residual divergence is itself the quality signal: every line in the
// per-case report is either a pptx-kit gap to close or a documented, accepted
// stylistic difference (see `parity-baseline.json` notes and the harness
// README). Not every divergence is a bug — PptxGenJS hard-codes black run
// colors and vertical-center anchors where pptx-kit inherits from the theme,
// and pptx-kit marks real text boxes with `txBox="1"` where PptxGenJS does
// not — but the harness keeps all of them visible and counted.

import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { afterAll, describe, expect, it } from 'vitest';
import { expectSchemaValid, isSchemaValidationAvailable } from '../lib/expect-schema-valid.ts';
import { diffLines, divergenceCount } from './canonical.ts';
import { CASES } from './cases.ts';
import { pptxGenJsAvailable, runCase } from './harness.ts';

const BASELINE_PATH = fileURLToPath(new URL('./parity-baseline.json', import.meta.url));
const OUT_DIR = fileURLToPath(new URL('./out/', import.meta.url));
const RECORD = process.env.CORPUS_RECORD === '1';

type Baseline = Record<string, number>;

const loadBaseline = (): Baseline => {
  try {
    return JSON.parse(readFileSync(BASELINE_PATH, 'utf8')) as Baseline;
  } catch {
    return {};
  }
};

const baseline = loadBaseline();
const recorded: Baseline = {};
const reportSections: string[] = [];

describe.skipIf(!pptxGenJsAvailable())('PptxGenJS parity corpus', () => {
  for (const c of CASES) {
    it(c.id, async () => {
      const r = await runCase(c);

      // 1. Hard gate: pptx-kit's slide must be schema-valid.
      if (isSchemaValidationAvailable()) {
        expectSchemaValid(r.kitXml, 'pml');
      }

      // 2. Ratchet: divergence from the reference may not grow.
      const n = divergenceCount(r.kitCanonical, r.pgjsCanonical);
      recorded[c.id] = n;

      const diff = diffLines(r.kitCanonical, r.pgjsCanonical);
      reportSections.push(
        `## ${c.id} — divergence ${n}\n\n${diff ? `\`\`\`diff\n${diff}\n\`\`\`` : '_identical_'}\n`,
      );

      if (!RECORD) {
        const limit = baseline[c.id];
        expect(
          limit,
          `no parity baseline for "${c.id}" — run \`CORPUS_RECORD=1 pnpm test\` to record one`,
        ).toBeTypeOf('number');
        expect(
          n,
          `parity regressed for "${c.id}": divergence ${n} > baseline ${limit}. ` +
            'See test/corpus/out/report.md for the diff.',
        ).toBeLessThanOrEqual(limit!);
      }
    });
  }

  afterAll(() => {
    mkdirSync(OUT_DIR, { recursive: true });
    const total = Object.values(recorded).reduce((a, b) => a + b, 0);
    writeFileSync(
      `${OUT_DIR}report.md`,
      `# PptxGenJS parity report\n\nTotal divergence across ${
        Object.keys(recorded).length
      } cases: **${total}**\n\n${reportSections.join('\n')}`,
    );
    if (RECORD) {
      const sorted = Object.fromEntries(
        Object.entries(recorded).sort(([a], [b]) => a.localeCompare(b)),
      );
      writeFileSync(BASELINE_PATH, `${JSON.stringify(sorted, null, 2)}\n`);
    }
  });
});

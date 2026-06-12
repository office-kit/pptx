// Baseline management for the preview-fidelity regression gate.
//
// The committed baseline.json records per-slide fg-SSIM scores from a known-
// good LibreOffice run. The gate compares every CI run against it:
//
//   --record   write a new baseline.json after a successful run
//   --check    compare the run against baseline.json; exit 1 on regression
//
// The comparison logic (compareWithBaseline) is pure so it can be unit-tested
// without LibreOffice. See test/fidelity-baseline.test.ts.

import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { resolveSoffice } from './ground-truth.ts';
import type { FileReport } from './report.ts';

// LibreOffice point-release rendering jitter moves absolute fg-SSIM scores by
// up to ~0.01 per release; 0.02 gives one full release of headroom without
// masking real regressions.
export const TOLERANCE = 0.02;

export interface BaselineData {
  readonly engine: string;
  readonly width: number;
  readonly sofficeVersion: string;
  readonly files: Record<string, number[]>;
}

export interface CompareFailure {
  readonly file: string;
  /** 1-indexed slide number; 0 means the mismatch applies to the whole file. */
  readonly slide: number;
  readonly reason: 'regression' | 'missing-in-baseline' | 'extra-in-baseline';
  readonly actual?: number;
  readonly baseline?: number;
  readonly delta?: number;
}

export interface CompareImprovement {
  readonly file: string;
  readonly slide: number;
  readonly actual: number;
  readonly baseline: number;
  readonly delta: number;
}

export interface CompareResult {
  readonly pass: boolean;
  readonly failures: readonly CompareFailure[];
  readonly improvements: readonly CompareImprovement[];
}

/**
 * Compare per-slide fgSsim scores from a run against the committed baseline.
 * Pure: no I/O, safe to unit-test without LibreOffice.
 */
export const compareWithBaseline = (
  baseline: BaselineData,
  run: { readonly files: Record<string, number[]> },
): CompareResult => {
  const failures: CompareFailure[] = [];
  const improvements: CompareImprovement[] = [];

  // Files present in baseline but absent from the run.
  for (const [file] of Object.entries(baseline.files).sort(([a], [b]) => a.localeCompare(b))) {
    if (!(file in run.files)) {
      failures.push({ file, slide: 0, reason: 'extra-in-baseline' });
    }
  }

  // Iterate run files in sorted order for deterministic output.
  for (const [file, slides] of Object.entries(run.files).sort(([a], [b]) => a.localeCompare(b))) {
    const baseSlides = baseline.files[file];

    if (baseSlides === undefined) {
      // File is new — no baseline entry to compare against.
      for (let i = 0; i < slides.length; i++) {
        failures.push({ file, slide: i + 1, reason: 'missing-in-baseline', actual: slides[i]! });
      }
      continue;
    }

    const maxLen = Math.max(slides.length, baseSlides.length);
    for (let i = 0; i < maxLen; i++) {
      const actual = slides[i];
      const base = baseSlides[i];

      if (actual === undefined) {
        // Baseline has a slide that no longer exists in the run. `base` is
        // defined here: i < maxLen with actual missing implies i < baseSlides.length.
        failures.push({ file, slide: i + 1, reason: 'extra-in-baseline', baseline: base! });
      } else if (base === undefined) {
        // New slide in the run has no baseline entry.
        failures.push({ file, slide: i + 1, reason: 'missing-in-baseline', actual });
      } else {
        // Round delta to 4 dp to match the precision of the stored fgSsim
        // scores, eliminating floating-point drift when comparing two 4-dp values
        // (e.g. 0.78 - 0.80 computes as -0.020000000000000062 in IEEE 754).
        const delta = parseFloat((actual - base).toFixed(4));
        if (delta < -TOLERANCE) {
          failures.push({
            file,
            slide: i + 1,
            reason: 'regression',
            actual,
            baseline: base,
            delta,
          });
        } else if (delta > TOLERANCE) {
          improvements.push({ file, slide: i + 1, actual, baseline: base, delta });
        }
      }
    }
  }

  return { pass: failures.length === 0, failures, improvements };
};

/** Run `soffice --version` via the resolved binary and return the trimmed string. */
export const getSofficeVersion = (): string => {
  const bin = resolveSoffice();
  const out = execFileSync(bin, ['--version'], { encoding: 'utf8' });
  return out.trim();
};

/** Build the `files` map from completed FileReports, rounding fgSsim to 4 dp. */
export const reportsToRunFiles = (reports: FileReport[]): Record<string, number[]> => {
  const entries: [string, number[]][] = reports.map((r) => [
    r.name,
    r.slides.map((s) => parseFloat(s.fgSsim.toFixed(4))),
  ]);
  // Sort by filename for a stable, human-readable baseline.
  entries.sort(([a], [b]) => a.localeCompare(b));
  return Object.fromEntries(entries);
};

export const loadBaseline = (path: string): BaselineData | null => {
  if (!existsSync(path)) return null;
  return JSON.parse(readFileSync(path, 'utf8')) as BaselineData;
};

// Hand-rolled serialization instead of JSON.stringify(data, null, 2): oxfmt
// reformats committed JSON with per-file score arrays inline, and emitting
// that shape directly lets a CI baseline candidate be committed unchanged.
export const writeBaselineFile = (path: string, data: BaselineData): void => {
  const fileLines = Object.entries(data.files)
    .map(([name, scores]) => `    ${JSON.stringify(name)}: [${scores.join(', ')}]`)
    .join(',\n');
  const json = [
    '{',
    `  "engine": ${JSON.stringify(data.engine)},`,
    `  "width": ${data.width},`,
    `  "sofficeVersion": ${JSON.stringify(data.sofficeVersion)},`,
    '  "files": {',
    fileLines,
    '  }',
    '}',
  ].join('\n');
  writeFileSync(path, json + '\n');
};

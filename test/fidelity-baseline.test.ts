// Unit coverage for the fidelity-gate baseline comparison logic. The function
// under test (compareWithBaseline) is pure — no I/O, no LibreOffice — so these
// run in CI without external renderers. See site/fidelity/README.md.

import { describe, expect, it } from 'vitest';
import { compareWithBaseline, TOLERANCE, type BaselineData } from '../site/fidelity/baseline.ts';

const makeBaseline = (files: Record<string, number[]>): BaselineData => ({
  engine: 'libreoffice',
  width: 1280,
  sofficeVersion: 'LibreOffice 7.5.0.3',
  files,
});

describe('compareWithBaseline', () => {
  it('passes when all slides are within tolerance', () => {
    const result = compareWithBaseline(makeBaseline({ 'a.pptx': [0.8, 0.9] }), {
      files: { 'a.pptx': [0.8, 0.9] },
    });
    expect(result.pass).toBe(true);
    expect(result.failures).toHaveLength(0);
    expect(result.improvements).toHaveLength(0);
  });

  it('passes at exactly -TOLERANCE (boundary: strictly-less-than, not regression)', () => {
    // delta = (0.8 - TOLERANCE) - 0.8 = -TOLERANCE; condition is delta < -TOLERANCE, which is false.
    const result = compareWithBaseline(makeBaseline({ 'a.pptx': [0.8] }), {
      files: { 'a.pptx': [0.8 - TOLERANCE] },
    });
    expect(result.pass).toBe(true);
    expect(result.failures).toHaveLength(0);
  });

  it('fails one tick below -TOLERANCE', () => {
    const result = compareWithBaseline(makeBaseline({ 'a.pptx': [0.8] }), {
      files: { 'a.pptx': [0.8 - TOLERANCE - 0.0001] },
    });
    expect(result.pass).toBe(false);
    expect(result.failures).toHaveLength(1);
    expect(result.failures[0]?.reason).toBe('regression');
  });

  it('regression failure carries correct actual, baseline, and delta', () => {
    const result = compareWithBaseline(makeBaseline({ 'a.pptx': [0.9] }), {
      files: { 'a.pptx': [0.7] },
    });
    expect(result.pass).toBe(false);
    const f = result.failures[0];
    expect(f?.reason).toBe('regression');
    expect(f?.slide).toBe(1);
    expect(f?.actual).toBeCloseTo(0.7, 4);
    expect(f?.baseline).toBeCloseTo(0.9, 4);
    expect(f?.delta).toBeCloseTo(-0.2, 4);
  });

  it('fails when run contains a file not present in baseline', () => {
    const result = compareWithBaseline(makeBaseline({ 'a.pptx': [0.8] }), {
      files: { 'a.pptx': [0.8], 'b.pptx': [0.7] },
    });
    expect(result.pass).toBe(false);
    expect(
      result.failures.some((f) => f.file === 'b.pptx' && f.reason === 'missing-in-baseline'),
    ).toBe(true);
  });

  it('fails when baseline contains a file absent from the run', () => {
    const result = compareWithBaseline(makeBaseline({ 'a.pptx': [0.8], 'b.pptx': [0.7] }), {
      files: { 'a.pptx': [0.8] },
    });
    expect(result.pass).toBe(false);
    expect(
      result.failures.some((f) => f.file === 'b.pptx' && f.reason === 'extra-in-baseline'),
    ).toBe(true);
  });

  it('fails when run has more slides than baseline', () => {
    const result = compareWithBaseline(makeBaseline({ 'a.pptx': [0.8] }), {
      files: { 'a.pptx': [0.8, 0.9] },
    });
    expect(result.pass).toBe(false);
    expect(
      result.failures.some(
        (f) => f.file === 'a.pptx' && f.slide === 2 && f.reason === 'missing-in-baseline',
      ),
    ).toBe(true);
  });

  it('fails when baseline has more slides than run', () => {
    const result = compareWithBaseline(makeBaseline({ 'a.pptx': [0.8, 0.9] }), {
      files: { 'a.pptx': [0.8] },
    });
    expect(result.pass).toBe(false);
    expect(
      result.failures.some(
        (f) => f.file === 'a.pptx' && f.slide === 2 && f.reason === 'extra-in-baseline',
      ),
    ).toBe(true);
  });

  it('reports improvements for slides more than TOLERANCE above baseline', () => {
    const result = compareWithBaseline(makeBaseline({ 'a.pptx': [0.6] }), {
      files: { 'a.pptx': [0.8] },
    });
    // Improvements do not fail the gate.
    expect(result.pass).toBe(true);
    expect(result.improvements).toHaveLength(1);
    const imp = result.improvements[0];
    expect(imp?.file).toBe('a.pptx');
    expect(imp?.slide).toBe(1);
    expect(imp?.delta).toBeCloseTo(0.2, 4);
  });

  it('does not report improvement at exactly TOLERANCE above baseline', () => {
    // delta = TOLERANCE; condition is delta > TOLERANCE, which is false.
    const result = compareWithBaseline(makeBaseline({ 'a.pptx': [0.6] }), {
      files: { 'a.pptx': [0.6 + TOLERANCE] },
    });
    expect(result.pass).toBe(true);
    expect(result.improvements).toHaveLength(0);
  });

  it('handles multiple files and slides with mixed outcomes', () => {
    const baseline = makeBaseline({
      'a.pptx': [0.8, 0.9],
      'b.pptx': [0.7],
    });
    const run = {
      files: {
        'a.pptx': [0.75, 0.5], // slide 2 regresses (0.5 - 0.9 = -0.4 < -0.02)
        'b.pptx': [0.95], // improvement (0.95 - 0.7 = 0.25 > 0.02)
      },
    };
    const result = compareWithBaseline(baseline, run);
    expect(result.pass).toBe(false);
    expect(
      result.failures.some(
        (f) => f.file === 'a.pptx' && f.slide === 2 && f.reason === 'regression',
      ),
    ).toBe(true);
    expect(result.improvements.some((i) => i.file === 'b.pptx' && i.slide === 1)).toBe(true);
  });
});

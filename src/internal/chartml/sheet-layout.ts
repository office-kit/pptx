// Where each chart channel lives in the embedded workbook.
//
// The chart XML and the workbook must agree cell for cell: every `<c:f>`
// formula in the chart names a range, and PowerPoint's "Edit data" rewrites
// the chart from whatever it finds there. Computing the grid and the
// formulas in one place is what keeps the two from drifting.

import { type SheetCell, cellAddr, cellRange } from './embedded-xlsx.ts';
import type { ReadChartSpec } from './types.ts';

const SHEET = 'Sheet1';

/** `<c:f>` formulas of one series' channels. */
export interface SeriesSheetRefs {
  /** The header cell holding the series name (`<c:tx>`). */
  readonly name: string;
  /** `<c:val>` / `<c:yVal>`. */
  readonly values: string;
  /** `<c:xVal>` — xy kinds only. */
  readonly xValues?: string;
  /** `<c:bubbleSize>` — bubble only. */
  readonly bubbleSizes?: string;
}

export interface ChartSheetLayout {
  readonly grid: ReadonlyArray<ReadonlyArray<SheetCell>>;
  /**
   * `<c:cat>` range shared by every series. Spans all the level columns
   * for a multi-level axis. Absent for the xy kinds, which have no
   * categories.
   */
  readonly categories?: string;
  readonly series: ReadonlyArray<SeriesSheetRefs>;
}

const isXyKind = (spec: ReadChartSpec): boolean =>
  spec.kind === 'scatter' || spec.kind === 'bubble';

/** Parses a date-axis category; the caller has validated it is numeric. */
export const dateSerial = (category: string): number => Number(category);

// Category kinds — the layout Excel itself produces for a chart sheet:
//
//   |  (level columns, outermost first)  | series 0 | series 1 | …
//   |  2024  |  Q1                       |    10    |    7     |
//   |        |  Q2                       |    12    |    9     |
const layoutCategorySheet = (spec: ReadChartSpec): ChartSheetLayout => {
  const outerLevels = [...(spec.categoryGroupLevels ?? [])].reverse();
  const levelCount = outerLevels.length + 1;
  const rowCount = spec.categories.length;
  const isDateAxis = spec.categoryAxisDate !== undefined;

  // A1 (and the rest of the level header) is written as an empty string
  // rather than omitted, the way Excel marks the corner of a chart range.
  const header: SheetCell[] = [
    ...Array.from({ length: levelCount }, () => ''),
    ...spec.series.map((s) => s.name),
  ];
  const grid: SheetCell[][] = [header];
  for (let r = 0; r < rowCount; r++) {
    const category = spec.categories[r] ?? '';
    const row: SheetCell[] = outerLevels.map((level) => {
      const label = level[r] ?? '';
      return label === '' ? null : label;
    });
    row.push(isDateAxis ? dateSerial(category) : category);
    for (const s of spec.series) row.push(s.values[r] ?? null);
    grid.push(row);
  }

  return {
    grid,
    categories: cellRange(SHEET, 1, 0, rowCount, levelCount - 1),
    series: spec.series.map((_, i) => ({
      name: cellAddr(SHEET, 0, levelCount + i),
      values: cellRange(SHEET, 1, levelCount + i, rowCount),
    })),
  };
};

// xy kinds — each series owns its x channel, so each gets its own block of
// columns (x, y[, size]). The series name heads the y column because that
// is the cell `<c:tx>` points at.
const layoutXySheet = (spec: ReadChartSpec): ChartSheetLayout => {
  const isBubble = spec.kind === 'bubble';
  const stride = isBubble ? 3 : 2;
  const rowCount = Math.max(
    0,
    ...spec.series.map((s) =>
      Math.max(s.values.length, s.xValues?.length ?? 0, s.bubbleSizes?.length ?? 0),
    ),
  );
  const header: SheetCell[] = [];
  for (const s of spec.series) {
    header.push('X', s.name);
    if (isBubble) header.push('Size');
  }
  const grid: SheetCell[][] = [header];
  for (let r = 0; r < rowCount; r++) {
    const row: SheetCell[] = [];
    for (const s of spec.series) {
      row.push(s.xValues?.[r] ?? null, s.values[r] ?? null);
      if (isBubble) row.push(s.bubbleSizes?.[r] ?? null);
    }
    grid.push(row);
  }
  return {
    grid,
    series: spec.series.map((s, i) => {
      const base = i * stride;
      const pointCount = s.values.length;
      return {
        name: cellAddr(SHEET, 0, base + 1),
        xValues: cellRange(SHEET, 1, base, pointCount),
        values: cellRange(SHEET, 1, base + 1, pointCount),
        ...(isBubble ? { bubbleSizes: cellRange(SHEET, 1, base + 2, pointCount) } : {}),
      };
    }),
  };
};

/** Lays `spec` out as a sheet grid plus the formulas that address it. */
export const layoutChartSheet = (spec: ReadChartSpec): ChartSheetLayout =>
  isXyKind(spec) ? layoutXySheet(spec) : layoutCategorySheet(spec);

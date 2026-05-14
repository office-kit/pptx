// Chart authoring types.
//
// Minimum-viable chart surface: one or more named series of numeric
// values plotted against shared string categories. Sufficient for the
// three most-requested chart types — bar, line, pie — and structured so
// that scatter / area / radar can extend it later without breaking the
// public shape.

/** Chart type tokens we support today. */
export type ChartKind = 'bar' | 'column' | 'line' | 'pie';

/** One labelled series of numeric values. */
export interface ChartSeries {
  /** Series label rendered in the legend. */
  readonly name: string;
  /**
   * Numeric values, one per category. `null` slots become empty cells in
   * the embedded workbook (PowerPoint draws them as a gap). Lengths
   * shorter than the category count are right-padded with `null`.
   */
  readonly values: ReadonlyArray<number | null>;
  /** Optional `#RRGGBB` fill override. Defaults to the theme's accent palette. */
  readonly color?: string;
}

/** Full chart specification. */
export interface ChartSpec {
  readonly kind: ChartKind;
  /** Category labels along the x-axis (or pie slice labels for `kind: 'pie'`). */
  readonly categories: ReadonlyArray<string>;
  readonly series: ReadonlyArray<ChartSeries>;
  /** Optional chart title rendered above the plot area. */
  readonly title?: string;
}

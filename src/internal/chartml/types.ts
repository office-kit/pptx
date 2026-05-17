// Chart authoring types.
//
// Minimum-viable chart surface: one or more named series of numeric
// values plotted against shared string categories. Sufficient for the
// three most-requested chart types — bar, line, pie — and structured so
// that scatter / area / radar can extend it later without breaking the
// public shape.

/** Chart type tokens we support today. */
export type ChartKind = 'bar' | 'column' | 'line' | 'pie' | 'doughnut' | 'area';

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
  /**
   * Optional per-data-point color overrides, indexed by point index
   * (`<c:dPt><c:idx val="N"/><c:spPr><a:solidFill>…`). Sparse — only
   * the indices that author an override appear. Pie / doughnut decks
   * almost always emit one of these per slice to break out of the
   * single-series-color default.
   */
  readonly pointColors?: ReadonlyArray<string | null>;
  /**
   * Line-smoothing toggle (`<c:smooth val="1"/>`) — only meaningful for
   * line / scatter / area series. When `true`, the renderer interpolates
   * a smooth curve through the data points instead of straight segments.
   */
  readonly smooth?: boolean;
  /**
   * Optional trendline overlay. ECMA-376 §21.2.2.211 allows several
   * regression types; we surface the most common subset. The line is
   * painted on top of the series in the renderer.
   */
  readonly trendline?: ChartTrendline;
}

/** A single trendline overlay for a series. */
export interface ChartTrendline {
  /** Regression type — linear / exp / log / poly / power / movingAvg. */
  readonly type: 'linear' | 'exp' | 'log' | 'poly' | 'power' | 'movingAvg';
  /** Optional moving-average period (only meaningful for type='movingAvg'). */
  readonly period?: number;
  /** Polynomial order (only meaningful for type='poly'). */
  readonly order?: number;
  /** Override stroke color; defaults to the series color. */
  readonly color?: string;
}

/**
 * Per-series data-label toggles read from `<c:dLbls>` (ECMA-376
 * §21.2.2.55). All four toggles default to `false` — renderers paint
 * labels only when the corresponding flag is `true`.
 */
export interface ChartDataLabels {
  /** Numeric value of each data point. */
  readonly showValue: boolean;
  /** Category label of each data point. */
  readonly showCategory: boolean;
  /** Series name on each data point. */
  readonly showSeriesName: boolean;
  /** Percentage of total (for pie / doughnut). */
  readonly showPercent: boolean;
}

/**
 * Authored value-axis scaling (`<c:valAx><c:scaling><c:min/>/<c:max/></c:scaling>`)
 * plus tick-spacing hints. When omitted, renderers compute the range
 * and tick spacing from the series values.
 */
export interface ChartAxisScaling {
  readonly min?: number;
  readonly max?: number;
  /** Major tick spacing (`<c:majorUnit val="N"/>`). */
  readonly majorUnit?: number;
  /** Minor tick spacing (`<c:minorUnit val="N"/>`). */
  readonly minorUnit?: number;
}

/**
 * Bar / column grouping per ECMA-376 §21.2.2.76 (`ST_BarGrouping`):
 *
 *   - `clustered` — bars within a category sit side-by-side (the default).
 *   - `stacked` — series values sit on top of each other; the y-axis spans
 *     0..max(sum of each category).
 *   - `percentStacked` — series values normalize to 100% per category.
 *   - `standard` — only meaningful for the 3D variants; renderers treat
 *     it as `clustered`.
 */
export type ChartGrouping = 'clustered' | 'stacked' | 'percentStacked' | 'standard';

/** Full chart specification. */
export interface ChartSpec {
  readonly kind: ChartKind;
  /** Category labels along the x-axis (or pie slice labels for `kind: 'pie'`). */
  readonly categories: ReadonlyArray<string>;
  readonly series: ReadonlyArray<ChartSeries>;
  /** Optional chart title rendered above the plot area. */
  readonly title?: string;
  /** Optional chart-level data-label toggles. */
  readonly dataLabels?: ChartDataLabels;
  /** Optional value-axis scaling override (min / max). */
  readonly valueAxis?: ChartAxisScaling;
  /** Optional axis title text (`<c:catAx><c:title>` / `<c:valAx><c:title>`). */
  readonly categoryAxisTitle?: string;
  readonly valueAxisTitle?: string;
  /** Bar / column / area grouping mode. Absent for line / pie. */
  readonly grouping?: ChartGrouping;
  /**
   * Gap between adjacent bar groups in `<c:gapWidth val="N"/>` units
   * (0..500, percent of bar width). Default 150 (= 1.5×) in PowerPoint.
   */
  readonly gapWidthPct?: number;
  /**
   * Overlap of adjacent bars within a category in `<c:overlap val="N"/>`
   * percent (-100..100). Negative pulls bars apart, positive overlaps.
   * Defaults to 0 (clustered) / 100 (stacked).
   */
  readonly overlapPct?: number;
  /**
   * Optional legend configuration. `position` mirrors
   * `<c:legend><c:legendPos val="…"/>` — `'r'` (right) is the default,
   * `'t'`, `'b'`, `'l'`, `'tr'` (top-right) the other tokens. `null`
   * for `position` means no legend at all (the chart explicitly hides
   * it). `undefined` overall means the chart didn't author a legend
   * element — renderers should fall back to their own default.
   */
  readonly legend?: { position: 'r' | 't' | 'b' | 'l' | 'tr' | null };
  /**
   * Pie / doughnut: angle (in degrees, 0–360) at which the first slice
   * starts, measured clockwise from 12 o'clock. Mirrors
   * `<c:firstSliceAng val="…"/>`. Default 0 = start at the top.
   */
  readonly firstSliceAngleDeg?: number;
  /**
   * Doughnut hole size as a percent of the outer radius (10..90).
   * Mirrors `<c:holeSize val="…"/>`. Default 50.
   */
  readonly holeSizePct?: number;
}

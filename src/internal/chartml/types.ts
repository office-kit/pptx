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
   * Optional line stroke width in EMU (`<c:ser><c:spPr><a:ln w="…"/>`).
   * Only meaningful for line / area / scatter series. Default falls
   * back to the renderer's own pick.
   */
  readonly lineWidthEmu?: number;
  /**
   * Optional line dash style token (`<c:ser><c:spPr><a:ln><a:prstDash
   * val="…"/>`), e.g. `'dash'`, `'dot'`, `'sysDash'`. Only
   * meaningful for line / area / scatter series.
   */
  readonly lineDash?: string;
  /**
   * Optional marker style for line / scatter series (`<c:ser><c:marker>`).
   * `none` hides the markers; the others render the matching glyph at
   * each data point. Defaults to `auto`, which renderers map to a small
   * filled circle.
   */
  readonly markerSymbol?:
    | 'none'
    | 'auto'
    | 'circle'
    | 'square'
    | 'diamond'
    | 'triangle'
    | 'star'
    | 'x'
    | 'plus'
    | 'dash'
    | 'dot'
    | 'picture';
  /**
   * Marker size in points (`<c:marker><c:size val="N"/>`). PowerPoint
   * default ~5. Only meaningful when `markerSymbol` isn't `none`.
   */
  readonly markerSizePt?: number;
  /**
   * Invert the series color for negative values (bar / column charts).
   * Mirrors `<c:ser><c:invertIfNegative val="1"/>`. Renderers typically
   * paint the negative bars in the inverted shade of the series color.
   */
  readonly invertIfNegative?: boolean;
  /**
   * Optional per-data-point color overrides, indexed by point index
   * (`<c:dPt><c:idx val="N"/><c:spPr><a:solidFill>…`). Sparse — only
   * the indices that author an override appear. Pie / doughnut decks
   * almost always emit one of these per slice to break out of the
   * single-series-color default.
   */
  readonly pointColors?: ReadonlyArray<string | null>;
  /**
   * Optional per-data-point pie/doughnut slice explosion percentages
   * (`<c:dPt><c:explosion val="N"/>`). Sparse — only the indices that
   * author an explosion appear. The value is the radial offset as a
   * percentage of the slice radius (`25` ≈ a quarter-radius pull-out).
   */
  readonly pointExplosions?: ReadonlyArray<number | null>;
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
  /**
   * Per-series data-label toggle overrides. Same shape as the chart-level
   * `ChartSpec.dataLabels`; overrides win when present. Read from
   * `<c:ser><c:dLbls>`.
   */
  readonly dataLabels?: ChartDataLabels;
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
  /**
   * Number-format code from `<c:dLbls><c:numFmt formatCode="…"/>`. When
   * set, value labels are projected through this Excel-style format
   * (same subset the value axis honors: `"0%"`, `"#,##0"`, `"$#,##0"`,
   * `"0.00"`). Independent of `ChartAxisScaling.numberFormat`.
   */
  readonly numberFormat?: string;
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
  /**
   * Number-format code from `<c:numFmt formatCode="…"/>`. Common
   * values: `"0%"`, `"0.0%"`, `"#,##0"`, `"$#,##0"`, `"yyyy-mm-dd"`.
   * Renderers project a subset of Excel-style formats to label text.
   */
  readonly numberFormat?: string;
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
  /**
   * Plot-area background fill — `<c:plotArea><c:spPr><a:solidFill>
   * <a:srgbClr val="…"/>`. `null` for no fill / unsupported fill kind.
   */
  readonly plotAreaFill?: string;
  /**
   * Chart-area background fill — `<c:chartSpace><c:spPr><a:solidFill>
   * <a:srgbClr val="…"/>`. Renderers can use this as the outer card
   * color instead of the hard-coded white.
   */
  readonly chartAreaFill?: string;
  /** Optional axis title text (`<c:catAx><c:title>` / `<c:valAx><c:title>`). */
  readonly categoryAxisTitle?: string;
  readonly valueAxisTitle?: string;
  /** When `true`, value axis is hidden (`<c:valAx><c:delete val="1"/>`). */
  readonly valueAxisHidden?: boolean;
  /** When `true`, category axis is hidden (`<c:catAx><c:delete val="1"/>`). */
  readonly categoryAxisHidden?: boolean;
  /** When the value-axis emits `<c:majorGridlines/>` — its gridlines are visible. */
  readonly valueAxisMajorGridlines?: boolean;
  /** When the value-axis emits `<c:minorGridlines/>` — minor gridlines are visible. */
  readonly valueAxisMinorGridlines?: boolean;
  /**
   * Category-axis tick label skip step (`<c:catAx><c:tickLblSkip val="N"/>`):
   * render every Nth category label. Commonly 2 / 5 / 10 on dense
   * time-series charts to keep labels from overlapping.
   */
  readonly categoryAxisTickLabelSkip?: number;
  /**
   * Category-axis tick label position (`<c:catAx><c:tickLblPos val="…"/>`):
   *   - `none` hides labels but keeps the axis line
   *   - `low` / `high` puts them at the start / end (rare)
   *   - `nextTo` (default) is the standard position next to the axis
   */
  readonly categoryAxisTickLabelPos?: 'none' | 'low' | 'high' | 'nextTo';
  /**
   * Category-axis order — `'minMax'` (the data's natural order) or
   * `'maxMin'` (reversed). For bar charts PowerPoint typically emits
   * `maxMin` so the first category sits at the top instead of the
   * bottom; honour the authored value when present.
   */
  readonly categoryAxisOrientation?: 'minMax' | 'maxMin';
  /** Same for the value axis. */
  readonly valueAxisOrientation?: 'minMax' | 'maxMin';
  /** Bar / column / area grouping mode. Absent for line / pie. */
  readonly grouping?: ChartGrouping;
  /**
   * Drop lines (`<c:dropLines>`) — vertical guide lines from each data
   * point down to the value axis. Common on line / area charts.
   */
  readonly dropLines?: boolean;
  /**
   * High-low lines (`<c:hiLowLines>`) — vertical lines spanning the
   * highest and lowest series value at each category. Used by stock
   * charts and side-by-side line charts.
   */
  readonly hiLowLines?: boolean;
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
  readonly legend?: {
    position: 'r' | 't' | 'b' | 'l' | 'tr' | null;
    /** When `true`, legend overlays the plot area instead of taking a strip. */
    readonly overlay?: boolean;
  };
  /** When `true`, the chart title overlays the plot area instead of taking a strip. */
  readonly titleOverlay?: boolean;
  /**
   * `<c:dispBlanksAs val="…"/>` — how line / area renderers should
   * treat `null` values in series:
   *
   *   - `'gap'`  — leave a gap (the default)
   *   - `'zero'` — substitute zero
   *   - `'span'` — connect the surrounding points across the gap
   */
  readonly dispBlanksAs?: 'gap' | 'zero' | 'span';
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

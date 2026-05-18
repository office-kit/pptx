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
  /**
   * Optional custom label for the trendline (`<c:trendline><c:name>…`).
   * Defaults to PowerPoint's auto-generated label
   * (e.g. "Linear (X)" / "MA(5) (X)") when omitted.
   */
  readonly name?: string;
  /** Regression type — linear / exp / log / poly / power / movingAvg. */
  readonly type: 'linear' | 'exp' | 'log' | 'poly' | 'power' | 'movingAvg';
  /**
   * Forward extension (`<c:forward val="N"/>`) — the trendline runs N
   * data-point periods past the last point. Used to project future
   * values from the regression line.
   */
  readonly forward?: number;
  /**
   * Backward extension (`<c:backward val="N"/>`) — the trendline runs N
   * data-point periods before the first point.
   */
  readonly backward?: number;
  /** Optional moving-average period (only meaningful for type='movingAvg'). */
  readonly period?: number;
  /** Polynomial order (only meaningful for type='poly'). */
  readonly order?: number;
  /** Override stroke color; defaults to the series color. */
  readonly color?: string;
  /**
   * Show the regression equation next to the trendline
   * (`<c:dispEq val="1"/>`). Defaults to `false`.
   */
  readonly displayEquation?: boolean;
  /**
   * Show the R² coefficient next to the trendline
   * (`<c:dispRSqr val="1"/>`). Defaults to `false`.
   */
  readonly displayRSquared?: boolean;
}

/**
 * Per-series data-label toggles read from `<c:dLbls>` (ECMA-376
 * §21.2.2.55). All four toggles default to `false` — renderers paint
 * labels only when the corresponding flag is `true`.
 */
/**
 * `<c:dLblPos val="…"/>` — where the data label sits relative to its
 * data point. Per ECMA-376 §21.2.2.51 the token universe varies by
 * chart kind:
 *
 *   - bar / column: `'ctr' | 'inEnd' | 'outEnd' | 'inBase'`
 *   - line / area:  `'ctr' | 't' | 'b' | 'l' | 'r'`
 *   - pie / doughnut: `'ctr' | 'inEnd' | 'outEnd' | 'bestFit'`
 *
 * The union below covers every token; renderers ignore tokens that
 * don't apply to the chart's kind.
 */
export type ChartDataLabelPosition =
  | 'ctr'
  | 'inEnd'
  | 'outEnd'
  | 'inBase'
  | 't'
  | 'b'
  | 'l'
  | 'r'
  | 'bestFit';

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
  /**
   * Position of the label relative to its data point — from
   * `<c:dLbls><c:dLblPos val="…"/>`.
   */
  readonly position?: ChartDataLabelPosition;
  /**
   * Separator between concatenated label parts when more than one of
   * `showValue` / `showCategory` / `showSeriesName` / `showPercent` is
   * set. Read from `<c:dLbls><c:separator>…</c:separator>` (a
   * leaf-text-only element). Common values: `" "` (default), `", "`,
   * `"\n"`, `"; "`.
   */
  readonly separator?: string;
  /**
   * Default-run text style for the labels — projected from
   * `<c:dLbls><c:txPr>…<a:defRPr/>…</c:txPr>`. Applies to whichever
   * label parts are turned on by the `show*` toggles. Independent of
   * the chart-title style and the axis-label style.
   */
  readonly textStyle?: ChartTextStyle;
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
  /**
   * Logarithmic base for the axis from `<c:scaling><c:logBase val="N"/>`.
   * Typical values are `2`, `10`, `Math.E`. When set, the renderer
   * projects values through `Math.log(v) / Math.log(logBase)` before
   * mapping to plot coordinates. Linear when omitted.
   */
  readonly logBase?: number;
  /**
   * Authored display-units scale from `<c:dispUnits><c:builtInUnit val="…"/>`.
   * Token values map to divisors:
   *
   *   - `'hundreds'` (100), `'thousands'` (1e3), `'tenThousands'` (1e4),
   *     `'hundredThousands'` (1e5), `'millions'` (1e6),
   *     `'tenMillions'` (1e7), `'hundredMillions'` (1e8),
   *     `'billions'` (1e9), `'trillions'` (1e12)
   *
   * Renderers divide axis labels by the divisor and may append a unit
   * suffix ("K", "M", "B") for readability.
   */
  readonly displayUnits?:
    | 'hundreds'
    | 'thousands'
    | 'tenThousands'
    | 'hundredThousands'
    | 'millions'
    | 'tenMillions'
    | 'hundredMillions'
    | 'billions'
    | 'trillions';
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

/**
 * Authored text style for a chart label (title / axis title / etc.).
 * Read from the label's first `<a:rPr>` (and `<a:defRPr>` as fallback).
 * All fields are optional — absent fields mean "fall back to the
 * renderer's default for this label position."
 */
export interface ChartTextStyle {
  /** Font size in points. From `<a:rPr sz="N"/>` where N is in 100ths of a pt. */
  readonly sizePt?: number;
  /** Bold flag from `<a:rPr b="1"/>`. */
  readonly bold?: boolean;
  /** Italic flag from `<a:rPr i="1"/>`. */
  readonly italic?: boolean;
  /** Fill color as `#RRGGBB` from `<a:rPr><a:solidFill><a:srgbClr/></a:solidFill>`. */
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
  /** Optional font / color overrides for the chart title. */
  readonly titleStyle?: ChartTextStyle;
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
   * Plot-area outline color from `<c:plotArea><c:spPr><a:ln>
   * <a:solidFill><a:srgbClr/>`. `undefined` means no outline.
   */
  readonly plotAreaStrokeColor?: string;
  /**
   * Chart-area background fill — `<c:chartSpace><c:spPr><a:solidFill>
   * <a:srgbClr val="…"/>`. Renderers can use this as the outer card
   * color instead of the hard-coded white.
   */
  readonly chartAreaFill?: string;
  /**
   * Chart-area outline color from `<c:chartSpace><c:spPr><a:ln>
   * <a:solidFill><a:srgbClr/>`. Overrides the renderer's default
   * `#E5E7EB` card border.
   */
  readonly chartAreaStrokeColor?: string;
  /** Optional axis title text (`<c:catAx><c:title>` / `<c:valAx><c:title>`). */
  readonly categoryAxisTitle?: string;
  /** Authored font / color on the category-axis title (same `<a:rPr>` shape as `titleStyle`). */
  readonly categoryAxisTitleStyle?: ChartTextStyle;
  /**
   * Rotation of the category-axis title, in degrees clockwise. Maps
   * to `<c:catAx><c:title><c:tx><c:rich><a:bodyPr rot="N"/>` (the OOXML
   * value is in 60000ths of a degree, but the API surface uses plain
   * degrees). Omit to inherit the default (`0`).
   */
  readonly categoryAxisTitleRotationDeg?: number;
  /** Authored font / color on the category-axis *tick labels* — `<c:catAx><c:txPr>`. */
  readonly categoryAxisLabelStyle?: ChartTextStyle;
  /**
   * Authored rotation on the category-axis tick labels, in degrees. From
   * `<c:catAx><c:txPr><a:bodyPr rot="N"/>` where N is in 60000ths of a
   * degree. Positive values rotate clockwise (PowerPoint convention),
   * matching the SVG `transform=rotate()` sense.
   */
  readonly categoryAxisLabelRotationDeg?: number;
  readonly valueAxisTitle?: string;
  /** Authored font / color on the value-axis title. */
  readonly valueAxisTitleStyle?: ChartTextStyle;
  /**
   * Rotation of the value-axis title, in degrees clockwise. PowerPoint
   * often emits `-90` (or `vert270`) so the title reads bottom-to-top
   * alongside the axis. Maps to `<c:valAx><c:title><c:tx><c:rich>
   * <a:bodyPr rot="N"/>` (60000ths of a degree on the wire).
   */
  readonly valueAxisTitleRotationDeg?: number;
  /** Authored font / color on the value-axis *tick labels* — `<c:valAx><c:txPr>`. */
  readonly valueAxisLabelStyle?: ChartTextStyle;
  /**
   * Authored rotation on the value-axis tick labels, in degrees. From
   * `<c:valAx><c:txPr><a:bodyPr rot="N"/>` (N in 60000ths of a degree).
   * Same sense as `categoryAxisLabelRotationDeg`.
   */
  readonly valueAxisLabelRotationDeg?: number;
  /** When `true`, value axis is hidden (`<c:valAx><c:delete val="1"/>`). */
  readonly valueAxisHidden?: boolean;
  /** When `true`, category axis is hidden (`<c:catAx><c:delete val="1"/>`). */
  readonly categoryAxisHidden?: boolean;
  /** When the value-axis emits `<c:majorGridlines/>` — its gridlines are visible. */
  readonly valueAxisMajorGridlines?: boolean;
  /**
   * Major-tick mark mode on the value axis (`<c:valAx><c:majorTickMark val="…"/>`):
   *
   *   - `'out'` — outside the plot edge (default)
   *   - `'in'` — inside the plot
   *   - `'cross'` — across the axis line
   *   - `'none'` — no tick marks
   */
  readonly valueAxisMajorTickMark?: 'in' | 'out' | 'cross' | 'none';
  /** Major-tick mark mode on the category axis (`<c:catAx><c:majorTickMark>`). */
  readonly categoryAxisMajorTickMark?: 'in' | 'out' | 'cross' | 'none';
  /** Minor-tick mark mode on the value axis (`<c:valAx><c:minorTickMark>`). */
  readonly valueAxisMinorTickMark?: 'in' | 'out' | 'cross' | 'none';
  /** Minor-tick mark mode on the category axis (`<c:catAx><c:minorTickMark>`). */
  readonly categoryAxisMinorTickMark?: 'in' | 'out' | 'cross' | 'none';
  /**
   * Authored color on the value-axis major gridlines — `<c:valAx>
   * <c:majorGridlines><c:spPr><a:ln><a:solidFill><a:srgbClr val="…"/>`.
   * Returned as `#RRGGBB`. `undefined` falls back to the renderer's
   * default (a light gray).
   */
  readonly valueAxisMajorGridlineColor?: string;
  /** Companion authored color on the value-axis minor gridlines. */
  readonly valueAxisMinorGridlineColor?: string;
  /** Authored color on the category-axis major gridlines (same shape as `valueAxisMajorGridlineColor`). */
  readonly categoryAxisMajorGridlineColor?: string;
  /** Authored color on the category-axis minor gridlines. */
  readonly categoryAxisMinorGridlineColor?: string;
  /**
   * Authored color on the value-axis line itself — `<c:valAx><c:spPr>
   * <a:ln><a:solidFill><a:srgbClr val="…"/>`. Returned as `#RRGGBB`.
   * `undefined` falls back to the renderer's default axis stroke.
   */
  readonly valueAxisLineColor?: string;
  /** Authored color on the category-axis line (same shape as `valueAxisLineColor`). */
  readonly categoryAxisLineColor?: string;
  /** When the value-axis emits `<c:minorGridlines/>` — minor gridlines are visible. */
  readonly valueAxisMinorGridlines?: boolean;
  /**
   * Whether the category axis emits `<c:majorGridlines/>`. Sounds odd
   * for column charts (where category gridlines are vertical, between
   * categories) but bar charts put the cat axis on the vertical edge
   * and use these as horizontal guide lines per category band.
   */
  readonly categoryAxisMajorGridlines?: boolean;
  /** Companion `<c:minorGridlines/>` on the category axis. */
  readonly categoryAxisMinorGridlines?: boolean;
  /**
   * Category-axis tick label skip step (`<c:catAx><c:tickLblSkip val="N"/>`):
   * render every Nth category label. Commonly 2 / 5 / 10 on dense
   * time-series charts to keep labels from overlapping.
   */
  readonly categoryAxisTickLabelSkip?: number;
  /**
   * Category-axis tick *mark* skip (`<c:catAx><c:tickMarkSkip val="N"/>`):
   * draw every Nth tick mark independently of the label-skip stride.
   * Useful when you want fewer label collisions but the same dense
   * tick lattice; defaults to 1 (every tick).
   */
  readonly categoryAxisTickMarkSkip?: number;
  /**
   * Category-axis tick label position (`<c:catAx><c:tickLblPos val="…"/>`):
   *   - `none` hides labels but keeps the axis line
   *   - `low` / `high` puts them at the start / end (rare)
   *   - `nextTo` (default) is the standard position next to the axis
   */
  readonly categoryAxisTickLabelPos?: 'none' | 'low' | 'high' | 'nextTo';
  /**
   * Distance from the axis line to the labels, expressed as a percent of
   * the chart text size — `<c:catAx><c:lblOffset val="N"/>` where N is
   * 0..1000 (default 100). Larger values push category labels further
   * from the axis. Per ECMA-376 §21.2.2.94.
   */
  readonly categoryAxisLabelOffset?: number;
  /**
   * Multi-line category-label alignment relative to the tick mark —
   * `<c:catAx><c:lblAlgn val="ctr|l|r"/>`. PowerPoint defaults to
   * `ctr` when omitted; the authored value wins.
   */
  readonly categoryAxisLabelAlign?: 'ctr' | 'l' | 'r';
  /**
   * Number-format code for the category-axis tick labels —
   * `<c:catAx><c:numFmt formatCode="…"/>`. Most useful on date-style
   * categories (`"mm/dd/yyyy"`, `"mmm"`, etc.) but accepts any Excel
   * format string. Independent of `valueAxis.numberFormat` (which
   * targets the value axis).
   */
  readonly categoryAxisNumberFormat?: string;
  /**
   * Category-axis order — `'minMax'` (the data's natural order) or
   * `'maxMin'` (reversed). For bar charts PowerPoint typically emits
   * `maxMin` so the first category sits at the top instead of the
   * bottom; honour the authored value when present.
   */
  readonly categoryAxisOrientation?: 'minMax' | 'maxMin';
  /** Same for the value axis. */
  readonly valueAxisOrientation?: 'minMax' | 'maxMin';
  /**
   * Where the category axis crosses the value axis. Either an enum
   * keyword (`<c:valAx><c:crosses val="autoZero|min|max"/>`) or a
   * specific numeric value (`<c:valAx><c:crossesAt val="N"/>`). The two
   * forms are mutually exclusive — PowerPoint emits one or the other.
   * Default is `autoZero` (the category axis sits at value 0 if the
   * range straddles zero, otherwise at the closer extreme).
   */
  readonly valueAxisCrosses?: 'autoZero' | 'min' | 'max' | { at: number };
  /**
   * Whether the value axis crosses the category axis *between* tick
   * marks (the default for bar / column / area) or *at* each tick mark
   * (the default for line / scatter). Maps to `<c:valAx>
   * <c:crossBetween val="between|midCat"/>`. PowerPoint emits this when
   * the chart kind makes the default value non-obvious — surface it
   * here so the round-trip preserves the authored intent.
   */
  readonly valueAxisCrossBetween?: 'between' | 'midCat';
  /**
   * When `false`, plot data from hidden cells in the embedded workbook
   * — maps to `<c:plotVisOnly val="0"/>`. PowerPoint's default is
   * `true` (only plot visible cells), so omitting this field emits
   * `val="1"` to stay round-trip-safe with PowerPoint-authored files.
   */
  readonly plotVisibleCellsOnly?: boolean;
  /**
   * Renders the chart area with rounded corners
   * (`<c:chartSpace><c:roundedCorners val="1"/>`). PowerPoint's default
   * is `false`; surface only when explicitly `true` so the round-trip
   * doesn't add a redundant `false`.
   */
  readonly roundedCorners?: boolean;
  /**
   * PowerPoint built-in chart-style preset
   * (`<c:chartSpace><c:style val="N"/>`), 1–48. Encodes a curated combo
   * of theme accent colors, gradients, effects, and font sizes that
   * PowerPoint applies when the user picks a chart style from the
   * "Chart Styles" gallery. Surface for round-trip parity; renderers in
   * pptx-kit don't (yet) interpret it.
   */
  readonly chartStyle?: number;
  /**
   * Language code for the chart's number / date formatters
   * (`<c:chartSpace><c:lang val="…"/>`). PowerPoint emits the user's
   * Office UI language (e.g. `'en-US'`, `'ja-JP'`). Carried for
   * round-trip parity; renderers in pptx-kit don't act on it yet.
   */
  readonly language?: string;
  /**
   * `<c:chartSpace><c:date1904 val="…"/>` — the Excel date-system flag.
   * `false` (the default) uses the 1900-epoch; `true` uses 1904.
   * Surface for parity; renderers don't act on it yet.
   */
  readonly date1904?: boolean;
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
    /**
     * Authored font / color on the legend's text. From `<c:legend><c:txPr>`'s
     * first `<a:p><a:pPr><a:defRPr>` (or `<a:r><a:rPr>` as fallback). Same
     * shape as `titleStyle`.
     */
    readonly textStyle?: ChartTextStyle;
    /**
     * Series indices that the legend hides via `<c:legend><c:legendEntry>
     * <c:idx val="N"/><c:delete val="1"/></c:legendEntry>`. Renderers
     * filter them from the legend list while still plotting them.
     * Common use: keep a trendline series in the data but drop its
     * legend entry.
     */
    readonly hiddenIndices?: ReadonlyArray<number>;
  };
  /** When `true`, the chart title overlays the plot area instead of taking a strip. */
  readonly titleOverlay?: boolean;
  /**
   * `<c:barChart><c:varyColors val="1"/>` etc. — when `true` and the
   * chart has a single series, each data point gets a distinct color
   * from the palette. Pie / doughnut already vary colors implicitly;
   * this flag is most useful for single-series bar / column.
   */
  readonly varyColors?: boolean;
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

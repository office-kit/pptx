// Chart authoring types.
//
// One or more named series of numeric values plotted against shared
// string categories (bar / column / line / pie / doughnut / area / radar /
// stock / surface), plus the xy(z)-tuple kinds (scatter / bubble). Every
// kind is both authorable and readable; together with the modifiers on
// `ChartSpec` (`view3D`, `ofPie`, `surfaceContour`) they cover all sixteen
// plot-group elements of CT_PlotArea (ECMA-376 §21.2.2.145).

/**
 * Chart type tokens. The kind names the data shape; the 3-D and
 * pie-of-pie variants are modifiers on `ChartSpec` rather than kinds of
 * their own:
 *
 *   | kind       | element              | modifier → element                        |
 *   | ---------- | -------------------- | ----------------------------------------- |
 *   | `bar`      | `<c:barChart>`       | `view3D` → `<c:bar3DChart>`               |
 *   | `column`   | `<c:barChart>`       | `view3D` → `<c:bar3DChart>`               |
 *   | `line`     | `<c:lineChart>`      | `view3D` → `<c:line3DChart>`              |
 *   | `area`     | `<c:areaChart>`      | `view3D` → `<c:area3DChart>`              |
 *   | `pie`      | `<c:pieChart>`       | `view3D` → `<c:pie3DChart>`, `ofPie` → `<c:ofPieChart>` |
 *   | `doughnut` | `<c:doughnutChart>`  |                                           |
 *   | `scatter`  | `<c:scatterChart>`   |                                           |
 *   | `bubble`   | `<c:bubbleChart>`    | `bubble3D` shades the bubbles             |
 *   | `radar`    | `<c:radarChart>`     |                                           |
 *   | `stock`    | `<c:stockChart>`     |                                           |
 *   | `surface`  | `<c:surface3DChart>` | `surfaceContour` → `<c:surfaceChart>`     |
 */
export type ChartKind =
  | 'bar'
  | 'column'
  | 'line'
  | 'pie'
  | 'doughnut'
  | 'area'
  | 'scatter'
  | 'radar'
  | 'bubble'
  | 'stock'
  | 'surface';

/** One labelled series of numeric values. */
export interface ChartSeries {
  /** Series label rendered in the legend. */
  readonly name: string;
  /**
   * Numeric values, one per category. `null` slots become empty cells in
   * the embedded workbook (PowerPoint draws them as a gap). Lengths
   * shorter than the category count are right-padded with `null`.
   *
   * For `scatter` / `bubble` kinds there are no categories: `values`
   * holds the series' y-channel (`<c:yVal>`), paired positionally with
   * `xValues`. For `radar` it behaves like `line` (values per category).
   *
   * A `stock` chart reads its series by position, not by name: three
   * series are high, low, close; four are open, high, low, close
   * (CT_StockChart allows exactly 3 or 4). A `surface` chart is a grid:
   * each series is one row of depth, each category one column.
   */
  readonly values: ReadonlyArray<number | null>;
  /**
   * X-channel values for `scatter` / `bubble` series (`<c:xVal>`), paired
   * positionally with `values`. Absent for every other kind, where
   * `values` is plotted against the shared `categories`.
   */
  readonly xValues?: ReadonlyArray<number | null>;
  /**
   * Per-point bubble sizes for `bubble` series (`<c:bubbleSize>`), paired
   * positionally with `xValues` / `values`. The renderer scales each
   * bubble's area (or width, per `ChartSpec.bubbleSizeRepresents`) to
   * this magnitude. Absent for non-bubble kinds.
   */
  readonly bubbleSizes?: ReadonlyArray<number | null>;
  /** Optional `#RRGGBB` fill override. Defaults to the theme's accent palette. */
  readonly color?: string;
  /**
   * Per-series chart-kind override for combo charts (e.g. a `line`
   * series overlaid on a `column` chart). Series sharing an effective
   * kind are grouped into one plot group (`<c:barChart>` /
   * `<c:lineChart>` / `<c:areaChart>`); the groups share the category
   * axis. Only the category kinds can mix — `bar` / `column` / `line` /
   * `area`. Absent = plotted with the chart's own `kind`.
   */
  readonly chartKind?: 'bar' | 'column' | 'line' | 'area';
  /**
   * Plot this series against the secondary value axis — the right-hand
   * axis PowerPoint shows for combo charts with mixed units
   * (`<c:valAx>` pair with `axPos="r"` plus a deleted companion
   * `<c:catAx>`). The builder emits the secondary axis pair on demand.
   * Only meaningful for the category kinds; rejected for pie / doughnut.
   * At least one series must stay on the primary axis — a chart whose
   * every series is secondary is rejected at authoring time.
   */
  readonly secondaryAxis?: boolean;
  /**
   * Optional line stroke width in EMU (`<c:ser><c:spPr><a:ln w="…"/>`).
   * Only meaningful for line / area / scatter series. Default falls
   * back to the renderer's own pick.
   */
  readonly lineWidthEmu?: number;
  /**
   * Optional stroke color of the series outline / line
   * (`<c:ser><c:spPr><a:ln><a:solidFill>`), as an sRGB hex. Defaults to
   * `color`; on a pie / doughnut series it is the border between slices.
   */
  readonly lineColor?: string;
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
   * Optional marker fill + outline color (`<c:marker><c:spPr>`), as an sRGB
   * hex. Defaults to `color`. The builder always writes it for line /
   * scatter / radar series, because PowerPoint paints a marker without
   * `<c:spPr>` in the theme's automatic color instead of the series color.
   */
  readonly markerColor?: string;
  /**
   * Optional marker outline color (`<c:marker><c:spPr><a:ln>`), as an sRGB
   * hex. Defaults to `markerColor`, then `color`.
   */
  readonly markerLineColor?: string;
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
   * Optional per-data-point data-label overrides
   * (`<c:dLbls><c:dLbl><c:idx val="N"/>…`). Sparse — `null` slots fall back
   * to the series-level `dataLabels`. PowerPoint draws the per-point element
   * over the series defaults, which is how pie / doughnut exporters give each
   * slice its own label content and font.
   */
  readonly pointDataLabels?: ReadonlyArray<ChartDataLabels | null>;
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
  /**
   * Opacity of the series fill, `0` (transparent) to `1` (opaque) —
   * `<a:srgbClr><a:alpha val="…"/>` on the series' `<a:solidFill>`. Applies
   * to `color` and to every `pointColors` override. Overlapping area,
   * radar and bubble series are the usual reason to set it.
   */
  readonly fillOpacity?: number;
  /**
   * Error bars along the value direction (`<c:errBars>`; `<c:errDir
   * val="y"/>` on scatter / bubble). Valid on bar / column / line / area /
   * scatter / bubble series; rejected elsewhere, where CT_*Ser has no
   * `errBars` child.
   */
  readonly errorBars?: ChartErrorBars;
  /**
   * Error bars along the x direction (`<c:errDir val="x"/>`). Only scatter
   * and bubble series have an x channel to attach them to.
   */
  readonly xErrorBars?: ChartErrorBars;
}

/**
 * How far an error bar reaches (`<c:errValType>` plus its operand).
 * `stdErr` takes no operand; `cust` carries per-point literal amounts.
 */
export type ChartErrorBarAmount =
  | { readonly type: 'fixedVal' | 'percentage' | 'stdDev'; readonly value: number }
  | { readonly type: 'stdErr' }
  | {
      readonly type: 'cust';
      /** Per-point positive amounts (`<c:plus><c:numLit>`). */
      readonly plus?: ReadonlyArray<number | null>;
      /** Per-point negative amounts (`<c:minus><c:numLit>`). */
      readonly minus?: ReadonlyArray<number | null>;
    };

/** Error bars of one series in one direction (ECMA-376 §21.2.2.55). */
export interface ChartErrorBars {
  /** Which side of the point gets a bar (`<c:errBarType>`). */
  readonly barType: 'both' | 'plus' | 'minus';
  readonly amount: ChartErrorBarAmount;
  /** Drop the T-shaped end caps (`<c:noEndCap val="1"/>`). */
  readonly noEndCap?: boolean;
  /** Bar stroke color as `#RRGGBB`. */
  readonly color?: string;
  /** Bar stroke width in EMU. */
  readonly lineWidthEmu?: number;
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
  // `ChartSpec` is bidirectional (authored into `addSlideChart`, read back from
  // `getShapeChartSpec`), and the reader always emits all four toggles, so they
  // are required: a sparse `{ showValue: true }` would widen what every reader
  // gets. Set the ones you want on and the rest to `false`.
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
  /**
   * Leader lines between a label and its data point
   * (`<c:dLbls><c:showLeaderLines val="…"/>`). Omitted = the
   * application default (shown). Series- and chart-level only: a
   * per-point `<c:dLbl>` has no such element, so `pointDataLabels`
   * ignores it on write and never reads it.
   */
  readonly showLeaderLines?: boolean;
  /** Bubble size of each data point (`<c:showBubbleSize>`); bubble charts only. */
  readonly showBubbleSize?: boolean;
  /** Legend key swatch next to each label (`<c:showLegendKey>`). */
  readonly showLegendKey?: boolean;
  /** Label background fill as `#RRGGBB` (`<c:spPr><a:solidFill>`). */
  readonly fillColor?: string;
  /**
   * Literal label text replacing the generated one (`<c:dLbl><c:tx>
   * <c:rich>`) — how a scatter point gets a name. Per-point only: the
   * series- and chart-level `<c:dLbls>` have no `<c:tx>`, so it is ignored
   * there on write and never read.
   */
  readonly text?: string;
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
  /**
   * Show the display-units caption next to the axis ("Thousands",
   * "Millions", …) — an empty `<c:dispUnitsLbl/>` under `<c:dispUnits>`.
   * Only meaningful together with `displayUnits`.
   */
  readonly displayUnitsLabel?: boolean;
}

/**
 * 3-D view of a chart (`<c:view3D>`, ECMA-376 §21.2.2.228). Every field is
 * optional; an empty object still selects the 3-D plot-group element and
 * leaves the camera to the application's defaults.
 */
export interface ChartView3D {
  /** Elevation in degrees, -90..90 (`<c:rotX>`). */
  readonly rotX?: number;
  /** Rotation about the vertical axis in degrees, 0..360 (`<c:rotY>`). */
  readonly rotY?: number;
  /**
   * Right-angle axes (`<c:rAngAx>`): an oblique projection that ignores
   * `perspective`. PowerPoint's default for bar / column / line / area.
   */
  readonly rightAngleAxes?: boolean;
  /** Field of view, 0..240 (`<c:perspective>`); PowerPoint's default is 30. */
  readonly perspective?: number;
  /** Depth as a percent of the chart width, 20..2000 (`<c:depthPercent>`). */
  readonly depthPercent?: number;
  /** Height as a percent of the chart width, 5..500 (`<c:hPercent>`). */
  readonly heightPercent?: number;
}

/**
 * Pie-of-pie / bar-of-pie settings (`<c:ofPieChart>`, ECMA-376
 * §21.2.2.126): the trailing points of a pie move into a second plot.
 */
export interface ChartOfPie {
  /** Shape of the second plot (`<c:ofPieType>`). */
  readonly type: 'pie' | 'bar';
  /**
   * What decides which points move (`<c:splitType>`): `pos` — the last
   * `splitPos` points; `val` — points below `splitPos`; `percent` — points
   * below `splitPos` percent of the total; `cust` — the `customSplit`
   * indices; `auto` — the application's choice.
   */
  readonly splitType?: 'auto' | 'cust' | 'percent' | 'pos' | 'val';
  /** Threshold for `splitType` `pos` / `val` / `percent` (`<c:splitPos>`). */
  readonly splitPos?: number;
  /** Point indices in the second plot for `splitType: 'cust'` (`<c:custSplit>`). */
  readonly customSplit?: ReadonlyArray<number>;
  /** Second plot size as a percent of the first, 5..200 (`<c:secondPieSize>`). */
  readonly secondPieSizePct?: number;
  /** Gap between the two plots, 0..500 percent (`<c:gapWidth>`). */
  readonly gapWidthPct?: number;
  /** Connector lines between the two plots (`<c:serLines>`). */
  readonly seriesLines?: boolean;
}

/** Up / down bars between the first and last line series (`<c:upDownBars>`). */
export interface ChartUpDownBars {
  /** Bar width gap, 0..500 percent (`<c:gapWidth>`); PowerPoint's default is 150. */
  readonly gapWidthPct?: number;
  /** Fill of the rising bars as `#RRGGBB` (`<c:upBars>`). */
  readonly upColor?: string;
  /** Fill of the falling bars as `#RRGGBB` (`<c:downBars>`). */
  readonly downColor?: string;
}

/** Data table under the plot area (`<c:dTable>`, ECMA-376 §21.2.2.54). */
export interface ChartDataTable {
  readonly showHorizontalBorder?: boolean;
  readonly showVerticalBorder?: boolean;
  readonly showOutline?: boolean;
  /** Legend keys next to the series names. */
  readonly showKeys?: boolean;
  readonly textStyle?: ChartTextStyle;
}

/** Calendar unit of a date axis (ST_TimeUnit). */
export type ChartTimeUnit = 'days' | 'months' | 'years';

/**
 * Date-axis settings. Its presence turns the category axis into a
 * `<c:dateAx>`: `ChartSpec.categories` must then hold date serial numbers
 * (days since the workbook epoch, see `ChartSpec.date1904`) as strings,
 * and they are written as a numeric channel so the axis can space points
 * by date. Pair it with `categoryAxisNumberFormat` (e.g. `"yyyy-mm"`).
 */
export interface ChartDateAxis {
  readonly baseTimeUnit?: ChartTimeUnit;
  readonly majorUnit?: number;
  readonly majorTimeUnit?: ChartTimeUnit;
  readonly minorUnit?: number;
  readonly minorTimeUnit?: ChartTimeUnit;
}

/**
 * Series (depth) axis of a 3-D chart (`<c:serAx>`). The builder emits the
 * axis whenever the plot-group element requires it; this object only
 * formats it.
 */
export interface ChartSeriesAxis {
  readonly hidden?: boolean;
  readonly title?: string;
  readonly titleStyle?: ChartTextStyle;
  readonly labelStyle?: ChartTextStyle;
  readonly majorGridlines?: boolean;
  readonly orientation?: 'minMax' | 'maxMin';
  readonly tickLabelPos?: 'none' | 'low' | 'high' | 'nextTo';
  readonly tickLabelSkip?: number;
  readonly lineColor?: string;
}

/**
 * Manual position of a chart element, as fractions (0..1) of the chart
 * space: `x` / `y` measured from its top-left corner, `w` / `h` as a
 * share of its size (`<c:layout><c:manualLayout>` with `xMode` / `yMode`
 * set to `edge`).
 */
export interface ChartManualLayout {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
  /**
   * Plot area only: whether the box bounds the plot alone (`inner`) or the
   * plot plus its tick labels and axis titles (`outer`, the default).
   */
  readonly target?: 'inner' | 'outer';
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
  /**
   * Font face applied to both the latin and east-asian typeface slots —
   * `<a:rPr><a:latin typeface="…"/><a:ea typeface="…"/>`. East-asian is
   * set alongside latin so a Japanese / CJK family (e.g. `'Yu Gothic'`)
   * renders the labels instead of the renderer's latin-only fallback.
   * Mirrors the latin/ea pairing PowerPoint emits for CJK chart fonts.
   */
  readonly font?: string;
  /**
   * Complex-script font face — `<a:rPr><a:cs typeface="…"/>`. Written only
   * when set: `font` fills the latin and east-asian slots and never the
   * complex-script one, so a chart that mixes Arabic / Hebrew / Thai labels
   * has to name the face here as well.
   */
  readonly fontComplexScript?: string;
  /** Font size in points. From `<a:rPr sz="N"/>` where N is in 100ths of a pt. */
  readonly sizePt?: number;
  /** Bold flag from `<a:rPr b="1"/>`. */
  readonly bold?: boolean;
  /** Italic flag from `<a:rPr i="1"/>`. */
  readonly italic?: boolean;
  /** Fill color as `#RRGGBB` from `<a:rPr><a:solidFill><a:srgbClr/></a:solidFill>`. */
  readonly color?: string;
}

/**
 * Formatting for the secondary value axis (`<c:valAx>` with `axPos="r"`)
 * that `ChartSeries.secondaryAxis` series plot against. The fields mirror
 * the primary `valueAxis*` fields of `ChartSpec` without the prefix, limited
 * to what a right-hand axis needs; the axis position (`r`) and crossing
 * (`max`) stay fixed. Omit the whole object for PowerPoint's defaults.
 */
export interface ChartSecondaryValueAxis {
  readonly scaling?: ChartAxisScaling;
  readonly title?: string;
  readonly titleStyle?: ChartTextStyle;
  readonly labelStyle?: ChartTextStyle;
  readonly majorGridlines?: boolean;
  readonly majorGridlineColor?: string;
  /** Major gridline width in EMU, the mirror of `ChartSpec.valueAxisMajorGridlineWidthEmu`. */
  readonly majorGridlineWidthEmu?: number;
  readonly lineColor?: string;
  /** Axis line width in EMU, the mirror of `ChartSpec.valueAxisLineWidthEmu`. */
  readonly lineWidthEmu?: number;
  readonly majorTickMark?: 'in' | 'out' | 'cross' | 'none';
  readonly minorTickMark?: 'in' | 'out' | 'cross' | 'none';
  /** Tick-label position, the mirror of `ChartSpec.valueAxisTickLabelPos`. */
  readonly tickLabelPos?: 'none' | 'low' | 'high' | 'nextTo';
  readonly crossBetween?: 'between' | 'midCat';
}

/**
 * Every field a chart can carry, with no cross-field rules applied — the
 * shape `readChartSpec` returns. A deck authored elsewhere can combine
 * fields that this library would never write together, so the read type
 * stays permissive; `ChartSpec` is the closed, write-side domain.
 */
export interface ReadChartSpec {
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
   * Formatting for the secondary value axis. Only meaningful when at least
   * one series sets `secondaryAxis: true`; ignored otherwise.
   */
  readonly secondaryValueAxis?: ChartSecondaryValueAxis;
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
  /**
   * Value-axis tick label position (`<c:valAx><c:tickLblPos val="…"/>`),
   * the same choices as `categoryAxisTickLabelPos`: `none` hides the
   * labels, `low` / `high` puts them at the plot's start / end, `nextTo`
   * (default) next to the axis. Primary axis only.
   */
  readonly valueAxisTickLabelPos?: 'none' | 'low' | 'high' | 'nextTo';
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
   * Authored width of the value-axis major gridlines in EMU — the `w` of
   * the same `<a:ln>` as `valueAxisMajorGridlineColor` (12700 = 1 pt).
   * `undefined` leaves the width to the application (PowerPoint: 0.75 pt).
   */
  readonly valueAxisMajorGridlineWidthEmu?: number;
  /** Companion width of the value-axis minor gridlines. */
  readonly valueAxisMinorGridlineWidthEmu?: number;
  /** Width of the category-axis major gridlines (same shape as `valueAxisMajorGridlineWidthEmu`). */
  readonly categoryAxisMajorGridlineWidthEmu?: number;
  /** Width of the category-axis minor gridlines. */
  readonly categoryAxisMinorGridlineWidthEmu?: number;
  /**
   * Authored color on the value-axis line itself — `<c:valAx><c:spPr>
   * <a:ln><a:solidFill><a:srgbClr val="…"/>`. Returned as `#RRGGBB`.
   * `undefined` falls back to the renderer's default axis stroke.
   */
  readonly valueAxisLineColor?: string;
  /** Authored color on the category-axis line (same shape as `valueAxisLineColor`). */
  readonly categoryAxisLineColor?: string;
  /**
   * Authored width of the value-axis line in EMU — the `w` of the same
   * `<a:ln>` as `valueAxisLineColor`. `undefined` leaves it to the application.
   */
  readonly valueAxisLineWidthEmu?: number;
  /** Width of the category-axis line (same shape as `valueAxisLineWidthEmu`). */
  readonly categoryAxisLineWidthEmu?: number;
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
   * Toggle multi-level (hierarchical) category labels —
   * `<c:catAx><c:noMultiLvlLbl val="0|1"/>`. PowerPoint defaults to
   * `0` (multi-level labels stack). Set to `true` to flatten
   * hierarchical categories into a single row.
   */
  readonly categoryAxisNoMultiLevelLabel?: boolean;
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
   * @office-kit/pptx don't (yet) interpret it.
   */
  readonly chartStyle?: number;
  /**
   * Language code for the chart's number / date formatters
   * (`<c:chartSpace><c:lang val="…"/>`). PowerPoint emits the user's
   * Office UI language (e.g. `'en-US'`, `'ja-JP'`). Carried for
   * round-trip parity; renderers in @office-kit/pptx don't act on it yet.
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
   * Whether a line chart draws point markers (`<c:lineChart><c:marker
   * val="1"/>`). `true` is PowerPoint's "Line with Markers" subtype;
   * absent / `false` is the plain "Line" subtype (no markers). Only
   * meaningful for `kind: 'line'`.
   */
  readonly lineMarkers?: boolean;
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
    /** Manual position of the legend box; `target` does not apply. */
    readonly layout?: ChartManualLayout;
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
  /**
   * Scatter sub-type from `<c:scatterChart><c:scatterStyle val="…"/>`
   * (ECMA-376 ST_ScatterStyle). Governs whether the renderer connects
   * points with straight / smoothed lines and whether it draws markers:
   *
   *   - `'marker'` (and the renderer's default when absent) — markers only
   *   - `'line'` / `'smooth'` — connecting line only
   *   - `'lineMarker'` / `'smoothMarker'` — line + markers
   *   - `'none'` — neither
   */
  readonly scatterStyle?: 'none' | 'line' | 'lineMarker' | 'marker' | 'smooth' | 'smoothMarker';
  /**
   * Radar sub-type from `<c:radarChart><c:radarStyle val="…"/>`
   * (ECMA-376 ST_RadarStyle): `'standard'` (polyline only), `'marker'`
   * (polyline + point markers), or `'filled'` (the polygon is filled at
   * reduced opacity).
   */
  readonly radarStyle?: 'standard' | 'marker' | 'filled';
  /**
   * Bubble-size scale percentage from `<c:bubbleChart><c:bubbleScale
   * val="N"/>` (0..300, PowerPoint default 100). Scales every bubble's
   * rendered radius proportionally.
   */
  readonly bubbleScale?: number;
  /**
   * Whether a bubble's `<c:bubbleSize>` maps to the bubble's area or its
   * width — `<c:bubbleChart><c:sizeRepresents val="area|w"/>` (ECMA-376
   * ST_SizeRepresents; the `'w'` token is surfaced as `'width'`).
   * PowerPoint's default is `'area'`, so radius scales with `sqrt(size)`;
   * `'width'` scales radius linearly with size.
   */
  readonly bubbleSizeRepresents?: 'area' | 'width';
  /**
   * Shade bubbles as spheres — `<c:bubble3D val="1"/>` on each series of the
   * bubble chart (PowerPoint's "3-D Bubble" subtype). Unlike the
   * other 3-D variants it has no camera, so `view3D` does not apply.
   */
  readonly bubble3D?: boolean;
  /** Plot bubbles with a negative size (`<c:showNegBubbles>`). */
  readonly showNegativeBubbles?: boolean;
  /**
   * 3-D view. Its presence selects the 3-D plot-group element for `bar` /
   * `column` / `line` / `area` / `pie`; for `surface` it only positions the
   * camera. Rejected for every other kind, and for combo charts (a 3-D
   * group cannot share a plot area with another group in PowerPoint).
   */
  readonly view3D?: ChartView3D;
  /** Shape of 3-D bars / columns (`<c:bar3DChart><c:shape>`). Requires `view3D`. */
  readonly bar3DShape?: 'box' | 'cone' | 'coneToMax' | 'cylinder' | 'pyramid' | 'pyramidToMax';
  /**
   * Depth gap between the series rows of a 3-D bar / column / line / area
   * chart, 0..500 percent (`<c:gapDepth>`). Requires `view3D`.
   */
  readonly gapDepthPct?: number;
  /** Series (depth) axis formatting for 3-D charts and `surface`. */
  readonly seriesAxis?: ChartSeriesAxis;
  /** Turns a `pie` into a pie-of-pie / bar-of-pie chart. */
  readonly ofPie?: ChartOfPie;
  /**
   * Surface charts: `true` draws the top-down contour plot
   * (`<c:surfaceChart>`), absent / `false` the 3-D surface
   * (`<c:surface3DChart>`).
   */
  readonly surfaceContour?: boolean;
  /** Surface charts: draw the mesh only, without color bands (`<c:wireframe>`). */
  readonly surfaceWireframe?: boolean;
  /**
   * Up / down bars. Valid on `line` charts (between the first and last
   * series) and on `stock` charts with an open series, where they are the
   * candlestick bodies.
   */
  readonly upDownBars?: ChartUpDownBars;
  /** Data table under the plot area. Rejected for the axis-less kinds. */
  readonly dataTable?: ChartDataTable;
  /** Turns the category axis into a date axis — see `ChartDateAxis`. */
  readonly categoryAxisDate?: ChartDateAxis;
  /**
   * Scaling of the horizontal axis where it is numeric: the x value axis
   * of `scatter` / `bubble` charts (every field applies) and a date axis
   * (`min` / `max` only, as date serials). Rejected on a text category
   * axis, which has no scale to set.
   */
  readonly categoryAxisScaling?: ChartAxisScaling;
  /**
   * Outer levels of a multi-level category axis (`<c:multiLvlStrRef>`),
   * ordered from the level next to `categories` outward. Each level has
   * one entry per category; an empty string continues the previous
   * group, so `['2024', '', '', '', '2025', '', '', '']` groups eight
   * quarters into two years.
   */
  readonly categoryGroupLevels?: ReadonlyArray<ReadonlyArray<string>>;
  /** Manual position of the plot area. */
  readonly plotAreaLayout?: ChartManualLayout;
  /**
   * Manual position of the chart title's top-left corner, as fractions of
   * the chart space. A title sizes itself to its text, so there is no
   * width / height to set.
   */
  readonly titleLayout?: { readonly x: number; readonly y: number };
  /**
   * Hide the value-axis line while keeping its tick labels
   * (`<c:valAx><c:spPr><a:ln><a:noFill/>`). `valueAxisHidden` removes the
   * labels too; this does not.
   */
  readonly valueAxisLineHidden?: boolean;
  /** Hide the category-axis line while keeping its labels. */
  readonly categoryAxisLineHidden?: boolean;
}

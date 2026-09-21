// The closed, write-side chart domain. `ReadChartSpec` names every field a
// chart can carry; this file says which of them each kind actually draws, so
// that a field the kind has no element for is a compile error rather than a
// value the builder silently drops.
//
// The field sets are runtime arrays and the types are derived from them, so
// the type and the `isChartSpec` guard cannot drift apart.

import type {
  ChartAxisScaling,
  ChartDateAxis,
  ChartOfPie,
  ChartSeries,
  ChartView3D,
  ReadChartSpec,
} from './types.ts';

/** Field of a chart whose availability depends on the kind. */
type SpecField = Exclude<keyof ReadChartSpec, 'kind' | 'series'>;

type FieldList = ReadonlyArray<SpecField>;

/**
 * A kind's field set: `K` as declared on `ReadChartSpec`, and every other
 * field closed off. The `?: never` arms are what turn a field the kind
 * cannot draw into a compile error.
 */
type Fields<K extends SpecField> = Pick<ReadChartSpec, K> & {
  readonly [P in Exclude<SpecField, K>]?: never;
};

/** Fields every chart kind carries. */
const COMMON_FIELDS = [
  'categories',
  'title',
  'titleStyle',
  'titleOverlay',
  'titleLayout',
  'dataLabels',
  'legend',
  'plotAreaLayout',
  'plotAreaFill',
  'plotAreaStrokeColor',
  'chartAreaFill',
  'chartAreaStrokeColor',
  'plotVisibleCellsOnly',
  'dispBlanksAs',
  'date1904',
  'language',
  'roundedCorners',
  'chartStyle',
  'varyColors',
] as const satisfies FieldList;

/** `<c:valAx>` — on every kind but the axis-less pie / doughnut. */
const VALUE_AXIS_FIELDS = [
  'valueAxis',
  'valueAxisTitle',
  'valueAxisTitleStyle',
  'valueAxisTitleRotationDeg',
  'valueAxisLabelStyle',
  'valueAxisLabelRotationDeg',
  'valueAxisHidden',
  'valueAxisMajorGridlines',
  'valueAxisMinorGridlines',
  'valueAxisMajorGridlineColor',
  'valueAxisMinorGridlineColor',
  'valueAxisMajorGridlineWidthEmu',
  'valueAxisMinorGridlineWidthEmu',
  'valueAxisMajorTickMark',
  'valueAxisMinorTickMark',
  'valueAxisTickLabelPos',
  'valueAxisLineColor',
  'valueAxisLineWidthEmu',
  'valueAxisLineHidden',
  'valueAxisOrientation',
  'valueAxisCrosses',
  'valueAxisCrossBetween',
] as const satisfies FieldList;

/**
 * The horizontal axis. One set of fields drives `<c:catAx>` on a category
 * chart and the x-channel `<c:valAx>` of a scatter / bubble chart.
 */
const HORIZONTAL_AXIS_FIELDS = [
  'categoryAxisTitle',
  'categoryAxisTitleStyle',
  'categoryAxisTitleRotationDeg',
  'categoryAxisLabelStyle',
  'categoryAxisLabelRotationDeg',
  'categoryAxisHidden',
  'categoryAxisMajorGridlines',
  'categoryAxisMinorGridlines',
  'categoryAxisMajorGridlineColor',
  'categoryAxisMinorGridlineColor',
  'categoryAxisMajorGridlineWidthEmu',
  'categoryAxisMinorGridlineWidthEmu',
  'categoryAxisMajorTickMark',
  'categoryAxisMinorTickMark',
  'categoryAxisTickLabelPos',
  'categoryAxisLineColor',
  'categoryAxisLineWidthEmu',
  'categoryAxisLineHidden',
  'categoryAxisOrientation',
  'categoryAxisNumberFormat',
] as const satisfies FieldList;

/** Tick and label controls only a discrete category axis has. */
const CATEGORY_AXIS_FIELDS = [
  'categoryAxisTickLabelSkip',
  'categoryAxisTickMarkSkip',
  'categoryAxisLabelOffset',
  'categoryAxisLabelAlign',
  'categoryAxisNoMultiLevelLabel',
] as const satisfies FieldList;

/** The three fields {@link CategoryAxisMode} arbitrates between. */
const CATEGORY_MODE_FIELDS = [
  'categoryAxisDate',
  'categoryAxisScaling',
  'categoryGroupLevels',
] as const satisfies FieldList;

/** Every field of a chart that has axes. */
const AXIS_CHART_FIELDS = [
  ...COMMON_FIELDS,
  ...VALUE_AXIS_FIELDS,
  ...HORIZONTAL_AXIS_FIELDS,
] as const;

/** Every field of a chart drawn against a discrete category axis. */
const CATEGORY_CHART_FIELDS = [
  ...AXIS_CHART_FIELDS,
  ...CATEGORY_AXIS_FIELDS,
  ...CATEGORY_MODE_FIELDS,
  'dataTable',
] as const;

/** Per-variant field sets; the variant types below are derived from these. */
const VARIANT_FIELDS = {
  combo: [
    ...CATEGORY_CHART_FIELDS,
    'grouping',
    'gapWidthPct',
    'overlapPct',
    'dropLines',
    'hiLowLines',
    'lineMarkers',
    'upDownBars',
    'secondaryValueAxis',
  ],
  bar3D: [
    ...CATEGORY_CHART_FIELDS,
    'grouping',
    'gapWidthPct',
    'gapDepthPct',
    'bar3DShape',
    'seriesAxis',
    'view3D',
  ],
  lineArea3D: [
    ...CATEGORY_CHART_FIELDS,
    'grouping',
    'dropLines',
    'gapDepthPct',
    'seriesAxis',
    'view3D',
  ],
  pie: [...COMMON_FIELDS, 'firstSliceAngleDeg'],
  ofPie: [...COMMON_FIELDS, 'ofPie'],
  pie3D: [...COMMON_FIELDS, 'view3D'],
  doughnut: [...COMMON_FIELDS, 'firstSliceAngleDeg', 'holeSizePct'],
  scatter: [...AXIS_CHART_FIELDS, 'categoryAxisScaling', 'scatterStyle'],
  bubble: [
    ...AXIS_CHART_FIELDS,
    'categoryAxisScaling',
    'bubbleScale',
    'bubbleSizeRepresents',
    'bubble3D',
    'showNegativeBubbles',
  ],
  radar: [...CATEGORY_CHART_FIELDS, 'radarStyle'],
  stock: [...CATEGORY_CHART_FIELDS, 'dropLines', 'hiLowLines'],
  candlestick: [...CATEGORY_CHART_FIELDS, 'dropLines', 'hiLowLines', 'upDownBars'],
  surface: [...CATEGORY_CHART_FIELDS, 'surfaceContour', 'surfaceWireframe', 'view3D', 'seriesAxis'],
} as const satisfies Record<string, FieldList>;

/** Name of one arm of {@link ChartSpec}. */
export type ChartVariant = keyof typeof VARIANT_FIELDS;

type VariantField<V extends ChartVariant> = (typeof VARIANT_FIELDS)[V][number];

/**
 * Series channels and combo placement are the only parts of a series that
 * depend on the chart kind; name, styling and per-point overrides are shared.
 */
type ChartSeriesCommon = Omit<
  ChartSeries,
  'xValues' | 'bubbleSizes' | 'chartKind' | 'secondaryAxis'
>;

/** A series plotted as one value channel against the categories. */
export type PlainChartSeries = ChartSeriesCommon & {
  readonly xValues?: never;
  readonly bubbleSizes?: never;
  readonly chartKind?: never;
  readonly secondaryAxis?: never;
};

/**
 * A series on a bar / column / line / area chart. Those four share a plot
 * area, so a series can be drawn as another of them (`chartKind`) or moved
 * to the secondary value axis.
 */
export type ComboChartSeries = ChartSeriesCommon & {
  readonly xValues?: never;
  readonly bubbleSizes?: never;
  readonly chartKind?: 'bar' | 'column' | 'line' | 'area';
  readonly secondaryAxis?: boolean;
};

/** A scatter series: an x channel paired with the y channel. */
export type ScatterChartSeries = ChartSeriesCommon & {
  readonly xValues: ReadonlyArray<number | null>;
  readonly bubbleSizes?: never;
  readonly chartKind?: never;
  readonly secondaryAxis?: never;
};

/** A bubble series: x, y, and the size channel that scales each bubble. */
export type BubbleChartSeries = ChartSeriesCommon & {
  readonly xValues: ReadonlyArray<number | null>;
  readonly bubbleSizes: ReadonlyArray<number | null>;
  readonly chartKind?: never;
  readonly secondaryAxis?: never;
};

/**
 * A category axis is either discrete — optionally grouped into outer
 * levels — or a date axis. Only a numeric axis has a scale, so
 * `categoryAxisScaling` rides with the date form.
 */
type CategoryAxisMode =
  | {
      readonly categoryAxisDate: ChartDateAxis;
      /** Date-serial range of the axis; only `min` / `max` apply. */
      readonly categoryAxisScaling?: ChartAxisScaling;
      readonly categoryGroupLevels?: never;
    }
  | {
      readonly categoryAxisDate?: never;
      readonly categoryAxisScaling?: never;
      readonly categoryGroupLevels?: ReadonlyArray<ReadonlyArray<string>>;
    };

/**
 * A bar / column / line / area chart. Because the four share a plot area,
 * every series-level and group-level control of the family is available
 * whichever of them is the base kind — that is what makes a combo chart.
 */
export type ComboChartSpec = Fields<VariantField<'combo'>> &
  CategoryAxisMode & {
    readonly kind: 'bar' | 'column' | 'line' | 'area';
    readonly series: ReadonlyArray<ComboChartSeries>;
  };

/** A 3-D bar / column chart (`<c:bar3DChart>`). */
export type Bar3DChartSpec = Fields<VariantField<'bar3D'>> &
  CategoryAxisMode & {
    readonly kind: 'bar' | 'column';
    readonly series: ReadonlyArray<PlainChartSeries>;
    /** Its presence is what selects the 3-D element; a combo cannot be 3-D. */
    readonly view3D: ChartView3D;
  };

/**
 * A 3-D line / area chart. High-low lines, up/down bars and point markers
 * have no 3-D element to live in.
 */
export type LineArea3DChartSpec = Fields<VariantField<'lineArea3D'>> &
  CategoryAxisMode & {
    readonly kind: 'line' | 'area';
    readonly series: ReadonlyArray<PlainChartSeries>;
    readonly view3D: ChartView3D;
  };

/** A pie chart. One series, one slice per category. */
export type PieChartSpec = Fields<VariantField<'pie'>> & {
  readonly kind: 'pie';
  readonly series: readonly [PlainChartSeries];
};

/** A pie-of-pie / bar-of-pie chart (`<c:ofPieChart>`). */
export type OfPieChartSpec = Fields<VariantField<'ofPie'>> & {
  readonly kind: 'pie';
  readonly series: readonly [PlainChartSeries];
  readonly ofPie: ChartOfPie;
};

/** A 3-D pie (`<c:pie3DChart>`), which turns with the camera rather than `firstSliceAngleDeg`. */
export type Pie3DChartSpec = Fields<VariantField<'pie3D'>> & {
  readonly kind: 'pie';
  readonly series: readonly [PlainChartSeries];
  readonly view3D: ChartView3D;
};

/** A doughnut chart. */
export type DoughnutChartSpec = Fields<VariantField<'doughnut'>> & {
  readonly kind: 'doughnut';
  readonly series: readonly [PlainChartSeries];
};

/** A scatter chart: both axes are numeric, so there are no categories to group. */
export type ScatterChartSpec = Fields<VariantField<'scatter'>> & {
  readonly kind: 'scatter';
  readonly series: ReadonlyArray<ScatterChartSeries>;
};

/** A bubble chart: a scatter chart whose third channel scales each marker. */
export type BubbleChartSpec = Fields<VariantField<'bubble'>> & {
  readonly kind: 'bubble';
  readonly series: ReadonlyArray<BubbleChartSeries>;
};

/** A radar chart. */
export type RadarChartSpec = Fields<VariantField<'radar'>> &
  CategoryAxisMode & {
    readonly kind: 'radar';
    readonly series: ReadonlyArray<PlainChartSeries>;
  };

/** A stock chart plotted from high, low and close. */
export type StockChartSpec = Fields<VariantField<'stock'>> &
  CategoryAxisMode & {
    readonly kind: 'stock';
    readonly series: readonly [PlainChartSeries, PlainChartSeries, PlainChartSeries];
  };

/**
 * A candlestick chart: a stock chart with an open series, whose up / down
 * bars span open → close.
 */
export type CandlestickChartSpec = Fields<VariantField<'candlestick'>> &
  CategoryAxisMode & {
    readonly kind: 'stock';
    readonly series: readonly [
      PlainChartSeries,
      PlainChartSeries,
      PlainChartSeries,
      PlainChartSeries,
    ];
  };

/** A surface chart. Its three axes make `view3D` a camera, not a subtype switch. */
export type SurfaceChartSpec = Fields<VariantField<'surface'>> &
  CategoryAxisMode & {
    readonly kind: 'surface';
    readonly series: ReadonlyArray<PlainChartSeries>;
  };

/**
 * A chart this library can draw. Each kind carries only the fields its
 * OOXML element has, so a field the chart cannot show is a compile error
 * rather than a value that serializes into nothing — or into a deck
 * PowerPoint offers to repair.
 *
 * `readChartSpec` returns the permissive `ReadChartSpec` instead, because a
 * deck authored elsewhere can combine fields no kind allows; narrow one back
 * with {@link isChartSpec}.
 */
export type ChartSpec =
  | ComboChartSpec
  | Bar3DChartSpec
  | LineArea3DChartSpec
  | PieChartSpec
  | OfPieChartSpec
  | Pie3DChartSpec
  | DoughnutChartSpec
  | ScatterChartSpec
  | BubbleChartSpec
  | RadarChartSpec
  | StockChartSpec
  | CandlestickChartSpec
  | SurfaceChartSpec;

// CT_StockChart takes 3 series (high, low, close) or 4 (open first).
export const STOCK_SERIES_WITHOUT_OPEN = 3;
export const STOCK_SERIES_WITH_OPEN = 4;

const fieldSet = (fields: FieldList): ReadonlySet<string> => new Set<string>(fields);

// Spelled out so the record stays total: a new variant that is not listed
// here, or one listed that no longer exists, fails to compile.
const VARIANT_FIELD_SETS: Readonly<Record<ChartVariant, ReadonlySet<string>>> = {
  combo: fieldSet(VARIANT_FIELDS.combo),
  bar3D: fieldSet(VARIANT_FIELDS.bar3D),
  lineArea3D: fieldSet(VARIANT_FIELDS.lineArea3D),
  pie: fieldSet(VARIANT_FIELDS.pie),
  ofPie: fieldSet(VARIANT_FIELDS.ofPie),
  pie3D: fieldSet(VARIANT_FIELDS.pie3D),
  doughnut: fieldSet(VARIANT_FIELDS.doughnut),
  scatter: fieldSet(VARIANT_FIELDS.scatter),
  bubble: fieldSet(VARIANT_FIELDS.bubble),
  radar: fieldSet(VARIANT_FIELDS.radar),
  stock: fieldSet(VARIANT_FIELDS.stock),
  candlestick: fieldSet(VARIANT_FIELDS.candlestick),
  surface: fieldSet(VARIANT_FIELDS.surface),
};

/**
 * Which arm of {@link ChartSpec} a spec would inhabit, from the fields that
 * pick the OOXML element. It says nothing about whether the rest of the
 * spec fits — {@link isChartSpec} decides that.
 */
const chartVariant = (spec: ReadChartSpec): ChartVariant => {
  switch (spec.kind) {
    case 'bar':
    case 'column':
      return spec.view3D !== undefined ? 'bar3D' : 'combo';
    case 'line':
    case 'area':
      return spec.view3D !== undefined ? 'lineArea3D' : 'combo';
    case 'pie':
      if (spec.view3D !== undefined) return 'pie3D';
      return spec.ofPie !== undefined ? 'ofPie' : 'pie';
    case 'stock':
      return spec.series.length === STOCK_SERIES_WITH_OPEN ? 'candlestick' : 'stock';
    default:
      return spec.kind;
  }
};

const SINGLE_SERIES_VARIANTS: ReadonlySet<ChartVariant> = new Set([
  'pie',
  'ofPie',
  'pie3D',
  'doughnut',
]);

const seriesFits = (spec: ReadChartSpec, variant: ChartVariant): boolean => {
  if (SINGLE_SERIES_VARIANTS.has(variant)) return spec.series.length === 1;
  if (variant === 'stock') return spec.series.length === STOCK_SERIES_WITHOUT_OPEN;
  if (variant === 'candlestick') return spec.series.length === STOCK_SERIES_WITH_OPEN;
  const wantsX = variant === 'scatter' || variant === 'bubble';
  return spec.series.every(
    (series) =>
      (series.xValues !== undefined) === wantsX &&
      (series.bubbleSizes !== undefined) === (variant === 'bubble') &&
      (variant === 'combo' ||
        (series.chartKind === undefined && series.secondaryAxis === undefined)),
  );
};

/**
 * Whether a spec read back from a deck also inhabits the write-side
 * {@link ChartSpec}. A deck authored elsewhere can combine fields no kind
 * draws, so a read-modify-write pass has to narrow before it writes.
 */
export const isChartSpec = (spec: ReadChartSpec): spec is ChartSpec => {
  const variant = chartVariant(spec);
  const allowed = VARIANT_FIELD_SETS[variant];
  for (const [field, value] of Object.entries(spec)) {
    if (field === 'kind' || field === 'series' || value === undefined) continue;
    if (!allowed.has(field)) return false;
  }
  // A text category axis has no scale, and a date axis has no outer levels.
  if (spec.categoryAxisDate === undefined) {
    if (spec.categoryAxisScaling !== undefined && spec.kind !== 'scatter' && spec.kind !== 'bubble')
      return false;
  } else if (spec.categoryGroupLevels !== undefined) {
    return false;
  }
  return seriesFits(spec, variant);
};

/** Every field of `ReadChartSpec` that some variant draws. Used by the tests. */
export const ALL_VARIANT_FIELDS: ReadonlySet<string> = new Set(
  Object.values(VARIANT_FIELDS).flat(),
);

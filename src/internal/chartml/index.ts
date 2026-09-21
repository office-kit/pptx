// internal/chartml — c: namespace: charts + minimal xlsx writer for embedded data.
// Allowed imports: internal/drawingml, internal/parts, internal/xml.

export type {
  ChartAxisScaling,
  ChartDataLabelPosition,
  ChartDataLabels,
  ChartDataTable,
  ChartDateAxis,
  ChartErrorBarAmount,
  ChartErrorBars,
  ChartGrouping,
  ChartKind,
  ChartManualLayout,
  ChartOfPie,
  ChartSecondaryValueAxis,
  ChartSeries,
  ChartSeriesAxis,
  ReadChartSpec,
  ChartTextStyle,
  ChartTimeUnit,
  ChartTrendline,
  ChartUpDownBars,
  ChartView3D,
} from './types.ts';
export type {
  Bar3DChartSpec,
  BubbleChartSeries,
  BubbleChartSpec,
  CandlestickChartSpec,
  ChartSpec,
  ChartVariant,
  ComboChartSeries,
  ComboChartSpec,
  DoughnutChartSpec,
  LineArea3DChartSpec,
  OfPieChartSpec,
  Pie3DChartSpec,
  PieChartSpec,
  PlainChartSeries,
  RadarChartSpec,
  ScatterChartSeries,
  ScatterChartSpec,
  StockChartSpec,
  SurfaceChartSpec,
} from './chart-spec.ts';
export { ALL_VARIANT_FIELDS, isChartSpec } from './chart-spec.ts';
export { buildChartSpaceDoc } from './chart-builder.ts';
export { readChartSpec } from './chart-reader.ts';
export type { SheetCell } from './embedded-xlsx.ts';
export { buildEmbeddedXlsx, cellAddr, cellRange } from './embedded-xlsx.ts';
export { layoutChartSheet } from './sheet-layout.ts';

import { oneOf } from '../bounds.ts';
import { LINE_DASHES } from '../enum-values.ts';
import type { ChartSpec, ChartDataLabels, ChartAxisScaling } from './types.ts';

export const LABEL_POSITIONS = [
  'ctr',
  'inEnd',
  'outEnd',
  'inBase',
  't',
  'b',
  'l',
  'r',
  'bestFit',
] as const;
export const DISPLAY_UNITS = [
  'hundreds',
  'thousands',
  'tenThousands',
  'hundredThousands',
  'millions',
  'tenMillions',
  'hundredMillions',
  'billions',
  'trillions',
] as const;
export const LABEL_ALIGNMENTS = ['ctr', 'l', 'r'] as const;
export const DISPLAY_BLANKS = ['gap', 'zero', 'span'] as const;
export const SCATTER_STYLES = [
  'none',
  'line',
  'lineMarker',
  'marker',
  'smooth',
  'smoothMarker',
] as const;
export const RADAR_STYLES = ['standard', 'marker', 'filled'] as const;
export const BAR_3D_SHAPES = [
  'box',
  'cone',
  'coneToMax',
  'cylinder',
  'pyramid',
  'pyramidToMax',
] as const;
export const OF_PIE_TYPES = ['pie', 'bar'] as const;
export const SPLIT_TYPES = ['auto', 'cust', 'percent', 'pos', 'val'] as const;
export const LAYOUT_TARGETS = ['inner', 'outer'] as const;
export const AXIS_CROSSES = ['autoZero', 'min', 'max'] as const;
export const LEGEND_POSITIONS = ['r', 't', 'b', 'l', 'tr'] as const;
export const MARKER_SYMBOLS = [
  'none',
  'auto',
  'circle',
  'square',
  'diamond',
  'triangle',
  'star',
  'x',
  'plus',
  'dash',
  'dot',
  'picture',
] as const;
export const TRENDLINE_TYPES = ['linear', 'exp', 'log', 'poly', 'power', 'movingAvg'] as const;
export const ERROR_BAR_TYPES = ['both', 'plus', 'minus'] as const;
export const ERROR_VALUE_TYPES = ['fixedVal', 'percentage', 'stdDev', 'stdErr', 'cust'] as const;

export const GROUPINGS = ['clustered', 'stacked', 'percentStacked', 'standard'] as const;
export const TICK_MARKS = ['in', 'out', 'cross', 'none'] as const;
export const TICK_LABEL_POSITIONS = ['none', 'low', 'high', 'nextTo'] as const;
export const ORIENTATIONS = ['minMax', 'maxMin'] as const;
export const CROSS_BETWEEN = ['between', 'midCat'] as const;
export const TIME_UNITS = ['days', 'months', 'years'] as const;

const validateLabels = (labels: ChartDataLabels | null | undefined, field: string): void => {
  if (labels?.position !== undefined) oneOf(labels.position, LABEL_POSITIONS, `${field}.position`);
};

const validateScaling = (scaling: ChartAxisScaling | undefined, field: string): void => {
  if (scaling?.displayUnits !== undefined)
    oneOf(scaling.displayUnits, DISPLAY_UNITS, `${field}.displayUnits`);
};

/** Validate before chart XML or its embedded workbook is changed. */
export const validateChartSpecEnums = (spec: ChartSpec, caller: string): void => {
  oneOf(
    spec.kind,
    [
      'bar',
      'column',
      'line',
      'pie',
      'doughnut',
      'area',
      'scatter',
      'radar',
      'bubble',
      'stock',
      'surface',
    ],
    `${caller}: kind`,
  );
  if (spec.grouping !== undefined) oneOf(spec.grouping, GROUPINGS, `${caller}: grouping`);
  if (spec.valueAxisMajorTickMark !== undefined)
    oneOf(spec.valueAxisMajorTickMark, TICK_MARKS, `${caller}: valueAxisMajorTickMark`);
  if (spec.valueAxisMinorTickMark !== undefined)
    oneOf(spec.valueAxisMinorTickMark, TICK_MARKS, `${caller}: valueAxisMinorTickMark`);
  if (spec.categoryAxisMajorTickMark !== undefined)
    oneOf(spec.categoryAxisMajorTickMark, TICK_MARKS, `${caller}: categoryAxisMajorTickMark`);
  if (spec.categoryAxisMinorTickMark !== undefined)
    oneOf(spec.categoryAxisMinorTickMark, TICK_MARKS, `${caller}: categoryAxisMinorTickMark`);
  if (spec.valueAxisTickLabelPos !== undefined)
    oneOf(spec.valueAxisTickLabelPos, TICK_LABEL_POSITIONS, `${caller}: valueAxisTickLabelPos`);
  if (spec.categoryAxisTickLabelPos !== undefined)
    oneOf(
      spec.categoryAxisTickLabelPos,
      TICK_LABEL_POSITIONS,
      `${caller}: categoryAxisTickLabelPos`,
    );
  if (spec.categoryAxisLabelAlign !== undefined)
    oneOf(spec.categoryAxisLabelAlign, LABEL_ALIGNMENTS, `${caller}: categoryAxisLabelAlign`);
  if (spec.categoryAxisOrientation !== undefined)
    oneOf(spec.categoryAxisOrientation, ORIENTATIONS, `${caller}: categoryAxisOrientation`);
  if (spec.valueAxisOrientation !== undefined)
    oneOf(spec.valueAxisOrientation, ORIENTATIONS, `${caller}: valueAxisOrientation`);
  if (spec.valueAxisCrossBetween !== undefined)
    oneOf(spec.valueAxisCrossBetween, CROSS_BETWEEN, `${caller}: valueAxisCrossBetween`);
  if (spec.dispBlanksAs !== undefined)
    oneOf(spec.dispBlanksAs, DISPLAY_BLANKS, `${caller}: dispBlanksAs`);
  if (spec.scatterStyle !== undefined)
    oneOf(spec.scatterStyle, SCATTER_STYLES, `${caller}: scatterStyle`);
  if (spec.radarStyle !== undefined) oneOf(spec.radarStyle, RADAR_STYLES, `${caller}: radarStyle`);
  if (spec.bubbleSizeRepresents !== undefined)
    oneOf(spec.bubbleSizeRepresents, ['area', 'width'], `${caller}: bubbleSizeRepresents`);
  if (spec.bar3DShape !== undefined) oneOf(spec.bar3DShape, BAR_3D_SHAPES, `${caller}: bar3DShape`);
  if (spec.secondaryValueAxis?.majorTickMark !== undefined)
    oneOf(
      spec.secondaryValueAxis.majorTickMark,
      TICK_MARKS,
      `${caller}: secondaryValueAxis.majorTickMark`,
    );
  if (spec.secondaryValueAxis?.minorTickMark !== undefined)
    oneOf(
      spec.secondaryValueAxis.minorTickMark,
      TICK_MARKS,
      `${caller}: secondaryValueAxis.minorTickMark`,
    );
  if (spec.secondaryValueAxis?.tickLabelPos !== undefined)
    oneOf(
      spec.secondaryValueAxis.tickLabelPos,
      TICK_LABEL_POSITIONS,
      `${caller}: secondaryValueAxis.tickLabelPos`,
    );
  if (spec.secondaryValueAxis?.crossBetween !== undefined)
    oneOf(
      spec.secondaryValueAxis.crossBetween,
      CROSS_BETWEEN,
      `${caller}: secondaryValueAxis.crossBetween`,
    );
  if (spec.seriesAxis?.orientation !== undefined)
    oneOf(spec.seriesAxis.orientation, ORIENTATIONS, `${caller}: seriesAxis.orientation`);
  if (spec.seriesAxis?.tickLabelPos !== undefined)
    oneOf(spec.seriesAxis.tickLabelPos, TICK_LABEL_POSITIONS, `${caller}: seriesAxis.tickLabelPos`);
  if (spec.categoryAxisDate?.baseTimeUnit !== undefined)
    oneOf(
      spec.categoryAxisDate.baseTimeUnit,
      TIME_UNITS,
      `${caller}: categoryAxisDate.baseTimeUnit`,
    );
  if (spec.categoryAxisDate?.majorTimeUnit !== undefined)
    oneOf(
      spec.categoryAxisDate.majorTimeUnit,
      TIME_UNITS,
      `${caller}: categoryAxisDate.majorTimeUnit`,
    );
  if (spec.categoryAxisDate?.minorTimeUnit !== undefined)
    oneOf(
      spec.categoryAxisDate.minorTimeUnit,
      TIME_UNITS,
      `${caller}: categoryAxisDate.minorTimeUnit`,
    );
  if (spec.ofPie?.type !== undefined) oneOf(spec.ofPie.type, OF_PIE_TYPES, `${caller}: ofPie.type`);
  if (spec.ofPie?.splitType !== undefined)
    oneOf(spec.ofPie.splitType, SPLIT_TYPES, `${caller}: ofPie.splitType`);
  if (spec.plotAreaLayout?.target !== undefined)
    oneOf(spec.plotAreaLayout.target, LAYOUT_TARGETS, `${caller}: plotAreaLayout.target`);
  if (spec.legend?.layout?.target !== undefined)
    oneOf(spec.legend.layout.target, LAYOUT_TARGETS, `${caller}: legend.layout.target`);
  if (spec.valueAxisCrosses !== undefined && typeof spec.valueAxisCrosses !== 'object')
    oneOf(spec.valueAxisCrosses, AXIS_CROSSES, `${caller}: valueAxisCrosses`);
  if (spec.legend?.position != null)
    oneOf(spec.legend.position, LEGEND_POSITIONS, `${caller}: legend.position`);
  validateLabels(spec.dataLabels, `${caller}: dataLabels`);
  validateScaling(spec.valueAxis, `${caller}: valueAxis`);
  validateScaling(spec.categoryAxisScaling, `${caller}: categoryAxisScaling`);
  validateScaling(spec.secondaryValueAxis?.scaling, `${caller}: secondaryValueAxis.scaling`);
  for (const [index, series] of spec.series.entries()) {
    const field = `${caller}: series[${index}]`;
    if (series.chartKind !== undefined)
      oneOf(series.chartKind, ['bar', 'column', 'line', 'area'], `${field}.chartKind`);
    if (series.lineDash !== undefined) oneOf(series.lineDash, LINE_DASHES, `${field}.lineDash`);
    if (series.markerSymbol !== undefined)
      oneOf(series.markerSymbol, MARKER_SYMBOLS, `${field}.markerSymbol`);
    if (series.trendline !== undefined)
      oneOf(series.trendline.type, TRENDLINE_TYPES, `${field}.trendline.type`);
    validateLabels(series.dataLabels, `${field}.dataLabels`);
    for (const [point, labels] of (series.pointDataLabels ?? []).entries())
      validateLabels(labels, `${field}.pointDataLabels[${point}]`);
    for (const direction of ['errorBars', 'xErrorBars'] as const) {
      const bars = series[direction];
      if (bars === undefined) continue;
      oneOf(bars.barType, ERROR_BAR_TYPES, `${field}.${direction}.barType`);
      oneOf(bars.amount.type, ERROR_VALUE_TYPES, `${field}.${direction}.amount.type`);
    }
  }
};

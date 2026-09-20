import { oneOf } from '../bounds.ts';
import { LINE_DASHES } from '../enum-values.ts';
import type { ChartSpec, ChartDataLabels, ChartAxisScaling } from './types.ts';

const TICK_MARKS = ['in', 'out', 'cross', 'none'] as const;
const TICK_LABEL_POSITIONS = ['none', 'low', 'high', 'nextTo'] as const;
const ORIENTATIONS = ['minMax', 'maxMin'] as const;
const CROSS_BETWEEN = ['between', 'midCat'] as const;
const TIME_UNITS = ['days', 'months', 'years'] as const;

const validateLabels = (labels: ChartDataLabels | null | undefined, field: string): void => {
  if (labels?.position !== undefined)
    oneOf(
      labels.position,
      ['ctr', 'inEnd', 'outEnd', 'inBase', 't', 'b', 'l', 'r', 'bestFit'],
      `${field}.position`,
    );
};

const validateScaling = (scaling: ChartAxisScaling | undefined, field: string): void => {
  if (scaling?.displayUnits !== undefined)
    oneOf(
      scaling.displayUnits,
      [
        'hundreds',
        'thousands',
        'tenThousands',
        'hundredThousands',
        'millions',
        'tenMillions',
        'hundredMillions',
        'billions',
        'trillions',
      ],
      `${field}.displayUnits`,
    );
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
  if (spec.grouping !== undefined)
    oneOf(
      spec.grouping,
      ['clustered', 'stacked', 'percentStacked', 'standard'],
      `${caller}: grouping`,
    );
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
    oneOf(spec.categoryAxisLabelAlign, ['ctr', 'l', 'r'], `${caller}: categoryAxisLabelAlign`);
  if (spec.categoryAxisOrientation !== undefined)
    oneOf(spec.categoryAxisOrientation, ORIENTATIONS, `${caller}: categoryAxisOrientation`);
  if (spec.valueAxisOrientation !== undefined)
    oneOf(spec.valueAxisOrientation, ORIENTATIONS, `${caller}: valueAxisOrientation`);
  if (spec.valueAxisCrossBetween !== undefined)
    oneOf(spec.valueAxisCrossBetween, CROSS_BETWEEN, `${caller}: valueAxisCrossBetween`);
  if (spec.dispBlanksAs !== undefined)
    oneOf(spec.dispBlanksAs, ['gap', 'zero', 'span'], `${caller}: dispBlanksAs`);
  if (spec.scatterStyle !== undefined)
    oneOf(
      spec.scatterStyle,
      ['none', 'line', 'lineMarker', 'marker', 'smooth', 'smoothMarker'],
      `${caller}: scatterStyle`,
    );
  if (spec.radarStyle !== undefined)
    oneOf(spec.radarStyle, ['standard', 'marker', 'filled'], `${caller}: radarStyle`);
  if (spec.bubbleSizeRepresents !== undefined)
    oneOf(spec.bubbleSizeRepresents, ['area', 'width'], `${caller}: bubbleSizeRepresents`);
  if (spec.bar3DShape !== undefined)
    oneOf(
      spec.bar3DShape,
      ['box', 'cone', 'coneToMax', 'cylinder', 'pyramid', 'pyramidToMax'],
      `${caller}: bar3DShape`,
    );
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
  if (spec.ofPie?.type !== undefined)
    oneOf(spec.ofPie.type, ['pie', 'bar'], `${caller}: ofPie.type`);
  if (spec.ofPie?.splitType !== undefined)
    oneOf(
      spec.ofPie.splitType,
      ['auto', 'cust', 'percent', 'pos', 'val'],
      `${caller}: ofPie.splitType`,
    );
  if (spec.plotAreaLayout?.target !== undefined)
    oneOf(spec.plotAreaLayout.target, ['inner', 'outer'], `${caller}: plotAreaLayout.target`);
  if (spec.legend?.layout?.target !== undefined)
    oneOf(spec.legend.layout.target, ['inner', 'outer'], `${caller}: legend.layout.target`);
  if (spec.valueAxisCrosses !== undefined && typeof spec.valueAxisCrosses !== 'object')
    oneOf(spec.valueAxisCrosses, ['autoZero', 'min', 'max'], `${caller}: valueAxisCrosses`);
  if (spec.legend?.position != null)
    oneOf(spec.legend.position, ['r', 't', 'b', 'l', 'tr'], `${caller}: legend.position`);
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
      oneOf(
        series.markerSymbol,
        [
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
        ],
        `${field}.markerSymbol`,
      );
    if (series.trendline !== undefined)
      oneOf(
        series.trendline.type,
        ['linear', 'exp', 'log', 'poly', 'power', 'movingAvg'],
        `${field}.trendline.type`,
      );
    validateLabels(series.dataLabels, `${field}.dataLabels`);
    for (const [point, labels] of (series.pointDataLabels ?? []).entries())
      validateLabels(labels, `${field}.pointDataLabels[${point}]`);
    for (const direction of ['errorBars', 'xErrorBars'] as const) {
      const bars = series[direction];
      if (bars === undefined) continue;
      oneOf(bars.barType, ['both', 'plus', 'minus'], `${field}.${direction}.barType`);
      oneOf(
        bars.amount.type,
        ['fixedVal', 'percentage', 'stdDev', 'stdErr', 'cust'],
        `${field}.${direction}.amount.type`,
      );
    }
  }
};

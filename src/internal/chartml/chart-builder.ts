// Chart XML builder.
//
// Produces a complete `<c:chartSpace>` for any `ChartSpec` — every
// plot-group element of CT_PlotArea. The chart references an embedded
// xlsx via `<c:externalData r:id="rId1">`; the calling layer is
// responsible for wiring that rel and writing the xlsx bytes. Inline
// `<c:strCache>` / `<c:numCache>` blocks carry the values so PowerPoint
// can render the chart without ever opening the workbook.

import {
  boundedInt,
  firstSliceAngle,
  gapAmountPercent,
  holeSizePercent,
  lineWidthEmu as validateLineWidthEmu,
  overlapPercent,
} from '../bounds.ts';
import { NS, type XmlDocument, type XmlElement, attr, elem, qname, text } from '../xml/index.ts';
import { type ChartSheetLayout, dateSerial, layoutChartSheet } from './sheet-layout.ts';
import type {
  ChartAxisScaling,
  ChartDataLabels,
  ChartErrorBars,
  ChartManualLayout,
  ChartTextStyle,
  ChartView3D,
  ReadChartSpec,
} from './types.ts';
import { STOCK_SERIES_WITHOUT_OPEN, STOCK_SERIES_WITH_OPEN, type ChartSpec } from './chart-spec.ts';

// QNames (chart `c:` namespace) --------------------------------------------

const NS_C = NS.chart;
const NS_A = NS.dml;
const NS_R = NS.officeDocRels;

const c = (local: string): { prefix: string; localName: string; namespaceURI: string } =>
  qname('c', local, NS_C);
const a = (local: string): { prefix: string; localName: string; namespaceURI: string } =>
  qname('a', local, NS_A);
const r = (local: string): { prefix: string; localName: string; namespaceURI: string } =>
  qname('r', local, NS_R);

const ATTR_VAL = qname('', 'val', '');
const ATTR_IDX = qname('', 'idx', '');

// Helpers ------------------------------------------------------------------

const valNode = (name: ReturnType<typeof c>, val: string | number): XmlElement =>
  elem(name, { attrs: [attr(ATTR_VAL, String(val))] });

const ptNode = (idx: number, value: string): XmlElement =>
  elem(c('pt'), {
    attrs: [attr(ATTR_IDX, String(idx))],
    children: [elem(c('v'), { children: [text(value)] })],
  });

const strRef = (formula: string, points: ReadonlyArray<string>): XmlElement => {
  const strCache = elem(c('strCache'), {
    children: [valNode(c('ptCount'), points.length), ...points.map((p, i) => ptNode(i, p))],
  });
  return elem(c('strRef'), {
    children: [elem(c('f'), { children: [text(formula)] }), strCache],
  });
};

// The CT_NumData body shared by `<c:numCache>` and `<c:numLit>`. Empty
// points are left out; `ptCount` still counts them.
const numDataChildren = (
  points: ReadonlyArray<number | null>,
  formatCode: string,
): XmlElement[] => [
  elem(c('formatCode'), { children: [text(formatCode)] }),
  valNode(c('ptCount'), points.length),
  ...points
    .map((v, i) => (v === null ? null : ptNode(i, String(v))))
    .filter((n): n is XmlElement => n !== null),
];

const numRef = (
  formula: string,
  points: ReadonlyArray<number | null>,
  formatCode = 'General',
): XmlElement =>
  elem(c('numRef'), {
    children: [
      elem(c('f'), { children: [text(formula)] }),
      elem(c('numCache'), { children: numDataChildren(points, formatCode) }),
    ],
  });

// `<c:multiLvlStrRef>` — `levels[0]` is the innermost level (the category
// labels themselves), as CT_MultiLvlStrData orders its `<c:lvl>` children.
// An outer level's empty entries mean "same group as before" and are left
// out, which is how PowerPoint spans one label over several categories.
const multiLvlStrRef = (
  formula: string,
  levels: ReadonlyArray<ReadonlyArray<string>>,
  pointCount: number,
): XmlElement =>
  elem(c('multiLvlStrRef'), {
    children: [
      elem(c('f'), { children: [text(formula)] }),
      elem(c('multiLvlStrCache'), {
        children: [
          valNode(c('ptCount'), pointCount),
          ...levels.map((level, depth) =>
            elem(c('lvl'), {
              children: level
                .map((label, i) => (depth > 0 && label === '' ? null : ptNode(i, label)))
                .filter((n): n is XmlElement => n !== null),
            }),
          ),
        ],
      }),
    ],
  });

const hexOf = (color: string): string => color.replace(/^#/, '').toUpperCase();

// `<a:solidFill><a:srgbClr val>` with an optional `<a:alpha>`; `opacity` is
// 0..1 and the wire unit is 1/1000 percent.
const solidFill = (color: string, opacity?: number): XmlElement =>
  elem(a('solidFill'), {
    children: [
      elem(a('srgbClr'), {
        attrs: [attr(ATTR_VAL, hexOf(color))],
        children:
          opacity !== undefined
            ? [
                valNode(
                  a('alpha'),
                  boundedInt(opacity * 100000, 'percent100000', 'chart series: fillOpacity'),
                ),
              ]
            : [],
      }),
    ],
  });

const solidFillSpPr = (color: string, opacity?: number): XmlElement =>
  elem(c('spPr'), { children: [solidFill(color, opacity)] });

// <c:majorGridlines|minorGridlines> with optional spPr/ln/solidFill/srgbClr
// for the gridline color. Centralizes the per-gridline color emit so all
// four axis-gridline slots share the same shape.
// `<c:spPr><a:ln w?>` with an optional solid color, or null when neither is
// authored. Every child of <a:ln> is optional, so a width alone is valid.
const lineSpPr = (
  color: string | undefined,
  widthEmu: number | undefined,
  field: string,
  // An axis line can be switched off while its labels stay; a hidden line
  // has no color or width left to write.
  hidden = false,
): XmlElement | null => {
  if (hidden) {
    return elem(c('spPr'), { children: [elem(a('ln'), { children: [elem(a('noFill'))] })] });
  }
  if (color === undefined && widthEmu === undefined) return null;
  const ln = elem(a('ln'), {
    attrs:
      widthEmu !== undefined
        ? [attr(qname('', 'w', ''), String(validateLineWidthEmu(widthEmu, field)))]
        : [],
    children:
      color !== undefined
        ? [
            elem(a('solidFill'), {
              children: [
                elem(a('srgbClr'), {
                  attrs: [attr(qname('', 'val', ''), color.replace(/^#/, '').toUpperCase())],
                }),
              ],
            }),
          ]
        : [],
  });
  return elem(c('spPr'), { children: [ln] });
};

const gridlinesElement = (
  local: 'majorGridlines' | 'minorGridlines',
  color: string | undefined,
  widthEmu: number | undefined,
  field: string,
): XmlElement => {
  const spPr = lineSpPr(color, widthEmu, field);
  return spPr === null ? elem(c(local)) : elem(c(local), { children: [spPr] });
};

// Generic <c:spPr> with optional fill color + line color. Used for the
// chart-area / plot-area background, where authors set one or both.
const spPrChildren = (fill: string | undefined, stroke: string | undefined): XmlElement => {
  const out: XmlElement[] = [];
  if (fill !== undefined) {
    out.push(
      elem(a('solidFill'), {
        children: [
          elem(a('srgbClr'), {
            attrs: [attr(qname('', 'val', ''), fill.replace(/^#/, '').toUpperCase())],
          }),
        ],
      }),
    );
  }
  if (stroke !== undefined) {
    out.push(
      elem(a('ln'), {
        children: [
          elem(a('solidFill'), {
            children: [
              elem(a('srgbClr'), {
                attrs: [attr(qname('', 'val', ''), stroke.replace(/^#/, '').toUpperCase())],
              }),
            ],
          }),
        ],
      }),
    );
  }
  return elem(c('spPr'), { children: out });
};

// Default theme accent palette (matches Office 2013+ default theme).
const DEFAULT_ACCENT_COLORS = [
  '4472C4', // accent1
  'ED7D31', // accent2
  'A5A5A5', // accent3
  'FFC000', // accent4
  '5B9BD5', // accent5
  '70AD47', // accent6
];

// Builds <c:spPr> for a series — solidFill + <a:ln> (line color, width,
// dash). The reader extracts color from solidFill and the line props from
// the ln; this writer keeps them in lock-step.
const seriesSpPr = (
  color: string,
  // `null` hides the connecting line (`<a:ln><a:noFill/>`). That, not the
  // chart-level style token, is what PowerPoint honors for a markers-only
  // scatter series and for the invisible series of a stock chart.
  lineColor: string | null,
  lineWidthEmu: number | undefined,
  lineDash: string | undefined,
  fillOpacity: number | undefined,
): XmlElement => {
  const lnChildren: XmlElement[] = [lineColor === null ? elem(a('noFill')) : solidFill(lineColor)];
  if (lineColor !== null && lineDash !== undefined) {
    lnChildren.push(elem(a('prstDash'), { attrs: [attr(qname('', 'val', ''), lineDash)] }));
  }
  const ln =
    lineWidthEmu !== undefined
      ? elem(a('ln'), {
          attrs: [
            attr(
              qname('', 'w', ''),
              String(validateLineWidthEmu(lineWidthEmu, 'chart series: lineWidthEmu')),
            ),
          ],
          children: lnChildren,
        })
      : elem(a('ln'), { children: lnChildren });
  return elem(c('spPr'), { children: [solidFill(color, fillOpacity), ln] });
};

// `<c:marker>` for a series. PowerPoint paints a marker without <c:spPr> in
// the theme's automatic series color, not the series' own color, so the fill
// and outline colors are always written.
const markerElement = (
  symbol: string | undefined,
  sizePt: number | undefined,
  color: string,
  lineColor: string,
): XmlElement => {
  const children: XmlElement[] = [];
  if (symbol !== undefined) children.push(valNode(c('symbol'), symbol));
  if (sizePt !== undefined) children.push(valNode(c('size'), Math.round(sizePt)));
  if (symbol !== 'none') {
    const fill = (hex: string): XmlElement =>
      elem(a('solidFill'), {
        children: [elem(a('srgbClr'), { attrs: [attr(qname('', 'val', ''), hex)] })],
      });
    children.push(
      elem(c('spPr'), { children: [fill(color), elem(a('ln'), { children: [fill(lineColor)] })] }),
    );
  }
  return elem(c('marker'), { children });
};

// Build the array of `<c:dPt>` overrides for a series — combines the
// sparse pointColors and pointExplosions maps. Each authored index
// emits a `<c:dPt>` with `<c:idx>` + `<c:bubble3D val="0"/>` (required
// by PowerPoint) + optional explosion + optional spPr/solidFill color.
const dPtElements = (
  colors: ReadonlyArray<string | null> | undefined,
  explosions: ReadonlyArray<number | null> | undefined,
  fillOpacity: number | undefined,
  bubble3D: boolean,
): XmlElement[] => {
  const out: XmlElement[] = [];
  const colorLen = colors?.length ?? 0;
  const explLen = explosions?.length ?? 0;
  const max = Math.max(colorLen, explLen);
  for (let i = 0; i < max; i++) {
    const color = colors?.[i] ?? null;
    const expl = explosions?.[i] ?? null;
    if (color === null && (expl === null || !Number.isFinite(expl))) continue;
    const children: XmlElement[] = [
      valNode(c('idx'), i),
      // PowerPoint expects bubble3D on every dPt, 0 outside 3-D bubble charts.
      valNode(c('bubble3D'), bubble3D ? '1' : '0'),
    ];
    if (expl !== null) children.push(valNode(c('explosion'), Math.round(expl)));
    if (color !== null) children.push(solidFillSpPr(color, fillOpacity));
    out.push(elem(c('dPt'), { children }));
  }
  return out;
};

// `<c:trendline>` for a series. Carries trendlineType + optional
// period (movingAvg) / order (poly) / forward / backward / spPr color.
const trendlineElement = (
  tl: NonNullable<NonNullable<ChartSpec['series'][number]['trendline']>>,
): XmlElement => {
  const children: XmlElement[] = [];
  // CT_Trendline schema order: <c:name> first, before <c:spPr>.
  if (tl.name !== undefined) {
    children.push(elem(c('name'), { children: [text(tl.name)] }));
  }
  if (tl.color !== undefined) {
    const hex = tl.color.replace(/^#/, '').toUpperCase();
    const ln = elem(a('ln'), {
      children: [
        elem(a('solidFill'), {
          children: [elem(a('srgbClr'), { attrs: [attr(qname('', 'val', ''), hex)] })],
        }),
      ],
    });
    children.push(elem(c('spPr'), { children: [ln] }));
  }
  children.push(valNode(c('trendlineType'), tl.type));
  if (tl.type === 'movingAvg' && tl.period !== undefined) {
    children.push(valNode(c('period'), tl.period));
  }
  if (tl.type === 'poly' && tl.order !== undefined) {
    children.push(valNode(c('order'), tl.order));
  }
  if (tl.forward !== undefined) children.push(valNode(c('forward'), tl.forward));
  if (tl.backward !== undefined) children.push(valNode(c('backward'), tl.backward));
  // CT_Trendline schema order: dispRSqr before dispEq, both after
  // forward/backward and before trendlineLbl/extLst.
  if (tl.displayRSquared) children.push(valNode(c('dispRSqr'), '1'));
  if (tl.displayEquation) children.push(valNode(c('dispEq'), '1'));
  return elem(c('trendline'), { children });
};

// `<c:errBars>` (CT_ErrBars). `direction` is written for the xy kinds only:
// their series take one element per direction, while a category series has
// a single value direction and PowerPoint leaves `<c:errDir>` out.
const errBarsElement = (bars: ChartErrorBars, direction: 'x' | 'y' | null): XmlElement => {
  const children: XmlElement[] = [];
  if (direction !== null) children.push(valNode(c('errDir'), direction));
  children.push(valNode(c('errBarType'), bars.barType), valNode(c('errValType'), bars.amount.type));
  if (bars.noEndCap !== undefined) children.push(valNode(c('noEndCap'), bars.noEndCap ? '1' : '0'));
  if (bars.amount.type === 'cust') {
    // Literal amounts: error margins rarely live in the chart's data sheet,
    // and a `<c:numLit>` needs no workbook range to stay valid.
    const literal = (name: 'plus' | 'minus', points: ReadonlyArray<number | null>): XmlElement =>
      elem(c(name), {
        children: [elem(c('numLit'), { children: numDataChildren(points, 'General') })],
      });
    if (bars.amount.plus !== undefined) children.push(literal('plus', bars.amount.plus));
    if (bars.amount.minus !== undefined) children.push(literal('minus', bars.amount.minus));
  } else if (bars.amount.type !== 'stdErr') {
    children.push(valNode(c('val'), bars.amount.value));
  }
  const spPr = lineSpPr(bars.color, bars.lineWidthEmu, 'chart series: errorBars.lineWidthEmu');
  if (spPr !== null) children.push(spPr);
  return elem(c('errBars'), { children });
};

// Which CT_*Ser a series serializes as. It decides the children a series
// may carry and their order; a stock chart's series are CT_LineSer.
type SeriesShape = 'bar' | 'line' | 'area' | 'pie' | 'scatter' | 'bubble' | 'radar' | 'surface';

const SERIES_SHAPE: Readonly<Record<ChartSpec['kind'], SeriesShape>> = {
  bar: 'bar',
  column: 'bar',
  line: 'line',
  stock: 'line',
  area: 'area',
  pie: 'pie',
  doughnut: 'pie',
  scatter: 'scatter',
  bubble: 'bubble',
  radar: 'radar',
  surface: 'surface',
};

// Children each series shape allows beyond the shared head, per dml-chart.xsd.
// Emitting one outside its shape is schema-invalid.
const SHAPES_WITH_MARKER: ReadonlySet<SeriesShape> = new Set(['line', 'scatter', 'radar']);
const SHAPES_WITH_TRENDLINE: ReadonlySet<SeriesShape> = new Set([
  'bar',
  'line',
  'area',
  'scatter',
  'bubble',
]);
const SHAPES_WITH_ERR_BARS = SHAPES_WITH_TRENDLINE;
const SHAPES_WITH_INVERT: ReadonlySet<SeriesShape> = new Set(['bar', 'bubble']);

// Scatter sub-types that draw the connecting line / the point markers.
const SCATTER_STYLES_WITH_LINE: ReadonlySet<string> = new Set([
  'line',
  'lineMarker',
  'smooth',
  'smoothMarker',
]);
const SCATTER_STYLES_WITH_MARKER: ReadonlySet<string> = new Set([
  'marker',
  'lineMarker',
  'smoothMarker',
]);

// The `<c:cat>` channel shared by every series of a category chart.
const categoryChannel = (spec: ChartSpec, formula: string): XmlElement => {
  if (spec.categoryAxisDate !== undefined) {
    return elem(c('cat'), {
      children: [
        numRef(
          formula,
          spec.categories.map(dateSerial),
          spec.categoryAxisNumberFormat ?? 'General',
        ),
      ],
    });
  }
  if (spec.categoryGroupLevels !== undefined && spec.categoryGroupLevels.length > 0) {
    return elem(c('cat'), {
      children: [
        multiLvlStrRef(
          formula,
          [spec.categories, ...spec.categoryGroupLevels],
          spec.categories.length,
        ),
      ],
    });
  }
  return elem(c('cat'), { children: [strRef(formula, spec.categories)] });
};

const padTo = (
  values: ReadonlyArray<number | null> | undefined,
  length: number,
): Array<number | null> => Array.from({ length }, (_, i) => values?.[i] ?? null);

const seriesElement = (
  spec: ChartSpec,
  seriesIdx: number,
  layout: ChartSheetLayout,
): XmlElement => {
  const series = spec.series[seriesIdx];
  const refs = layout.series[seriesIdx];
  if (!series || !refs) throw new Error('seriesElement: out of range');

  const color =
    series.color !== undefined
      ? hexOf(series.color)
      : (DEFAULT_ACCENT_COLORS[seriesIdx % DEFAULT_ACCENT_COLORS.length] ?? '4472C4');

  // A combo series is shaped by its own plot group, not the chart's kind.
  const kind = series.chartKind ?? spec.kind;
  const shape = SERIES_SHAPE[kind];
  const children: XmlElement[] = [
    valNode(c('idx'), seriesIdx),
    valNode(c('order'), seriesIdx),
    elem(c('tx'), { children: [strRef(refs.name, [series.name])] }),
  ];

  // The data channels. Category shapes right-pad `values` to the category
  // count so the points line up with the axis; xy shapes pair their channels
  // point by point against the series' own length.
  const isXy = shape === 'scatter' || shape === 'bubble';
  const pointCount = isXy ? series.values.length : spec.categories.length;
  const channels: XmlElement[] = [];
  if (isXy) {
    if (refs.xValues === undefined) throw new Error('seriesElement: xy series without an x range');
    channels.push(
      elem(c('xVal'), { children: [numRef(refs.xValues, padTo(series.xValues, pointCount))] }),
      elem(c('yVal'), { children: [numRef(refs.values, padTo(series.values, pointCount))] }),
    );
    if (shape === 'bubble') {
      if (refs.bubbleSizes === undefined) {
        throw new Error('seriesElement: bubble series without a size range');
      }
      channels.push(
        elem(c('bubbleSize'), {
          children: [numRef(refs.bubbleSizes, padTo(series.bubbleSizes, pointCount))],
        }),
        valNode(c('bubble3D'), spec.bubble3D === true ? '1' : '0'),
      );
    }
  } else {
    if (layout.categories === undefined) {
      throw new Error('seriesElement: category series without a category range');
    }
    channels.push(
      categoryChannel(spec, layout.categories),
      elem(c('val'), { children: [numRef(refs.values, padTo(series.values, pointCount))] }),
    );
  }

  // A surface is colored by value band, not by series, and CT_SurfaceSer
  // has no child besides its channels — nothing below applies to it.
  if (shape === 'surface') {
    return elem(c('ser'), { children: [...children, ...channels] });
  }

  // spPr shape depends on what the series actually paints. A line series'
  // visible element is its stroke, so the color MUST live on <a:ln>; a bare
  // <a:solidFill> (correct for bar/area fills) leaves the line uncolored and
  // PowerPoint falls back to its automatic series palette — the line then
  // renders in the wrong color. seriesSpPr emits the color on both <a:ln>
  // and <a:solidFill>, so it colors the line for PowerPoint while keeping the
  // solidFill the reader round-trips. Bar / column / pie keep the legacy
  // solid-fill-only shape for tight round-trip compatibility with fixtures.
  const scatterStyle = spec.scatterStyle ?? 'marker';
  // Stock series are invisible carriers for the high-low lines and up/down
  // bars; a scatter series without a line sub-type shows its markers only.
  const lineHidden =
    spec.kind === 'stock' || (shape === 'scatter' && !SCATTER_STYLES_WITH_LINE.has(scatterStyle));
  if (
    lineHidden ||
    shape === 'line' ||
    shape === 'scatter' ||
    shape === 'radar' ||
    series.lineWidthEmu !== undefined ||
    series.lineColor !== undefined ||
    series.lineDash !== undefined
  ) {
    const lineColor = lineHidden
      ? null
      : series.lineColor !== undefined
        ? hexOf(series.lineColor)
        : color;
    children.push(
      seriesSpPr(color, lineColor, series.lineWidthEmu, series.lineDash, series.fillOpacity),
    );
  } else {
    children.push(solidFillSpPr(color, series.fillOpacity));
  }
  // invertIfNegative only exists on CT_BarSer and CT_BubbleSer. Emitting it on
  // a line/pie/area series is schema-invalid, so gate on those shapes.
  if (series.invertIfNegative === true && SHAPES_WITH_INVERT.has(shape)) {
    children.push(valNode(c('invertIfNegative'), '1'));
  }
  // <c:marker> is only valid on the marker-bearing series types
  // (CT_LineSer / CT_ScatterSer / CT_RadarSer). Emitting it on bar/column/
  // pie/doughnut/area produces schema-invalid XML.
  if (SHAPES_WITH_MARKER.has(shape)) {
    const markerColor = series.markerColor !== undefined ? hexOf(series.markerColor) : color;
    const markerLineColor =
      series.markerLineColor !== undefined ? hexOf(series.markerLineColor) : markerColor;
    // Markers are on by default, so the sub-types that have none must say so
    // on every series: stock charts, marker-less scatter styles, and the
    // standard / filled radar.
    const markersOff =
      spec.kind === 'stock' ||
      (shape === 'scatter' && !SCATTER_STYLES_WITH_MARKER.has(scatterStyle)) ||
      (shape === 'radar' && spec.radarStyle !== 'marker');
    children.push(
      markerElement(
        series.markerSymbol ?? (markersOff ? 'none' : undefined),
        series.markerSizePt,
        markerColor,
        markerLineColor,
      ),
    );
  }
  // <c:dPt> overrides go after invertIfNegative / marker per schema.
  for (const dPt of dPtElements(
    series.pointColors,
    series.pointExplosions,
    series.fillOpacity,
    spec.bubble3D === true,
  )) {
    children.push(dPt);
  }
  // Without the series-level show* group, LibreOffice treats the missing
  // toggles as on and paints every label kind on the points that have no
  // override — so per-point labels require the series defaults.
  if (series.pointDataLabels !== undefined && series.dataLabels === undefined) {
    throw new Error(
      `chart: series '${series.name}' sets pointDataLabels without dataLabels; add the series-level dataLabels the overrides fall back to`,
    );
  }
  const serDLbls = buildDLblsFromLabels(series.dataLabels, series.pointDataLabels);
  if (serDLbls !== null) children.push(serDLbls);
  // <c:trendline> exists on CT_BarSer/LineSer/AreaSer/ScatterSer/BubbleSer but
  // NOT on CT_PieSer (pie/doughnut) or CT_RadarSer — emitting it there is
  // schema-invalid, so gate on the trendline-bearing kinds.
  if (series.trendline && SHAPES_WITH_TRENDLINE.has(shape) && spec.kind !== 'stock') {
    children.push(trendlineElement(series.trendline));
  }
  if (series.errorBars !== undefined || series.xErrorBars !== undefined) {
    if (!SHAPES_WITH_ERR_BARS.has(shape)) {
      throw new Error(
        `chart: series '${series.name}' sets error bars, which a ${kind} series cannot carry (bar / column / line / area / scatter / bubble only)`,
      );
    }
    if (series.xErrorBars !== undefined && !isXy) {
      throw new Error(
        `chart: series '${series.name}' sets xErrorBars, which only scatter / bubble series have an x channel for`,
      );
    }
    // CT_ErrBars order is free between the two, PowerPoint writes x then y.
    if (series.xErrorBars !== undefined) children.push(errBarsElement(series.xErrorBars, 'x'));
    if (series.errorBars !== undefined) {
      children.push(errBarsElement(series.errorBars, isXy ? 'y' : null));
    }
  }
  // CT_BubbleSer closes with bubbleSize + bubble3D; nothing follows them.
  children.push(...channels);
  // Line series always get an explicit <c:smooth>: the schema default for an
  // absent element is val="1", so LibreOffice draws an unauthored line as a
  // smooth curve while PowerPoint draws it straight. PowerPoint itself always
  // writes the element; doing the same keeps every renderer straight unless
  // smoothing was asked for. (Only CT_LineSer / CT_ScatterSer carry smooth —
  // emitting it on a bar/pie series would be schema-invalid.)
  if (shape === 'line') {
    children.push(valNode(c('smooth'), series.smooth === true ? '1' : '0'));
  } else if (shape === 'scatter') {
    const smooth = series.smooth ?? (scatterStyle === 'smooth' || scatterStyle === 'smoothMarker');
    children.push(valNode(c('smooth'), smooth ? '1' : '0'));
  }
  return elem(c('ser'), { children });
};

// Axis ids — arbitrary distinct positive 32-bit integers PowerPoint just
// needs them stable within the chart for the `<c:crossAx>` back-pointer.
const CAT_AX_ID = 111111111;
const VAL_AX_ID = 222222222;
// Secondary axis pair for combo charts (series with `secondaryAxis: true`).
const SEC_CAT_AX_ID = 333333333;
const SEC_VAL_AX_ID = 444444444;

interface AxisIdPair {
  readonly cat: number;
  readonly val: number;
}

const PRIMARY_AXES: AxisIdPair = { cat: CAT_AX_ID, val: VAL_AX_ID };
const SECONDARY_AXES: AxisIdPair = { cat: SEC_CAT_AX_ID, val: SEC_VAL_AX_ID };

/** Category kinds that can participate in a combo plot-group split. */
const COMBO_KINDS = new Set<string>(['bar', 'column', 'line', 'area']);

const effectiveSeriesKind = (spec: ChartSpec, seriesIdx: number): ChartSpec['kind'] =>
  spec.series[seriesIdx]?.chartKind ?? spec.kind;

/**
 * Splits the series into plot groups keyed by (effective kind, axis).
 * Primary-axis groups come first so secondary overlays paint on top,
 * and bar groups precede line/area within each axis for the same reason
 * (matching PowerPoint's combo emit order).
 */
const comboPlotGroups = (
  spec: ChartSpec,
): { kind: ChartSpec['kind']; secondary: boolean; indices: number[] }[] => {
  const groups = new Map<
    string,
    { kind: ChartSpec['kind']; secondary: boolean; indices: number[] }
  >();
  for (let i = 0; i < spec.series.length; i++) {
    const kind = effectiveSeriesKind(spec, i);
    const secondary = spec.series[i]?.secondaryAxis === true;
    const key = `${kind}|${secondary ? '1' : '0'}`;
    const group = groups.get(key);
    if (group) group.indices.push(i);
    else groups.set(key, { kind, secondary, indices: [i] });
  }
  const paintOrder = (g: { kind: ChartSpec['kind']; secondary: boolean }): number =>
    (g.secondary ? 2 : 0) + (g.kind === 'line' || g.kind === 'area' ? 1 : 0);
  return [...groups.values()].sort((a, b) => paintOrder(a) - paintOrder(b));
};

// Build a `<c:txPr>` block carrying axis tick-label font / color and an
// optional `<a:bodyPr rot="N"/>` rotation. Returns null when neither
// labelStyle nor rotation are set.
const axisTxPrElement = (
  style: ChartTextStyle | undefined,
  rotationDeg: number | undefined,
): XmlElement | null => {
  if (style === undefined && rotationDeg === undefined) return null;
  const bodyAttrs: ReturnType<typeof attr>[] = [];
  if (rotationDeg !== undefined) {
    bodyAttrs.push(attr(qname('', 'rot', ''), String(Math.round(rotationDeg * 60000))));
  }
  bodyAttrs.push(attr(qname('', 'vert', ''), 'horz'));
  const bodyPr = elem(a('bodyPr'), { attrs: bodyAttrs });
  const { attrs: defAttrs, children: defChildren } = rPrAttrsFromStyle(style);
  const defRPr = elem(a('defRPr'), { attrs: defAttrs, children: defChildren });
  const pPr = elem(a('pPr'), { children: [defRPr] });
  // An empty <a:p> with just pPr is the canonical "defaults only" shape.
  return elem(c('txPr'), {
    children: [bodyPr, elem(a('lstStyle')), elem(a('p'), { children: [pPr] })],
  });
};

const catAxis = (spec: ChartSpec): XmlElement => {
  const catOrientation = spec.categoryAxisOrientation ?? 'minMax';
  const dateAxis = spec.categoryAxisDate;
  // CT_Scaling order: logBase, orientation, max, min. Only a date axis has a
  // numeric range to bound; `validateSpec` rejects scaling on a text axis.
  const scalingChildren: XmlElement[] = [valNode(c('orientation'), catOrientation)];
  if (spec.categoryAxisScaling?.max !== undefined) {
    scalingChildren.push(valNode(c('max'), spec.categoryAxisScaling.max));
  }
  if (spec.categoryAxisScaling?.min !== undefined) {
    scalingChildren.push(valNode(c('min'), spec.categoryAxisScaling.min));
  }
  const children: XmlElement[] = [
    valNode(c('axId'), CAT_AX_ID),
    elem(c('scaling'), { children: scalingChildren }),
    valNode(c('delete'), spec.categoryAxisHidden ? '1' : '0'),
    valNode(c('axPos'), isHorizontalBar(spec) ? 'l' : 'b'),
  ];
  if (spec.categoryAxisMajorGridlines) {
    children.push(
      gridlinesElement(
        'majorGridlines',
        spec.categoryAxisMajorGridlineColor,
        spec.categoryAxisMajorGridlineWidthEmu,
        'chart: categoryAxisMajorGridlineWidthEmu',
      ),
    );
  }
  if (spec.categoryAxisMinorGridlines) {
    children.push(
      gridlinesElement(
        'minorGridlines',
        spec.categoryAxisMinorGridlineColor,
        spec.categoryAxisMinorGridlineWidthEmu,
        'chart: categoryAxisMinorGridlineWidthEmu',
      ),
    );
  }
  if (spec.categoryAxisTitle !== undefined) {
    children.push(
      titleElement(
        spec.categoryAxisTitle,
        spec.categoryAxisTitleStyle,
        spec.categoryAxisTitleRotationDeg,
      ),
    );
  }
  if (spec.categoryAxisNumberFormat !== undefined) {
    children.push(
      elem(c('numFmt'), {
        attrs: [
          attr(qname('', 'formatCode', ''), spec.categoryAxisNumberFormat),
          attr(qname('', 'sourceLinked', ''), '0'),
        ],
      }),
    );
  }
  if (spec.categoryAxisMajorTickMark !== undefined) {
    children.push(valNode(c('majorTickMark'), spec.categoryAxisMajorTickMark));
  }
  if (spec.categoryAxisMinorTickMark !== undefined) {
    children.push(valNode(c('minorTickMark'), spec.categoryAxisMinorTickMark));
  }
  if (spec.categoryAxisTickLabelPos !== undefined) {
    children.push(valNode(c('tickLblPos'), spec.categoryAxisTickLabelPos));
  }
  const catLine = lineSpPr(
    spec.categoryAxisLineColor,
    spec.categoryAxisLineWidthEmu,
    'chart: categoryAxisLineWidthEmu',
    spec.categoryAxisLineHidden === true,
  );
  if (catLine !== null) children.push(catLine);
  const catTxPr = axisTxPrElement(spec.categoryAxisLabelStyle, spec.categoryAxisLabelRotationDeg);
  if (catTxPr) children.push(catTxPr);
  children.push(valNode(c('crossAx'), VAL_AX_ID));
  if (dateAxis !== undefined) {
    // CT_DateAx tail: auto, lblOffset, then the unit pairs. `auto` is off —
    // with it on PowerPoint re-decides between a text and a date axis from
    // the data and may discard the authored units.
    children.push(valNode(c('auto'), '0'));
    if (spec.categoryAxisLabelOffset !== undefined) {
      children.push(valNode(c('lblOffset'), spec.categoryAxisLabelOffset));
    }
    if (dateAxis.baseTimeUnit !== undefined) {
      children.push(valNode(c('baseTimeUnit'), dateAxis.baseTimeUnit));
    }
    if (dateAxis.majorUnit !== undefined)
      children.push(valNode(c('majorUnit'), dateAxis.majorUnit));
    if (dateAxis.majorTimeUnit !== undefined) {
      children.push(valNode(c('majorTimeUnit'), dateAxis.majorTimeUnit));
    }
    if (dateAxis.minorUnit !== undefined)
      children.push(valNode(c('minorUnit'), dateAxis.minorUnit));
    if (dateAxis.minorTimeUnit !== undefined) {
      children.push(valNode(c('minorTimeUnit'), dateAxis.minorTimeUnit));
    }
    return elem(c('dateAx'), { children });
  }
  // CT_CatAx schema order: lblAlgn / lblOffset precede the skip pair.
  if (spec.categoryAxisLabelAlign !== undefined) {
    children.push(valNode(c('lblAlgn'), spec.categoryAxisLabelAlign));
  }
  if (spec.categoryAxisLabelOffset !== undefined) {
    children.push(valNode(c('lblOffset'), spec.categoryAxisLabelOffset));
  }
  if (spec.categoryAxisTickLabelSkip !== undefined) {
    children.push(valNode(c('tickLblSkip'), spec.categoryAxisTickLabelSkip));
  }
  if (spec.categoryAxisTickMarkSkip !== undefined) {
    children.push(valNode(c('tickMarkSkip'), spec.categoryAxisTickMarkSkip));
  }
  if (spec.categoryAxisNoMultiLevelLabel) children.push(valNode(c('noMultiLvlLbl'), '1'));
  return elem(c('catAx'), { children });
};

// The primary and secondary value axes share one element shape; only the
// axis ids, the position, and the crossing differ. Every field is spelled
// out (with `undefined`) so both callers must decide about each one.
interface ValueAxisFields {
  readonly scaling: ChartAxisScaling | undefined;
  readonly orientation: 'minMax' | 'maxMin' | undefined;
  readonly hidden: boolean | undefined;
  readonly majorGridlines: boolean | undefined;
  readonly majorGridlineColor: string | undefined;
  readonly majorGridlineWidthEmu: number | undefined;
  readonly minorGridlines: boolean | undefined;
  readonly minorGridlineColor: string | undefined;
  readonly minorGridlineWidthEmu: number | undefined;
  readonly title: string | undefined;
  readonly titleStyle: ChartTextStyle | undefined;
  readonly titleRotationDeg: number | undefined;
  readonly majorTickMark: 'in' | 'out' | 'cross' | 'none' | undefined;
  readonly minorTickMark: 'in' | 'out' | 'cross' | 'none' | undefined;
  readonly tickLabelPos: 'none' | 'low' | 'high' | 'nextTo' | undefined;
  readonly lineColor: string | undefined;
  readonly lineWidthEmu: number | undefined;
  readonly lineHidden: boolean | undefined;
  /** Spec paths of this axis' width fields, for validation messages. */
  readonly widthFields: {
    readonly majorGridline: string;
    readonly minorGridline: string;
    readonly line: string;
  };
  readonly labelStyle: ChartTextStyle | undefined;
  readonly labelRotationDeg: number | undefined;
  readonly crosses: ChartSpec['valueAxisCrosses'];
  readonly crossBetween: 'between' | 'midCat' | undefined;
}

const valueAxisElement = (
  f: ValueAxisFields,
  ids: { axId: number; crossAxId: number; axPos: 'l' | 'r' | 'b' | 't' },
): XmlElement => {
  // CT_Scaling sequence (dml-chart.xsd): logBase, orientation, max, min, extLst.
  // `max` MUST precede `min` — emitting them in the other order is rejected by
  // the schema ("element max: expected extLst").
  const scalingChildren: XmlElement[] = [];
  if (f.scaling?.logBase !== undefined) {
    scalingChildren.push(valNode(c('logBase'), f.scaling.logBase));
  }
  scalingChildren.push(valNode(c('orientation'), f.orientation ?? 'minMax'));
  if (f.scaling?.max !== undefined) {
    scalingChildren.push(valNode(c('max'), f.scaling.max));
  }
  if (f.scaling?.min !== undefined) {
    scalingChildren.push(valNode(c('min'), f.scaling.min));
  }
  const children: XmlElement[] = [
    valNode(c('axId'), ids.axId),
    elem(c('scaling'), { children: scalingChildren }),
    valNode(c('delete'), f.hidden ? '1' : '0'),
    valNode(c('axPos'), ids.axPos),
  ];
  if (f.majorGridlines) {
    children.push(
      gridlinesElement(
        'majorGridlines',
        f.majorGridlineColor,
        f.majorGridlineWidthEmu,
        `chart: ${f.widthFields.majorGridline}`,
      ),
    );
  }
  if (f.minorGridlines) {
    children.push(
      gridlinesElement(
        'minorGridlines',
        f.minorGridlineColor,
        f.minorGridlineWidthEmu,
        `chart: ${f.widthFields.minorGridline}`,
      ),
    );
  }
  if (f.title !== undefined) {
    children.push(titleElement(f.title, f.titleStyle, f.titleRotationDeg));
  }
  if (f.scaling?.numberFormat !== undefined) {
    children.push(
      elem(c('numFmt'), {
        attrs: [
          attr(qname('', 'formatCode', ''), f.scaling.numberFormat),
          attr(qname('', 'sourceLinked', ''), '0'),
        ],
      }),
    );
  }
  if (f.majorTickMark !== undefined) {
    children.push(valNode(c('majorTickMark'), f.majorTickMark));
  }
  if (f.minorTickMark !== undefined) {
    children.push(valNode(c('minorTickMark'), f.minorTickMark));
  }
  if (f.tickLabelPos !== undefined) {
    children.push(valNode(c('tickLblPos'), f.tickLabelPos));
  }
  const axisLine = lineSpPr(
    f.lineColor,
    f.lineWidthEmu,
    `chart: ${f.widthFields.line}`,
    f.lineHidden === true,
  );
  if (axisLine !== null) children.push(axisLine);
  const valTxPr = axisTxPrElement(f.labelStyle, f.labelRotationDeg);
  if (valTxPr) children.push(valTxPr);
  children.push(valNode(c('crossAx'), ids.crossAxId));
  // <c:crosses val>/`crossesAt val>` — mutually exclusive per the
  // schema. Object form `{ at: N }` → crossesAt; string form → crosses.
  const xross = f.crosses;
  if (xross !== undefined) {
    if (typeof xross === 'string') {
      children.push(valNode(c('crosses'), xross));
    } else {
      children.push(valNode(c('crossesAt'), String(xross.at)));
    }
  }
  if (f.crossBetween !== undefined) {
    children.push(valNode(c('crossBetween'), f.crossBetween));
  }
  if (f.scaling?.majorUnit !== undefined) {
    children.push(valNode(c('majorUnit'), f.scaling.majorUnit));
  }
  if (f.scaling?.minorUnit !== undefined) {
    children.push(valNode(c('minorUnit'), f.scaling.minorUnit));
  }
  if (f.scaling?.displayUnits !== undefined) {
    children.push(
      elem(c('dispUnits'), {
        children: [
          valNode(c('builtInUnit'), f.scaling.displayUnits),
          // An empty label element is "show the caption, positioned and
          // worded by the application".
          ...(f.scaling.displayUnitsLabel === true ? [elem(c('dispUnitsLbl'))] : []),
        ],
      }),
    );
  }
  return elem(c('valAx'), { children });
};

// Horizontal bars swap the axis positions: categories run down the left
// edge and values along the bottom. PowerPoint and PptxGenJS both write
// `catAx axPos="l"` / `valAx axPos="b"` for `barDir="bar"`.
const isHorizontalBar = (spec: ChartSpec): boolean => spec.kind === 'bar';

const valAxis = (spec: ChartSpec): XmlElement =>
  valueAxisElement(
    {
      scaling: spec.valueAxis,
      orientation: spec.valueAxisOrientation,
      hidden: spec.valueAxisHidden,
      majorGridlines: spec.valueAxisMajorGridlines,
      majorGridlineColor: spec.valueAxisMajorGridlineColor,
      majorGridlineWidthEmu: spec.valueAxisMajorGridlineWidthEmu,
      minorGridlines: spec.valueAxisMinorGridlines,
      minorGridlineColor: spec.valueAxisMinorGridlineColor,
      minorGridlineWidthEmu: spec.valueAxisMinorGridlineWidthEmu,
      title: spec.valueAxisTitle,
      titleStyle: spec.valueAxisTitleStyle,
      titleRotationDeg: spec.valueAxisTitleRotationDeg,
      majorTickMark: spec.valueAxisMajorTickMark,
      minorTickMark: spec.valueAxisMinorTickMark,
      tickLabelPos: spec.valueAxisTickLabelPos,
      lineColor: spec.valueAxisLineColor,
      lineWidthEmu: spec.valueAxisLineWidthEmu,
      lineHidden: spec.valueAxisLineHidden,
      widthFields: {
        majorGridline: 'valueAxisMajorGridlineWidthEmu',
        minorGridline: 'valueAxisMinorGridlineWidthEmu',
        line: 'valueAxisLineWidthEmu',
      },
      labelStyle: spec.valueAxisLabelStyle,
      labelRotationDeg: spec.valueAxisLabelRotationDeg,
      crosses: spec.valueAxisCrosses,
      crossBetween: spec.valueAxisCrossBetween,
    },
    { axId: VAL_AX_ID, crossAxId: CAT_AX_ID, axPos: isHorizontalBar(spec) ? 'b' : 'l' },
  );

/**
 * The horizontal axis of a scatter / bubble chart. Both of their axes are
 * `<c:valAx>`; this one takes the spec's `categoryAxis*` fields, so an
 * author formats "the axis along the bottom" the same way for every kind.
 * It keeps the category axis' id, which the y axis already crosses.
 */
const xValAxis = (spec: ChartSpec): XmlElement =>
  valueAxisElement(
    {
      scaling: {
        ...spec.categoryAxisScaling,
        ...(spec.categoryAxisNumberFormat !== undefined
          ? { numberFormat: spec.categoryAxisNumberFormat }
          : {}),
      },
      orientation: spec.categoryAxisOrientation,
      hidden: spec.categoryAxisHidden,
      majorGridlines: spec.categoryAxisMajorGridlines,
      majorGridlineColor: spec.categoryAxisMajorGridlineColor,
      majorGridlineWidthEmu: spec.categoryAxisMajorGridlineWidthEmu,
      minorGridlines: spec.categoryAxisMinorGridlines,
      minorGridlineColor: spec.categoryAxisMinorGridlineColor,
      minorGridlineWidthEmu: spec.categoryAxisMinorGridlineWidthEmu,
      title: spec.categoryAxisTitle,
      titleStyle: spec.categoryAxisTitleStyle,
      titleRotationDeg: spec.categoryAxisTitleRotationDeg,
      majorTickMark: spec.categoryAxisMajorTickMark,
      minorTickMark: spec.categoryAxisMinorTickMark,
      tickLabelPos: spec.categoryAxisTickLabelPos,
      lineColor: spec.categoryAxisLineColor,
      lineWidthEmu: spec.categoryAxisLineWidthEmu,
      lineHidden: spec.categoryAxisLineHidden,
      widthFields: {
        majorGridline: 'categoryAxisMajorGridlineWidthEmu',
        minorGridline: 'categoryAxisMinorGridlineWidthEmu',
        line: 'categoryAxisLineWidthEmu',
      },
      labelStyle: spec.categoryAxisLabelStyle,
      labelRotationDeg: spec.categoryAxisLabelRotationDeg,
      crosses: undefined,
      // Points sit on the tick marks of a numeric axis, not between them.
      crossBetween: 'midCat',
    },
    { axId: CAT_AX_ID, crossAxId: VAL_AX_ID, axPos: 'b' },
  );

// Series (depth) axis id — see the axis-id note above.
const SER_AX_ID = 555555555;

/** `<c:serAx>` — the depth axis of a 3-D or surface chart; it crosses the value axis. */
const serAxis = (spec: ChartSpec): XmlElement => {
  const s = spec.seriesAxis ?? {};
  const children: XmlElement[] = [
    valNode(c('axId'), SER_AX_ID),
    elem(c('scaling'), { children: [valNode(c('orientation'), s.orientation ?? 'minMax')] }),
    valNode(c('delete'), s.hidden ? '1' : '0'),
    valNode(c('axPos'), 'b'),
  ];
  if (s.majorGridlines) children.push(elem(c('majorGridlines')));
  if (s.title !== undefined) children.push(titleElement(s.title, s.titleStyle));
  if (s.tickLabelPos !== undefined) children.push(valNode(c('tickLblPos'), s.tickLabelPos));
  const line = lineSpPr(s.lineColor, undefined, 'chart: seriesAxis');
  if (line !== null) children.push(line);
  const txPr = axisTxPrElement(s.labelStyle, undefined);
  if (txPr !== null) children.push(txPr);
  children.push(valNode(c('crossAx'), VAL_AX_ID));
  if (s.tickLabelSkip !== undefined) children.push(valNode(c('tickLblSkip'), s.tickLabelSkip));
  return elem(c('serAx'), { children });
};

// Build `<c:dLbls>` from a ChartDataLabels (showVal / showCatName /
// showSerName / showPercent toggles plus optional numFmt, position,
// separator). Returns `null` when no dataLabels were authored so
// callers know to skip the element entirely.
const buildDLblsFromLabels = (
  dl: ChartDataLabels | undefined,
  pointLabels: ReadonlyArray<ChartDataLabels | null> | undefined = undefined,
): XmlElement | null => {
  // CT_DLbls: the per-point <c:dLbl> overrides precede the series-level
  // group, and a <c:dLbls> holding only overrides is schema-valid.
  const children: XmlElement[] = [];
  pointLabels?.forEach((pl, i) => {
    if (pl === null) return;
    children.push(
      elem(c('dLbl'), {
        children: [
          valNode(c('idx'), i),
          // CT_DLbl only: a literal <c:tx> sits between idx and the shared group.
          ...(pl.text !== undefined ? [elem(c('tx'), { children: [richText(pl.text)] })] : []),
          ...dLblGroupChildren(pl),
        ],
      }),
    );
  });
  if (dl) {
    children.push(...dLblGroupChildren(dl));
    // Group_DLbls only: showLeaderLines follows separator in schema order.
    if (dl.showLeaderLines !== undefined) {
      children.push(valNode(c('showLeaderLines'), dl.showLeaderLines ? '1' : '0'));
    }
  }
  return children.length === 0 ? null : elem(c('dLbls'), { children });
};

// The shared tail of CT_DLbl / CT_DLbls (Group_DLbl / Group_DLbls): numFmt,
// txPr, dLblPos, the show* toggles, separator — in schema order.
const dLblGroupChildren = (dl: ChartDataLabels): XmlElement[] => {
  const children: XmlElement[] = [];
  if (dl.numberFormat !== undefined) {
    children.push(
      elem(c('numFmt'), {
        attrs: [
          attr(qname('', 'formatCode', ''), dl.numberFormat),
          attr(qname('', 'sourceLinked', ''), '0'),
        ],
      }),
    );
  }
  if (dl.fillColor !== undefined) children.push(solidFillSpPr(dl.fillColor));
  if (dl.textStyle !== undefined) {
    // CT_DLbls schema order places <c:txPr> before <c:dLblPos>; reusing
    // `axisTxPrElement` keeps the formatting parity with axis / legend.
    const txPr = axisTxPrElement(dl.textStyle, undefined);
    if (txPr !== null) children.push(txPr);
  }
  if (dl.position !== undefined) children.push(valNode(c('dLblPos'), dl.position));
  children.push(
    valNode(c('showLegendKey'), dl.showLegendKey ? '1' : '0'),
    valNode(c('showVal'), dl.showValue ? '1' : '0'),
    valNode(c('showCatName'), dl.showCategory ? '1' : '0'),
    valNode(c('showSerName'), dl.showSeriesName ? '1' : '0'),
    valNode(c('showPercent'), dl.showPercent ? '1' : '0'),
    valNode(c('showBubbleSize'), dl.showBubbleSize ? '1' : '0'),
  );
  if (dl.separator !== undefined) {
    children.push(elem(c('separator'), { children: [text(dl.separator)] }));
  }
  return children;
};

const dLblsElement = (spec: ChartSpec): XmlElement | null => buildDLblsFromLabels(spec.dataLabels);

// Plot groups ---------------------------------------------------------------

/** What every plot-group builder needs: the chart, its sheet, its series, its axes. */
interface PlotGroup {
  readonly spec: ChartSpec;
  readonly layout: ChartSheetLayout;
  readonly seriesIndices: ReadonlyArray<number>;
  readonly axes: AxisIdPair;
}

const groupSeries = (g: PlotGroup): XmlElement[] =>
  g.seriesIndices.map((i) => seriesElement(g.spec, i, g.layout));

const groupDLbls = (g: PlotGroup): XmlElement[] => {
  const dl = dLblsElement(g.spec);
  return dl ? [dl] : [];
};

const isThreeD = (spec: ChartSpec): boolean => spec.view3D !== undefined;

/**
 * Whether the chart carries a series (depth) axis. `<c:line3DChart>` and
 * the surface elements require three axes; 3-D bar / area charts only lay
 * their series out in depth — and so only have the axis — under the
 * `standard` grouping.
 */
const hasSeriesAxis = (spec: ChartSpec): boolean => {
  if (spec.kind === 'surface') return true;
  if (!isThreeD(spec)) return false;
  if (spec.kind === 'line') return true;
  return (
    (spec.kind === 'bar' || spec.kind === 'column' || spec.kind === 'area') &&
    spec.grouping === 'standard'
  );
};

const axIdNodes = (g: PlotGroup): XmlElement[] => [
  valNode(c('axId'), g.axes.cat),
  valNode(c('axId'), g.axes.val),
  ...(hasSeriesAxis(g.spec) ? [valNode(c('axId'), SER_AX_ID)] : []),
];

const gapDepthNodes = (spec: ChartSpec): XmlElement[] =>
  spec.gapDepthPct !== undefined
    ? [valNode(c('gapDepth'), gapAmountPercent(spec.gapDepthPct, 'chart: gapDepthPct'))]
    : [];

const buildBarChart = (g: PlotGroup, direction: 'col' | 'bar'): XmlElement => {
  const { spec } = g;
  const grouping = spec.grouping ?? 'clustered';
  const children: XmlElement[] = [
    valNode(c('barDir'), direction),
    valNode(c('grouping'), grouping),
    valNode(c('varyColors'), spec.varyColors ? '1' : '0'),
    ...groupSeries(g),
    ...groupDLbls(g),
  ];
  if (spec.gapWidthPct !== undefined) {
    children.push(valNode(c('gapWidth'), gapAmountPercent(spec.gapWidthPct, 'chart: gapWidthPct')));
  }
  if (isThreeD(spec)) {
    // CT_Bar3DChart swaps <c:overlap> for <c:gapDepth> + <c:shape>.
    children.push(...gapDepthNodes(spec));
    if (spec.bar3DShape !== undefined) children.push(valNode(c('shape'), spec.bar3DShape));
    children.push(...axIdNodes(g));
    return elem(c('bar3DChart'), { children });
  }
  // Stacked / 100%-stacked bars must overlap fully (overlap=100), otherwise
  // PowerPoint draws each series in its own sub-slot and the "stack" spreads
  // sideways across the category. PowerPoint always writes overlap=100 for
  // these groupings; default to it when the caller didn't set an explicit
  // overlap. Clustered keeps PowerPoint's own default (no element emitted).
  const overlapPct =
    spec.overlapPct ?? (grouping === 'stacked' || grouping === 'percentStacked' ? 100 : undefined);
  if (overlapPct !== undefined) {
    children.push(valNode(c('overlap'), overlapPercent(overlapPct, 'chart: overlapPct')));
  }
  children.push(...axIdNodes(g));
  return elem(c('barChart'), { children });
};

// CT_Grouping (line / area) has no `clustered`, unlike CT_BarGrouping; a
// combo chart's bar grouping maps to the plain `standard` for those groups.
const lineAreaGrouping = (spec: ChartSpec): 'standard' | 'stacked' | 'percentStacked' =>
  spec.grouping === 'stacked' || spec.grouping === 'percentStacked' ? spec.grouping : 'standard';

// `<c:upDownBars>`; an empty <c:upBars/> / <c:downBars/> takes the
// application's automatic fill (white up, black down in PowerPoint).
const upDownBarsElement = (bars: NonNullable<ChartSpec['upDownBars']>): XmlElement => {
  const bar = (name: 'upBars' | 'downBars', color: string | undefined): XmlElement =>
    elem(c(name), { children: color !== undefined ? [solidFillSpPr(color)] : [] });
  return elem(c('upDownBars'), {
    children: [
      ...(bars.gapWidthPct !== undefined
        ? [
            valNode(
              c('gapWidth'),
              gapAmountPercent(bars.gapWidthPct, 'chart: upDownBars.gapWidthPct'),
            ),
          ]
        : []),
      bar('upBars', bars.upColor),
      bar('downBars', bars.downColor),
    ],
  });
};

const buildLineChart = (g: PlotGroup): XmlElement => {
  const { spec } = g;
  const children: XmlElement[] = [
    valNode(c('grouping'), lineAreaGrouping(spec)),
    valNode(c('varyColors'), spec.varyColors ? '1' : '0'),
    ...groupSeries(g),
    ...groupDLbls(g),
  ];
  if (spec.dropLines) children.push(elem(c('dropLines')));
  if (isThreeD(spec)) {
    // CT_Line3DChart ends the shared head here: no high-low lines, up/down
    // bars or markers in 3-D.
    if (spec.hiLowLines || spec.upDownBars !== undefined) {
      throw new Error('chart: hiLowLines / upDownBars are not available on a 3-D line chart');
    }
    children.push(...gapDepthNodes(spec), ...axIdNodes(g));
    return elem(c('line3DChart'), { children });
  }
  if (spec.hiLowLines) children.push(elem(c('hiLowLines')));
  if (spec.upDownBars !== undefined) children.push(upDownBarsElement(spec.upDownBars));
  // <c:marker val> selects the line subtype: "1" → Line with Markers,
  // "0" → plain Line. Default on, preserving the historical output;
  // authors opt out of markers with `lineMarkers: false`.
  children.push(valNode(c('marker'), spec.lineMarkers === false ? '0' : '1'), ...axIdNodes(g));
  return elem(c('lineChart'), { children });
};

const buildAreaChart = (g: PlotGroup): XmlElement => {
  const { spec } = g;
  const children: XmlElement[] = [
    valNode(c('grouping'), lineAreaGrouping(spec)),
    valNode(c('varyColors'), spec.varyColors ? '1' : '0'),
    ...groupSeries(g),
    ...groupDLbls(g),
  ];
  // Drop lines belong to the group they were authored for: in a combo chart
  // the line group already carries them.
  if (spec.dropLines && spec.kind === 'area') children.push(elem(c('dropLines')));
  if (isThreeD(spec)) {
    children.push(...gapDepthNodes(spec), ...axIdNodes(g));
    return elem(c('area3DChart'), { children });
  }
  children.push(...axIdNodes(g));
  return elem(c('areaChart'), { children });
};

const requireSingleSeries = (spec: ChartSpec): void => {
  if (spec.series.length !== 1) {
    throw new Error(`${spec.kind} chart requires exactly one series`);
  }
};

const firstSliceAngNodes = (spec: ChartSpec): XmlElement[] =>
  spec.firstSliceAngleDeg !== undefined
    ? [
        valNode(
          c('firstSliceAng'),
          firstSliceAngle(spec.firstSliceAngleDeg, 'chart: firstSliceAngleDeg'),
        ),
      ]
    : [];

const buildOfPieChart = (g: PlotGroup, ofPie: NonNullable<ChartSpec['ofPie']>): XmlElement => {
  const children: XmlElement[] = [
    valNode(c('ofPieType'), ofPie.type),
    valNode(c('varyColors'), '1'),
    ...groupSeries(g),
    ...groupDLbls(g),
  ];
  if (ofPie.gapWidthPct !== undefined) {
    children.push(
      valNode(c('gapWidth'), gapAmountPercent(ofPie.gapWidthPct, 'chart: ofPie.gapWidthPct')),
    );
  }
  if (ofPie.customSplit !== undefined && ofPie.splitType !== 'cust') {
    throw new Error("chart: ofPie.customSplit requires ofPie.splitType 'cust'");
  }
  if (ofPie.splitType !== undefined) children.push(valNode(c('splitType'), ofPie.splitType));
  if (ofPie.splitPos !== undefined) children.push(valNode(c('splitPos'), ofPie.splitPos));
  if (ofPie.customSplit !== undefined) {
    children.push(
      elem(c('custSplit'), {
        children: ofPie.customSplit.map((idx) => valNode(c('secondPiePt'), idx)),
      }),
    );
  }
  if (ofPie.secondPieSizePct !== undefined) {
    children.push(
      valNode(
        c('secondPieSize'),
        boundedInt(ofPie.secondPieSizePct, 'secondPieSize', 'chart: ofPie.secondPieSizePct'),
      ),
    );
  }
  if (ofPie.seriesLines) children.push(elem(c('serLines')));
  return elem(c('ofPieChart'), { children });
};

const buildPieChart = (g: PlotGroup): XmlElement => {
  const { spec } = g;
  requireSingleSeries(spec);
  if (spec.ofPie !== undefined) {
    if (isThreeD(spec)) throw new Error('chart: ofPie and view3D cannot be combined');
    return buildOfPieChart(g, spec.ofPie);
  }
  const children: XmlElement[] = [
    valNode(c('varyColors'), '1'),
    ...groupSeries(g),
    ...groupDLbls(g),
  ];
  if (isThreeD(spec)) {
    // CT_Pie3DChart has no <c:firstSliceAng>: a 3-D pie turns with the camera.
    if (spec.firstSliceAngleDeg !== undefined) {
      throw new Error('chart: a 3-D pie has no firstSliceAngleDeg; rotate it with view3D.rotY');
    }
    return elem(c('pie3DChart'), { children });
  }
  children.push(...firstSliceAngNodes(spec));
  return elem(c('pieChart'), { children });
};

const buildDoughnutChart = (g: PlotGroup): XmlElement => {
  const { spec } = g;
  requireSingleSeries(spec);
  return elem(c('doughnutChart'), {
    children: [
      valNode(c('varyColors'), '1'),
      ...groupSeries(g),
      ...groupDLbls(g),
      ...firstSliceAngNodes(spec),
      valNode(c('holeSize'), holeSizePercent(spec.holeSizePct ?? 50, 'chart: holeSizePct')),
    ],
  });
};

const buildScatterChart = (g: PlotGroup): XmlElement =>
  elem(c('scatterChart'), {
    children: [
      // The token is written as authored so it reads back unchanged, but
      // PowerPoint draws from the per-series line / marker properties —
      // `seriesElement` sets those to match.
      valNode(c('scatterStyle'), g.spec.scatterStyle ?? 'marker'),
      valNode(c('varyColors'), g.spec.varyColors ? '1' : '0'),
      ...groupSeries(g),
      ...groupDLbls(g),
      ...axIdNodes(g),
    ],
  });

const buildBubbleChart = (g: PlotGroup): XmlElement => {
  const { spec } = g;
  const children: XmlElement[] = [
    valNode(c('varyColors'), spec.varyColors ? '1' : '0'),
    ...groupSeries(g),
    ...groupDLbls(g),
  ];
  // `spec.bubble3D` is written per series only. CT_BubbleChart allows a
  // chart-level `<c:bubble3D>`, but PowerPoint (16.x, macOS) offers to repair
  // any deck that has one, whatever its value.
  if (spec.bubbleScale !== undefined) {
    children.push(
      valNode(c('bubbleScale'), boundedInt(spec.bubbleScale, 'bubbleScale', 'chart: bubbleScale')),
    );
  }
  if (spec.showNegativeBubbles !== undefined) {
    children.push(valNode(c('showNegBubbles'), spec.showNegativeBubbles ? '1' : '0'));
  }
  if (spec.bubbleSizeRepresents !== undefined) {
    // ST_SizeRepresents spells width as `w`.
    children.push(
      valNode(c('sizeRepresents'), spec.bubbleSizeRepresents === 'width' ? 'w' : 'area'),
    );
  }
  children.push(...axIdNodes(g));
  return elem(c('bubbleChart'), { children });
};

const buildRadarChart = (g: PlotGroup): XmlElement =>
  elem(c('radarChart'), {
    children: [
      valNode(c('radarStyle'), g.spec.radarStyle ?? 'standard'),
      valNode(c('varyColors'), g.spec.varyColors ? '1' : '0'),
      ...groupSeries(g),
      ...groupDLbls(g),
      ...axIdNodes(g),
    ],
  });

const buildStockChart = (g: PlotGroup): XmlElement => {
  const { spec } = g;
  const count = spec.series.length;
  if (count !== STOCK_SERIES_WITHOUT_OPEN && count !== STOCK_SERIES_WITH_OPEN) {
    throw new Error(
      `stock chart requires 3 series (high, low, close) or 4 (open, high, low, close); got ${count}`,
    );
  }
  const children: XmlElement[] = [...groupSeries(g), ...groupDLbls(g)];
  if (spec.dropLines) children.push(elem(c('dropLines')));
  // The series themselves are invisible; the high-low lines are the chart.
  // Opt out with `hiLowLines: false`.
  if (spec.hiLowLines !== false) children.push(elem(c('hiLowLines')));
  // Up/down bars span open → close, so they need the open series; with one
  // they are the candlestick bodies and on by default.
  if (spec.upDownBars !== undefined && count !== STOCK_SERIES_WITH_OPEN) {
    throw new Error('stock chart: upDownBars require the 4-series form (open, high, low, close)');
  }
  if (count === STOCK_SERIES_WITH_OPEN) children.push(upDownBarsElement(spec.upDownBars ?? {}));
  children.push(...axIdNodes(g));
  return elem(c('stockChart'), { children });
};

const buildSurfaceChart = (g: PlotGroup): XmlElement => {
  const { spec } = g;
  const children: XmlElement[] = [];
  if (spec.surfaceWireframe !== undefined) {
    children.push(valNode(c('wireframe'), spec.surfaceWireframe ? '1' : '0'));
  }
  children.push(...groupSeries(g), ...axIdNodes(g));
  return elem(c(spec.surfaceContour === true ? 'surfaceChart' : 'surface3DChart'), { children });
};

/** One plot group, dispatched by its effective kind. */
const buildPlotGroup = (g: PlotGroup, kind: ChartSpec['kind']): XmlElement => {
  switch (kind) {
    case 'column':
      return buildBarChart(g, 'col');
    case 'bar':
      return buildBarChart(g, 'bar');
    case 'line':
      return buildLineChart(g);
    case 'area':
      return buildAreaChart(g);
    case 'pie':
      return buildPieChart(g);
    case 'doughnut':
      return buildDoughnutChart(g);
    case 'scatter':
      return buildScatterChart(g);
    case 'bubble':
      return buildBubbleChart(g);
    case 'radar':
      return buildRadarChart(g);
    case 'stock':
      return buildStockChart(g);
    case 'surface':
      return buildSurfaceChart(g);
    default: {
      const exhaustive: never = kind;
      throw new Error(`unsupported chart kind: ${String(exhaustive)}`);
    }
  }
};

/**
 * Secondary value axis (`axPos="r"`, crossing at the category maximum) —
 * the right-hand axis PowerPoint pairs with `secondaryAxis` series.
 * Formatting comes from `spec.secondaryValueAxis`; PowerPoint's defaults
 * apply for everything left out.
 */
const secondaryValAxis = (spec: ChartSpec): XmlElement => {
  const s = spec.secondaryValueAxis ?? {};
  return valueAxisElement(
    {
      scaling: s.scaling,
      orientation: undefined,
      hidden: undefined,
      majorGridlines: s.majorGridlines,
      majorGridlineColor: s.majorGridlineColor,
      majorGridlineWidthEmu: s.majorGridlineWidthEmu,
      minorGridlines: undefined,
      minorGridlineColor: undefined,
      minorGridlineWidthEmu: undefined,
      title: s.title,
      titleStyle: s.titleStyle,
      titleRotationDeg: undefined,
      majorTickMark: s.majorTickMark,
      minorTickMark: s.minorTickMark,
      tickLabelPos: s.tickLabelPos,
      lineColor: s.lineColor,
      lineWidthEmu: s.lineWidthEmu,
      lineHidden: undefined,
      widthFields: {
        majorGridline: 'secondaryValueAxis.majorGridlineWidthEmu',
        minorGridline: 'secondaryValueAxis.minorGridlineWidthEmu',
        line: 'secondaryValueAxis.lineWidthEmu',
      },
      labelStyle: s.labelStyle,
      labelRotationDeg: undefined,
      crosses: 'max',
      crossBetween: s.crossBetween,
    },
    { axId: SEC_VAL_AX_ID, crossAxId: SEC_CAT_AX_ID, axPos: 'r' },
  );
};

/**
 * Deleted companion category axis for the secondary pair. PowerPoint
 * requires every plot group's axId pair to resolve to a cat+val pair,
 * so the secondary group gets its own (hidden) category axis.
 */
const secondaryCatAxis = (spec: ChartSpec): XmlElement =>
  elem(c('catAx'), {
    children: [
      valNode(c('axId'), SEC_CAT_AX_ID),
      elem(c('scaling'), { children: [valNode(c('orientation'), 'minMax')] }),
      valNode(c('delete'), '1'),
      valNode(c('axPos'), isHorizontalBar(spec) ? 'l' : 'b'),
      valNode(c('crossAx'), SEC_VAL_AX_ID),
    ],
  });

// Builds an <a:rPr ...><a:solidFill><a:srgbClr/></a:solidFill></a:rPr>
// payload from a ChartTextStyle. Returns the rPr children to splice
// into the parent run / def-run-properties node.
const rPrAttrsFromStyle = (
  style: ChartTextStyle | undefined,
): {
  attrs: ReturnType<typeof attr>[];
  children: XmlElement[];
} => {
  const attrs: ReturnType<typeof attr>[] = [];
  const children: XmlElement[] = [];
  if (style?.sizePt !== undefined) {
    attrs.push(attr(qname('', 'sz', ''), String(Math.round(style.sizePt * 100))));
  }
  if (style?.bold === true) attrs.push(attr(qname('', 'b', ''), '1'));
  if (style?.bold === false) attrs.push(attr(qname('', 'b', ''), '0'));
  if (style?.italic === true) attrs.push(attr(qname('', 'i', ''), '1'));
  if (style?.italic === false) attrs.push(attr(qname('', 'i', ''), '0'));
  if (style?.color !== undefined) {
    const hex = style.color.startsWith('#') ? style.color.slice(1) : style.color;
    children.push(
      elem(a('solidFill'), {
        children: [elem(a('srgbClr'), { attrs: [attr(qname('', 'val', ''), hex.toUpperCase())] })],
      }),
    );
  }
  // <a:latin> / <a:ea> / <a:cs> follow the fill group in
  // CT_TextCharacterProperties, in that order; ea carries the same face as
  // latin so CJK glyphs aren't dropped to the renderer's latin-only fallback
  // (the latin slot is ignored for east-asian script). cs stays separate:
  // `font` alone must not change how complex scripts render.
  if (style?.font !== undefined) {
    const typeface = attr(qname('', 'typeface', ''), style.font);
    children.push(elem(a('latin'), { attrs: [typeface] }));
    children.push(elem(a('ea'), { attrs: [typeface] }));
  }
  if (style?.fontComplexScript !== undefined) {
    children.push(
      elem(a('cs'), { attrs: [attr(qname('', 'typeface', ''), style.fontComplexScript)] }),
    );
  }
  return { attrs, children };
};

// `<c:layout><c:manualLayout>`. x / y are measured from the chart's top-left
// corner (`edge`); w / h stay in the default `factor` mode, where they are a
// share of the chart size — under `edge` they would be the right / bottom
// coordinates instead. `size` is absent for a title, which sizes to its text.
const manualLayoutElement = (
  position: { readonly x: number; readonly y: number },
  size?: { readonly w: number; readonly h: number },
  target?: ChartManualLayout['target'],
): XmlElement =>
  elem(c('layout'), {
    children: [
      elem(c('manualLayout'), {
        children: [
          ...(target !== undefined ? [valNode(c('layoutTarget'), target)] : []),
          valNode(c('xMode'), 'edge'),
          valNode(c('yMode'), 'edge'),
          valNode(c('x'), position.x),
          valNode(c('y'), position.y),
          ...(size !== undefined ? [valNode(c('w'), size.w), valNode(c('h'), size.h)] : []),
        ],
      }),
    ],
  });

// A one-run `<c:rich>` body with no formatting of its own.
const richText = (value: string): XmlElement =>
  elem(c('rich'), {
    children: [
      elem(a('bodyPr')),
      elem(a('lstStyle')),
      elem(a('p'), {
        children: [elem(a('r'), { children: [elem(a('t'), { children: [text(value)] })] })],
      }),
    ],
  });

const titleElement = (
  title: string,
  style?: ChartTextStyle,
  rotationDeg?: number,
  placement?: { readonly overlay?: boolean; readonly layout?: { x: number; y: number } },
): XmlElement => {
  // The paragraph defaults stay empty so an unset size falls back to the
  // application default; the run-level <a:rPr> carries the authored style.
  //
  // No `lang`: it names the language of the run's text, and hard-coding
  // `en-US` on a title we did not author the words of is a claim we cannot
  // make. The chart-level language tag is `ChartSpec.language` (`<c:lang>`).
  const pPr = elem(a('pPr'), { children: [elem(a('defRPr'))] });
  const { attrs: runAttrs, children: runChildren } = rPrAttrsFromStyle(style);
  const runRPr = elem(a('rPr'), { attrs: runAttrs, children: runChildren });
  const tRun = elem(a('r'), {
    children: [runRPr, elem(a('t'), { children: [text(title)] })],
  });
  const para = elem(a('p'), { children: [pPr, tRun] });
  // `rot` and `vert` are only written when authored: PowerPoint reads a
  // `vert="horz"` without `rot` as "not rotated" and lays a value-axis title
  // out horizontally instead of applying its vertical default.
  const rich = elem(c('rich'), {
    children: [
      elem(a('bodyPr'), {
        attrs: [
          ...(rotationDeg !== undefined
            ? [attr(qname('', 'rot', ''), String(Math.round(rotationDeg * 60000)))]
            : []),
          attr(qname('', 'spcFirstLastPara', ''), '1'),
          attr(qname('', 'vertOverflow', ''), 'ellipsis'),
          ...(rotationDeg !== undefined ? [attr(qname('', 'vert', ''), 'horz')] : []),
          attr(qname('', 'wrap', ''), 'square'),
          attr(qname('', 'anchor', ''), 'ctr'),
          attr(qname('', 'anchorCtr', ''), '1'),
        ],
      }),
      elem(a('lstStyle')),
      para,
    ],
  });
  // CT_Title order: tx, layout, overlay.
  return elem(c('title'), {
    children: [
      elem(c('tx'), { children: [rich] }),
      ...(placement?.layout !== undefined ? [manualLayoutElement(placement.layout)] : []),
      valNode(c('overlay'), placement?.overlay === true ? '1' : '0'),
    ],
  });
};

// Camera defaults PowerPoint itself writes for each 3-D chart family. They
// are spelled out so the view never depends on what a consuming application
// falls back to for an empty <c:view3D/>.
const view3DDefaults = (spec: ChartSpec): ChartView3D => {
  if (spec.kind === 'surface' && spec.surfaceContour === true) {
    // A contour plot is the surface seen from straight above.
    return { rotX: 90, rotY: 0, rightAngleAxes: false, perspective: 0 };
  }
  if (spec.kind === 'pie') return { rotX: 30, rotY: 0, rightAngleAxes: false };
  if (spec.kind === 'surface')
    return { rotX: 15, rotY: 20, rightAngleAxes: false, perspective: 30 };
  return { rotX: 15, rotY: 20, rightAngleAxes: true };
};

// `<c:view3D>` in CT_View3D order: rotX, hPercent, rotY, depthPercent, rAngAx, perspective.
const view3DElement = (spec: ChartSpec): XmlElement => {
  const v: ChartView3D = { ...view3DDefaults(spec), ...spec.view3D };
  const children: XmlElement[] = [];
  if (v.rotX !== undefined) {
    children.push(valNode(c('rotX'), boundedInt(v.rotX, 'rotX', 'chart: view3D.rotX')));
  }
  if (v.heightPercent !== undefined) {
    children.push(
      valNode(
        c('hPercent'),
        boundedInt(v.heightPercent, 'heightPercent', 'chart: view3D.heightPercent'),
      ),
    );
  }
  if (v.rotY !== undefined) {
    children.push(valNode(c('rotY'), boundedInt(v.rotY, 'rotY', 'chart: view3D.rotY')));
  }
  if (v.depthPercent !== undefined) {
    children.push(
      valNode(
        c('depthPercent'),
        boundedInt(v.depthPercent, 'depthPercent', 'chart: view3D.depthPercent'),
      ),
    );
  }
  if (v.rightAngleAxes !== undefined) {
    children.push(valNode(c('rAngAx'), v.rightAngleAxes ? '1' : '0'));
  }
  if (v.perspective !== undefined) {
    children.push(
      valNode(
        c('perspective'),
        boundedInt(v.perspective, 'perspective', 'chart: view3D.perspective'),
      ),
    );
  }
  return elem(c('view3D'), { children });
};

const KINDS_WITH_VIEW_3D: ReadonlySet<ChartSpec['kind']> = new Set([
  'bar',
  'column',
  'line',
  'area',
  'pie',
  'surface',
]);
const AXISLESS_KINDS: ReadonlySet<ChartSpec['kind']> = new Set(['pie', 'doughnut']);
const XY_KINDS: ReadonlySet<ChartSpec['kind']> = new Set(['scatter', 'bubble']);

// Cross-field rules of a spec that no single element builder owns. Each one
// would otherwise serialize into a chart PowerPoint repairs or misdraws.
// `ChartSpec` already rules most of them out at compile time; this is the
// guard for the specs that reach the builder untyped (JS callers, a spec
// `isChartSpec` has not narrowed).
const validateSpec = (spec: ReadChartSpec, usesComboFields: boolean): void => {
  if (usesComboFields && !COMBO_KINDS.has(spec.kind)) {
    throw new Error(
      `chart kind '${spec.kind}' does not support per-series chartKind / secondaryAxis (combo charts require a bar / column / line / area base kind)`,
    );
  }
  if (spec.view3D !== undefined) {
    if (!KINDS_WITH_VIEW_3D.has(spec.kind)) {
      throw new Error(
        `chart kind '${spec.kind}' has no 3-D variant; view3D applies to bar / column / line / area / pie / surface`,
      );
    }
    // PowerPoint cannot mix a 3-D plot group with any other group.
    if (usesComboFields) throw new Error('chart: view3D cannot be combined with a combo chart');
  }
  if (
    (spec.bar3DShape !== undefined || spec.gapDepthPct !== undefined) &&
    spec.view3D === undefined
  ) {
    throw new Error('chart: bar3DShape / gapDepthPct require view3D');
  }
  if (spec.ofPie !== undefined && spec.kind !== 'pie') {
    throw new Error(`chart: ofPie applies to kind 'pie', not '${spec.kind}'`);
  }
  if (spec.dataTable !== undefined && (AXISLESS_KINDS.has(spec.kind) || XY_KINDS.has(spec.kind))) {
    throw new Error(`chart kind '${spec.kind}' cannot show a data table`);
  }
  if (spec.categoryAxisDate !== undefined) {
    if (AXISLESS_KINDS.has(spec.kind) || XY_KINDS.has(spec.kind)) {
      throw new Error(`chart kind '${spec.kind}' has no category axis to make a date axis of`);
    }
    if (spec.categoryGroupLevels !== undefined) {
      throw new Error('chart: a date axis cannot have categoryGroupLevels');
    }
    spec.categories.forEach((category, i) => {
      if (category.trim() === '' || !Number.isFinite(Number(category))) {
        throw new Error(
          `chart: categories[${i}] is ${JSON.stringify(category)}; a date axis needs date serial numbers (e.g. "45292" for 2024-01-01)`,
        );
      }
    });
  }
  if (
    spec.categoryAxisScaling !== undefined &&
    spec.categoryAxisDate === undefined &&
    !XY_KINDS.has(spec.kind)
  ) {
    throw new Error(
      'chart: categoryAxisScaling needs a numeric horizontal axis — a scatter / bubble chart or a date axis',
    );
  }
  spec.categoryGroupLevels?.forEach((level, i) => {
    if (level.length !== spec.categories.length) {
      throw new Error(
        `chart: categoryGroupLevels[${i}] has ${level.length} entries for ${spec.categories.length} categories`,
      );
    }
  });
  if (XY_KINDS.has(spec.kind)) {
    spec.series.forEach((series) => {
      if (series.xValues === undefined) {
        throw new Error(`chart: ${spec.kind} series '${series.name}' needs xValues`);
      }
      if (spec.kind === 'bubble' && series.bubbleSizes === undefined) {
        throw new Error(`chart: bubble series '${series.name}' needs bubbleSizes`);
      }
    });
  }
};

// `<c:dTable>`. All four toggles are always written, as PowerPoint does, so a
// consumer's default for an absent toggle never decides how the table looks.
const dataTableElement = (table: NonNullable<ChartSpec['dataTable']>): XmlElement => {
  const children: XmlElement[] = [
    valNode(c('showHorzBorder'), table.showHorizontalBorder === false ? '0' : '1'),
    valNode(c('showVertBorder'), table.showVerticalBorder === false ? '0' : '1'),
    valNode(c('showOutline'), table.showOutline === false ? '0' : '1'),
    valNode(c('showKeys'), table.showKeys === false ? '0' : '1'),
  ];
  const txPr = axisTxPrElement(table.textStyle, undefined);
  if (txPr !== null) children.push(txPr);
  return elem(c('dTable'), { children });
};

/**
 * Builds a complete `<c:chartSpace>` document for `spec`. The caller
 * wires the `<c:externalData>` rel; this builder always emits
 * `r:id="rId1"` for the embedded workbook.
 */
export const buildChartSpaceDoc = (spec: ChartSpec): XmlDocument => {
  const usesComboFields = spec.series.some(
    (series) => series.chartKind !== undefined || series.secondaryAxis === true,
  );
  validateSpec(spec, usesComboFields);

  const layout = layoutChartSheet(spec);
  let plottedGroups: XmlElement[];
  let hasSecondary = false;
  if (usesComboFields) {
    const groups = comboPlotGroups(spec);
    hasSecondary = groups.some((group) => group.secondary);
    plottedGroups = groups.map((group) =>
      buildPlotGroup(
        {
          spec,
          layout,
          seriesIndices: group.indices,
          axes: group.secondary ? SECONDARY_AXES : PRIMARY_AXES,
        },
        group.kind,
      ),
    );
  } else {
    plottedGroups = [
      buildPlotGroup(
        { spec, layout, seriesIndices: spec.series.map((_, i) => i), axes: PRIMARY_AXES },
        spec.kind,
      ),
    ];
  }

  const plotLayout = spec.plotAreaLayout;
  const plotAreaChildren: XmlElement[] = [
    plotLayout !== undefined
      ? manualLayoutElement(plotLayout, plotLayout, plotLayout.target)
      : elem(c('layout')),
    ...plottedGroups,
  ];
  if (!AXISLESS_KINDS.has(spec.kind)) {
    plotAreaChildren.push(XY_KINDS.has(spec.kind) ? xValAxis(spec) : catAxis(spec), valAxis(spec));
    if (hasSeriesAxis(spec)) plotAreaChildren.push(serAxis(spec));
    if (hasSecondary) {
      plotAreaChildren.push(secondaryValAxis(spec), secondaryCatAxis(spec));
    }
  }
  if (spec.dataTable !== undefined) plotAreaChildren.push(dataTableElement(spec.dataTable));
  // <c:plotArea><c:spPr><a:solidFill> + optional <a:ln><a:solidFill>.
  if (spec.plotAreaFill !== undefined || spec.plotAreaStrokeColor !== undefined) {
    plotAreaChildren.push(spPrChildren(spec.plotAreaFill, spec.plotAreaStrokeColor));
  }
  const plotArea = elem(c('plotArea'), { children: plotAreaChildren });

  const chartChildren: XmlElement[] = [];
  if (spec.title !== undefined) {
    chartChildren.push(
      titleElement(spec.title, spec.titleStyle, undefined, {
        ...(spec.titleOverlay !== undefined ? { overlay: spec.titleOverlay } : {}),
        ...(spec.titleLayout !== undefined ? { layout: spec.titleLayout } : {}),
      }),
    );
  }
  chartChildren.push(valNode(c('autoTitleDeleted'), spec.title !== undefined ? '0' : '1'));
  // A surface chart is 3-D (or a top-down view of one) whether or not the
  // author positioned the camera, so it always carries <c:view3D>.
  if (spec.view3D !== undefined || spec.kind === 'surface') {
    chartChildren.push(view3DElement(spec));
  }
  chartChildren.push(plotArea);
  // <c:legend> after plotArea but before plotVisOnly / dispBlanksAs.
  // Skip entirely when `position` is explicitly null (author wants no
  // legend at all).
  if (spec.legend !== undefined && spec.legend.position !== null) {
    const legendChildren: XmlElement[] = [valNode(c('legendPos'), spec.legend.position)];
    for (const idx of spec.legend.hiddenIndices ?? []) {
      legendChildren.push(
        elem(c('legendEntry'), {
          children: [valNode(c('idx'), idx), valNode(c('delete'), '1')],
        }),
      );
    }
    // CT_Legend order: legendPos, legendEntry*, layout, overlay.
    if (spec.legend.layout !== undefined) {
      legendChildren.push(manualLayoutElement(spec.legend.layout, spec.legend.layout));
    }
    legendChildren.push(valNode(c('overlay'), spec.legend.overlay ? '1' : '0'));
    // <c:legend><c:txPr> carries authored legend font / color.
    if (spec.legend.textStyle !== undefined) {
      const txPr = axisTxPrElement(spec.legend.textStyle, undefined);
      if (txPr !== null) legendChildren.push(txPr);
    }
    chartChildren.push(elem(c('legend'), { children: legendChildren }));
  }
  chartChildren.push(
    valNode(c('plotVisOnly'), spec.plotVisibleCellsOnly === false ? '0' : '1'),
    valNode(c('dispBlanksAs'), spec.dispBlanksAs ?? 'gap'),
  );
  const chart = elem(c('chart'), { children: chartChildren });

  const externalData = elem(c('externalData'), {
    attrs: [attr(r('id'), 'rId1')],
    children: [valNode(c('autoUpdate'), '0')],
  });

  // <c:chartSpace><c:spPr> sits at the root and styles the entire card.
  // CT_ChartSpace schema order: date1904 → lang → roundedCorners →
  // style → chart.
  const rootChildren: XmlElement[] = [];
  if (spec.date1904) rootChildren.push(valNode(c('date1904'), '1'));
  if (spec.language !== undefined) rootChildren.push(valNode(c('lang'), spec.language));
  if (spec.roundedCorners) rootChildren.push(valNode(c('roundedCorners'), '1'));
  if (spec.chartStyle !== undefined) {
    rootChildren.push(valNode(c('style'), Math.round(spec.chartStyle)));
  }
  rootChildren.push(chart);
  if (spec.chartAreaFill !== undefined || spec.chartAreaStrokeColor !== undefined) {
    rootChildren.push(spPrChildren(spec.chartAreaFill, spec.chartAreaStrokeColor));
  }
  rootChildren.push(externalData);
  const root = elem(c('chartSpace'), {
    prefixDecls: new Map([
      ['c', NS_C],
      ['a', NS_A],
      ['r', NS_R],
    ]),
    children: rootChildren,
  });

  return {
    kind: 'document',
    decl: { version: '1.0', encoding: 'UTF-8', standalone: 'yes' },
    prolog: [],
    root,
    epilog: [],
  };
};

export type { ChartKind } from './types.ts';
export type { ChartSpec } from './chart-spec.ts';

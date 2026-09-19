// Chart XML builder.
//
// Produces a complete `<c:chartSpace>` for one of the supported chart
// kinds (bar / column / line / pie). The chart references an embedded
// xlsx via `<c:externalData r:id="rId1">`; the calling layer is
// responsible for wiring that rel and writing the xlsx bytes. Inline
// `<c:strCache>` / `<c:numCache>` blocks carry the values so PowerPoint
// can render the chart without ever opening the workbook.

import {
  firstSliceAngle,
  gapAmountPercent,
  holeSizePercent,
  lineWidthEmu as validateLineWidthEmu,
  overlapPercent,
} from '../bounds.ts';
import { NS, type XmlDocument, type XmlElement, attr, elem, qname, text } from '../xml/index.ts';
import type { ChartAxisScaling, ChartDataLabels, ChartSpec, ChartTextStyle } from './types.ts';

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

const numRef = (formula: string, points: ReadonlyArray<number | null>): XmlElement => {
  const numCache = elem(c('numCache'), {
    children: [
      elem(c('formatCode'), { children: [text('General')] }),
      valNode(c('ptCount'), points.length),
      ...points
        .map((v, i) => (v === null ? null : ptNode(i, String(v))))
        .filter((n): n is XmlElement => n !== null),
    ],
  });
  return elem(c('numRef'), {
    children: [elem(c('f'), { children: [text(formula)] }), numCache],
  });
};

const solidFillSpPr = (color: string): XmlElement => {
  const srgbClr = elem(a('srgbClr'), {
    attrs: [attr(qname('', 'val', ''), color.replace(/^#/, '').toUpperCase())],
  });
  const solidFill = elem(a('solidFill'), { children: [srgbClr] });
  return elem(c('spPr'), { children: [solidFill] });
};

// <c:majorGridlines|minorGridlines> with optional spPr/ln/solidFill/srgbClr
// for the gridline color. Centralizes the per-gridline color emit so all
// four axis-gridline slots share the same shape.
// `<c:spPr><a:ln w?>` with an optional solid color, or null when neither is
// authored. Every child of <a:ln> is optional, so a width alone is valid.
const lineSpPr = (
  color: string | undefined,
  widthEmu: number | undefined,
  field: string,
): XmlElement | null => {
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
  lineColor: string,
  lineWidthEmu: number | undefined,
  lineDash: string | undefined,
): XmlElement => {
  const lnChildren: XmlElement[] = [
    elem(a('solidFill'), {
      children: [elem(a('srgbClr'), { attrs: [attr(qname('', 'val', ''), lineColor)] })],
    }),
  ];
  if (lineDash !== undefined) {
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
  return elem(c('spPr'), {
    children: [
      elem(a('solidFill'), {
        children: [elem(a('srgbClr'), { attrs: [attr(qname('', 'val', ''), color)] })],
      }),
      ln,
    ],
  });
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
      // PowerPoint expects bubble3D=0 on every dPt outside bubble charts.
      valNode(c('bubble3D'), '0'),
    ];
    if (expl !== null) children.push(valNode(c('explosion'), Math.round(expl)));
    if (color !== null) {
      const hex = color.replace(/^#/, '').toUpperCase();
      children.push(solidFillSpPr(hex));
    }
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

const seriesElement = (spec: ChartSpec, seriesIdx: number, sheet: string): XmlElement => {
  const series = spec.series[seriesIdx];
  if (!series) throw new Error('seriesElement: out of range');

  const headerCellFormula = `${sheet}!$${String.fromCharCode(66 + seriesIdx)}$1`;
  const catRange = `${sheet}!$A$2:$A$${spec.categories.length + 1}`;
  const valRange = `${sheet}!$${String.fromCharCode(66 + seriesIdx)}$2:$${String.fromCharCode(66 + seriesIdx)}$${spec.categories.length + 1}`;

  const color =
    series.color !== undefined
      ? series.color.replace(/^#/, '').toUpperCase()
      : (DEFAULT_ACCENT_COLORS[seriesIdx % DEFAULT_ACCENT_COLORS.length] ?? '4472C4');

  // Right-pad values to category count so the chart aligns visually.
  const paddedValues: Array<number | null> = [];
  for (let i = 0; i < spec.categories.length; i++) {
    paddedValues.push(i < series.values.length ? (series.values[i] ?? null) : null);
  }

  // A combo series is shaped by its own plot group, not the chart's kind.
  const kind = series.chartKind ?? spec.kind;
  const children: XmlElement[] = [
    valNode(c('idx'), seriesIdx),
    valNode(c('order'), seriesIdx),
    elem(c('tx'), { children: [strRef(headerCellFormula, [series.name])] }),
  ];
  // spPr shape depends on what the series actually paints. A line series'
  // visible element is its stroke, so the color MUST live on <a:ln>; a bare
  // <a:solidFill> (correct for bar/area fills) leaves the line uncolored and
  // PowerPoint falls back to its automatic series palette — the line then
  // renders in the wrong color. seriesSpPr emits the color on both <a:ln>
  // and <a:solidFill>, so it colors the line for PowerPoint while keeping the
  // solidFill the reader round-trips. Bar / column / pie keep the legacy
  // solid-fill-only shape for tight round-trip compatibility with fixtures.
  if (
    kind === 'line' ||
    series.lineWidthEmu !== undefined ||
    series.lineColor !== undefined ||
    series.lineDash !== undefined
  ) {
    const lineColor =
      series.lineColor !== undefined ? series.lineColor.replace(/^#/, '').toUpperCase() : color;
    children.push(seriesSpPr(color, lineColor, series.lineWidthEmu, series.lineDash));
  } else {
    children.push(solidFillSpPr(color));
  }
  // invertIfNegative only exists on CT_BarSer (bar/column). Emitting it on a
  // line/pie/area series is schema-invalid, so gate on the bar-family kinds.
  if (series.invertIfNegative === true && (kind === 'bar' || kind === 'column')) {
    children.push(valNode(c('invertIfNegative'), '1'));
  }
  // <c:marker> is only valid on the marker-bearing series types
  // (CT_LineSer / CT_ScatterSer / CT_RadarSer). Emitting it on bar/column/
  // pie/doughnut/area produces schema-invalid XML.
  if (kind === 'line' || kind === 'scatter' || kind === 'radar') {
    const markerColor =
      series.markerColor !== undefined ? series.markerColor.replace(/^#/, '').toUpperCase() : color;
    const markerLineColor =
      series.markerLineColor !== undefined
        ? series.markerLineColor.replace(/^#/, '').toUpperCase()
        : markerColor;
    children.push(
      markerElement(series.markerSymbol, series.markerSizePt, markerColor, markerLineColor),
    );
  }
  // <c:dPt> overrides go after invertIfNegative / marker per schema.
  for (const dPt of dPtElements(series.pointColors, series.pointExplosions)) {
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
  if (series.trendline && kind !== 'pie' && kind !== 'doughnut' && kind !== 'radar') {
    children.push(trendlineElement(series.trendline));
  }
  children.push(elem(c('cat'), { children: [strRef(catRange, spec.categories)] }));
  children.push(elem(c('val'), { children: [numRef(valRange, paddedValues)] }));
  // Line series always get an explicit <c:smooth>: the schema default for an
  // absent element is val="1", so LibreOffice draws an unauthored line as a
  // smooth curve while PowerPoint draws it straight. PowerPoint itself always
  // writes the element; doing the same keeps every renderer straight unless
  // smoothing was asked for. (Only CT_LineSer carries smooth — emitting it on
  // a bar/pie series would be schema-invalid.)
  if (kind === 'line') {
    children.push(valNode(c('smooth'), series.smooth === true ? '1' : '0'));
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

const effectiveSeriesKind = (spec: ChartSpec, seriesIdx: number): string =>
  spec.series[seriesIdx]?.chartKind ?? spec.kind;

/**
 * Splits the series into plot groups keyed by (effective kind, axis).
 * Primary-axis groups come first so secondary overlays paint on top,
 * and bar groups precede line/area within each axis for the same reason
 * (matching PowerPoint's combo emit order).
 */
const comboPlotGroups = (
  spec: ChartSpec,
): { kind: string; secondary: boolean; indices: number[] }[] => {
  const groups = new Map<string, { kind: string; secondary: boolean; indices: number[] }>();
  for (let i = 0; i < spec.series.length; i++) {
    const kind = effectiveSeriesKind(spec, i);
    const secondary = spec.series[i]?.secondaryAxis === true;
    const key = `${kind}|${secondary ? '1' : '0'}`;
    const group = groups.get(key);
    if (group) group.indices.push(i);
    else groups.set(key, { kind, secondary, indices: [i] });
  }
  const paintOrder = (g: { kind: string; secondary: boolean }): number =>
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
  const children: XmlElement[] = [
    valNode(c('axId'), CAT_AX_ID),
    elem(c('scaling'), { children: [valNode(c('orientation'), catOrientation)] }),
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
  );
  if (catLine !== null) children.push(catLine);
  const catTxPr = axisTxPrElement(spec.categoryAxisLabelStyle, spec.categoryAxisLabelRotationDeg);
  if (catTxPr) children.push(catTxPr);
  children.push(valNode(c('crossAx'), VAL_AX_ID));
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
  const axisLine = lineSpPr(f.lineColor, f.lineWidthEmu, `chart: ${f.widthFields.line}`);
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
        children: [valNode(c('builtInUnit'), f.scaling.displayUnits)],
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
    children.push(elem(c('dLbl'), { children: [valNode(c('idx'), i), ...dLblGroupChildren(pl)] }));
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
  if (dl.textStyle !== undefined) {
    // CT_DLbls schema order places <c:txPr> before <c:dLblPos>; reusing
    // `axisTxPrElement` keeps the formatting parity with axis / legend.
    const txPr = axisTxPrElement(dl.textStyle, undefined);
    if (txPr !== null) children.push(txPr);
  }
  if (dl.position !== undefined) children.push(valNode(c('dLblPos'), dl.position));
  children.push(
    valNode(c('showLegendKey'), '0'),
    valNode(c('showVal'), dl.showValue ? '1' : '0'),
    valNode(c('showCatName'), dl.showCategory ? '1' : '0'),
    valNode(c('showSerName'), dl.showSeriesName ? '1' : '0'),
    valNode(c('showPercent'), dl.showPercent ? '1' : '0'),
    valNode(c('showBubbleSize'), '0'),
  );
  if (dl.separator !== undefined) {
    children.push(elem(c('separator'), { children: [text(dl.separator)] }));
  }
  return children;
};

const dLblsElement = (spec: ChartSpec): XmlElement | null => buildDLblsFromLabels(spec.dataLabels);

const buildBarChart = (
  spec: ChartSpec,
  sheet: string,
  direction: 'col' | 'bar',
  seriesIndices: ReadonlyArray<number>,
  axes: AxisIdPair,
): XmlElement => {
  const ser = seriesIndices.map((i) => seriesElement(spec, i, sheet));
  const dl = dLblsElement(spec);
  const grouping = spec.grouping ?? 'clustered';
  const children: XmlElement[] = [
    valNode(c('barDir'), direction),
    valNode(c('grouping'), grouping),
    valNode(c('varyColors'), spec.varyColors ? '1' : '0'),
    ...ser,
    ...(dl ? [dl] : []),
  ];
  if (spec.gapWidthPct !== undefined) {
    children.push(valNode(c('gapWidth'), gapAmountPercent(spec.gapWidthPct, 'chart: gapWidthPct')));
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
  children.push(valNode(c('axId'), axes.cat), valNode(c('axId'), axes.val));
  return elem(c(direction === 'col' ? 'barChart' : 'barChart'), { children });
};

// CT_Grouping (line / area) has no `clustered`, unlike CT_BarGrouping; a
// combo chart's bar grouping maps to the plain `standard` for those groups.
const lineAreaGrouping = (spec: ChartSpec): 'standard' | 'stacked' | 'percentStacked' =>
  spec.grouping === 'stacked' || spec.grouping === 'percentStacked' ? spec.grouping : 'standard';

const buildLineChart = (
  spec: ChartSpec,
  sheet: string,
  seriesIndices: ReadonlyArray<number>,
  axes: AxisIdPair,
): XmlElement => {
  const ser = seriesIndices.map((i) => seriesElement(spec, i, sheet));
  const dl = dLblsElement(spec);
  const children: XmlElement[] = [
    valNode(c('grouping'), lineAreaGrouping(spec)),
    valNode(c('varyColors'), spec.varyColors ? '1' : '0'),
    ...ser,
    ...(dl ? [dl] : []),
  ];
  if (spec.dropLines) children.push(elem(c('dropLines')));
  if (spec.hiLowLines) children.push(elem(c('hiLowLines')));
  // <c:marker val> selects the line subtype: "1" → Line with Markers,
  // "0" → plain Line. Default on, preserving the historical output;
  // authors opt out of markers with `lineMarkers: false`.
  children.push(
    valNode(c('marker'), spec.lineMarkers === false ? '0' : '1'),
    valNode(c('axId'), axes.cat),
    valNode(c('axId'), axes.val),
  );
  return elem(c('lineChart'), { children });
};

const buildPieChart = (spec: ChartSpec, sheet: string): XmlElement => {
  if (spec.series.length !== 1) {
    throw new Error('pie chart requires exactly one series');
  }
  const ser = seriesElement(spec, 0, sheet);
  const dl = dLblsElement(spec);
  const children: XmlElement[] = [valNode(c('varyColors'), '1'), ser, ...(dl ? [dl] : [])];
  if (spec.firstSliceAngleDeg !== undefined) {
    children.push(
      valNode(
        c('firstSliceAng'),
        firstSliceAngle(spec.firstSliceAngleDeg, 'chart: firstSliceAngleDeg'),
      ),
    );
  }
  return elem(c('pieChart'), { children });
};

const buildDoughnutChart = (spec: ChartSpec, sheet: string): XmlElement => {
  if (spec.series.length !== 1) {
    throw new Error('doughnut chart requires exactly one series');
  }
  const ser = seriesElement(spec, 0, sheet);
  const dl = dLblsElement(spec);
  const children: XmlElement[] = [valNode(c('varyColors'), '1'), ser, ...(dl ? [dl] : [])];
  if (spec.firstSliceAngleDeg !== undefined) {
    children.push(
      valNode(
        c('firstSliceAng'),
        firstSliceAngle(spec.firstSliceAngleDeg, 'chart: firstSliceAngleDeg'),
      ),
    );
  }
  children.push(
    valNode(c('holeSize'), holeSizePercent(spec.holeSizePct ?? 50, 'chart: holeSizePct')),
  );
  return elem(c('doughnutChart'), { children });
};

const buildAreaChart = (
  spec: ChartSpec,
  sheet: string,
  seriesIndices: ReadonlyArray<number>,
  axes: AxisIdPair,
): XmlElement => {
  const ser = seriesIndices.map((i) => seriesElement(spec, i, sheet));
  const dl = dLblsElement(spec);
  return elem(c('areaChart'), {
    children: [
      valNode(c('grouping'), lineAreaGrouping(spec)),
      valNode(c('varyColors'), spec.varyColors ? '1' : '0'),
      ...ser,
      ...(dl ? [dl] : []),
      valNode(c('axId'), axes.cat),
      valNode(c('axId'), axes.val),
    ],
  });
};

/** One combo plot group, dispatched by its effective kind. */
const buildComboGroupChart = (
  spec: ChartSpec,
  sheet: string,
  kind: string,
  seriesIndices: ReadonlyArray<number>,
  axes: AxisIdPair,
): XmlElement => {
  switch (kind) {
    case 'column':
      return buildBarChart(spec, sheet, 'col', seriesIndices, axes);
    case 'bar':
      return buildBarChart(spec, sheet, 'bar', seriesIndices, axes);
    case 'line':
      return buildLineChart(spec, sheet, seriesIndices, axes);
    case 'area':
      return buildAreaChart(spec, sheet, seriesIndices, axes);
    default:
      throw new Error(`combo chart: series chartKind '${kind}' is not authorable`);
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

const titleElement = (title: string, style?: ChartTextStyle, rotationDeg?: number): XmlElement => {
  // The paragraph defaults stay empty so an unset size falls back to the
  // application default; the run-level <a:rPr> carries the authored style.
  const pPr = elem(a('pPr'), { children: [elem(a('defRPr'))] });
  const { attrs: runAttrs, children: runChildren } = rPrAttrsFromStyle(style);
  const runRPr = elem(a('rPr'), {
    attrs: [attr(qname('', 'lang', ''), 'en-US'), ...runAttrs],
    children: runChildren,
  });
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
  return elem(c('title'), {
    children: [elem(c('tx'), { children: [rich] }), valNode(c('overlay'), '0')],
  });
};

/**
 * Builds a complete `<c:chartSpace>` document for `spec`. The caller
 * wires the `<c:externalData>` rel; this builder always emits
 * `r:id="rId1"` for the embedded workbook.
 */
export const buildChartSpaceDoc = (spec: ChartSpec): XmlDocument => {
  const sheet = 'Sheet1';

  const usesComboFields = spec.series.some(
    (series) => series.chartKind !== undefined || series.secondaryAxis === true,
  );
  if (usesComboFields && !COMBO_KINDS.has(spec.kind)) {
    throw new Error(
      `chart kind '${spec.kind}' does not support per-series chartKind / secondaryAxis (combo charts require a bar / column / line / area base kind)`,
    );
  }

  const allIndices = spec.series.map((_, i) => i);
  let plottedGroups: XmlElement[];
  let hasSecondary = false;
  switch (spec.kind) {
    case 'column':
    case 'bar':
    case 'line':
    case 'area': {
      if (usesComboFields) {
        const groups = comboPlotGroups(spec);
        hasSecondary = groups.some((group) => group.secondary);
        plottedGroups = groups.map((group) =>
          buildComboGroupChart(
            spec,
            sheet,
            group.kind,
            group.indices,
            group.secondary ? SECONDARY_AXES : PRIMARY_AXES,
          ),
        );
      } else {
        plottedGroups = [buildComboGroupChart(spec, sheet, spec.kind, allIndices, PRIMARY_AXES)];
      }
      break;
    }
    case 'pie':
      plottedGroups = [buildPieChart(spec, sheet)];
      break;
    case 'doughnut':
      plottedGroups = [buildDoughnutChart(spec, sheet)];
      break;
    case 'scatter':
    case 'radar':
    case 'bubble':
      // Read + render only (plan W4): the builder can't serialize the
      // xy(z) tuple channels these kinds need, so reject rather than
      // silently emit a malformed or wrong-kind chart. `readChartSpec`
      // surfaces these kinds, but `addSlideChart` / `setChartSpec` won't
      // write them.
      throw new Error(
        `chart kind '${spec.kind}' is read-only; authoring scatter / radar / bubble charts is not yet supported`,
      );
    default: {
      const exhaustive: never = spec.kind;
      throw new Error(`unsupported chart kind: ${String(exhaustive)}`);
    }
  }

  const axisless = spec.kind === 'pie' || spec.kind === 'doughnut';
  const plotAreaChildren: XmlElement[] = [elem(c('layout')), ...plottedGroups];
  if (!axisless) {
    plotAreaChildren.push(catAxis(spec), valAxis(spec));
    if (hasSecondary) {
      plotAreaChildren.push(secondaryValAxis(spec), secondaryCatAxis(spec));
    }
  }
  // <c:plotArea><c:spPr><a:solidFill> + optional <a:ln><a:solidFill>.
  if (spec.plotAreaFill !== undefined || spec.plotAreaStrokeColor !== undefined) {
    plotAreaChildren.push(spPrChildren(spec.plotAreaFill, spec.plotAreaStrokeColor));
  }
  const plotArea = elem(c('plotArea'), { children: plotAreaChildren });

  const chartChildren: XmlElement[] = [];
  if (spec.title !== undefined) {
    const titleEl = titleElement(spec.title, spec.titleStyle);
    // <c:title> can also carry <c:overlay val="1"/> — append it after
    // the <c:tx> child but before the synthesized overlay node from
    // titleElement so we don't end up with two overlay elements.
    if (spec.titleOverlay !== undefined) {
      const filtered = titleEl.children.filter(
        (c2) =>
          !(
            c2.kind === 'element' &&
            c2.name.namespaceURI === NS_C &&
            c2.name.localName === 'overlay'
          ),
      );
      titleEl.children = [...filtered, valNode(c('overlay'), spec.titleOverlay ? '1' : '0')];
    }
    chartChildren.push(titleEl);
  }
  chartChildren.push(
    valNode(c('autoTitleDeleted'), spec.title !== undefined ? '0' : '1'),
    plotArea,
  );
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

export type { ChartKind, ChartSpec } from './types.ts';

// Chart XML builder.
//
// Produces a complete `<c:chartSpace>` for one of the supported chart
// kinds (bar / column / line / pie). The chart references an embedded
// xlsx via `<c:externalData r:id="rId1">`; the calling layer is
// responsible for wiring that rel and writing the xlsx bytes. Inline
// `<c:strCache>` / `<c:numCache>` blocks carry the values so PowerPoint
// can render the chart without ever opening the workbook.

import { NS, type XmlDocument, type XmlElement, attr, elem, qname, text } from '../xml/index.ts';
import type { ChartSpec, ChartTextStyle } from './types.ts';

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
const gridlinesElement = (
  local: 'majorGridlines' | 'minorGridlines',
  color: string | undefined,
): XmlElement => {
  if (color === undefined) return elem(c(local));
  const hex = color.replace(/^#/, '').toUpperCase();
  const ln = elem(a('ln'), {
    children: [
      elem(a('solidFill'), {
        children: [elem(a('srgbClr'), { attrs: [attr(qname('', 'val', ''), hex)] })],
      }),
    ],
  });
  return elem(c(local), { children: [elem(c('spPr'), { children: [ln] })] });
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

// Builds <c:spPr> for a series — solidFill + optional <a:ln> (line
// width + dash). The reader extracts color from solidFill and width
// /dash from the ln; this writer keeps them in lock-step.
const seriesSpPr = (
  color: string,
  lineWidthEmu: number | undefined,
  lineDash: string | undefined,
): XmlElement => {
  const lnChildren: XmlElement[] = [
    elem(a('solidFill'), {
      children: [elem(a('srgbClr'), { attrs: [attr(qname('', 'val', ''), color)] })],
    }),
  ];
  if (lineDash !== undefined) {
    lnChildren.push(elem(a('prstDash'), { attrs: [attr(qname('', 'val', ''), lineDash)] }));
  }
  const ln =
    lineWidthEmu !== undefined
      ? elem(a('ln'), {
          attrs: [attr(qname('', 'w', ''), String(lineWidthEmu))],
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

// `<c:marker>` for a series. PowerPoint accepts symbol + size; the
// inner <c:spPr> can carry per-marker fill/stroke but we leave that to
// the existing series color.
const markerElement = (
  symbol: string | undefined,
  sizePt: number | undefined,
): XmlElement | null => {
  if (symbol === undefined && sizePt === undefined) return null;
  const children: XmlElement[] = [];
  if (symbol !== undefined) children.push(valNode(c('symbol'), symbol));
  if (sizePt !== undefined) children.push(valNode(c('size'), Math.round(sizePt)));
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

  const children: XmlElement[] = [
    valNode(c('idx'), seriesIdx),
    valNode(c('order'), seriesIdx),
    elem(c('tx'), { children: [strRef(headerCellFormula, [series.name])] }),
  ];
  // spPr — emit the richer version (with ln) only when authored fields
  // would otherwise be lost; otherwise keep the legacy solid-fill-only
  // shape for tight round-trip compatibility with the existing fixtures.
  if (series.lineWidthEmu !== undefined || series.lineDash !== undefined) {
    children.push(seriesSpPr(color, series.lineWidthEmu, series.lineDash));
  } else {
    children.push(solidFillSpPr(color));
  }
  if (series.invertIfNegative === true) {
    children.push(valNode(c('invertIfNegative'), '1'));
  }
  const mk = markerElement(series.markerSymbol, series.markerSizePt);
  if (mk !== null) children.push(mk);
  // <c:dPt> overrides go after invertIfNegative / marker per schema.
  for (const dPt of dPtElements(series.pointColors, series.pointExplosions)) {
    children.push(dPt);
  }
  const serDLbls = buildDLblsFromLabels(series.dataLabels);
  if (serDLbls !== null) children.push(serDLbls);
  if (series.trendline) children.push(trendlineElement(series.trendline));
  children.push(elem(c('cat'), { children: [strRef(catRange, spec.categories)] }));
  children.push(elem(c('val'), { children: [numRef(valRange, paddedValues)] }));
  if (series.smooth === true) {
    children.push(valNode(c('smooth'), '1'));
  }
  return elem(c('ser'), { children });
};

// Axis ids — arbitrary distinct positive 32-bit integers PowerPoint just
// needs them stable within the chart for the `<c:crossAx>` back-pointer.
const CAT_AX_ID = 111111111;
const VAL_AX_ID = 222222222;

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
    valNode(c('axPos'), 'b'),
  ];
  if (spec.categoryAxisMajorGridlines) {
    children.push(gridlinesElement('majorGridlines', spec.categoryAxisMajorGridlineColor));
  }
  if (spec.categoryAxisMinorGridlines) {
    children.push(gridlinesElement('minorGridlines', spec.categoryAxisMinorGridlineColor));
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
  if (spec.categoryAxisLineColor !== undefined) {
    children.push(spPrChildren(undefined, spec.categoryAxisLineColor));
  }
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
  return elem(c('catAx'), { children });
};

const valAxis = (spec: ChartSpec): XmlElement => {
  // <c:scaling> ordering matters: logBase, orientation, min, max.
  const scalingChildren: XmlElement[] = [];
  if (spec.valueAxis?.logBase !== undefined) {
    scalingChildren.push(valNode(c('logBase'), spec.valueAxis.logBase));
  }
  scalingChildren.push(valNode(c('orientation'), spec.valueAxisOrientation ?? 'minMax'));
  if (spec.valueAxis?.min !== undefined) {
    scalingChildren.push(valNode(c('min'), spec.valueAxis.min));
  }
  if (spec.valueAxis?.max !== undefined) {
    scalingChildren.push(valNode(c('max'), spec.valueAxis.max));
  }
  const children: XmlElement[] = [
    valNode(c('axId'), VAL_AX_ID),
    elem(c('scaling'), { children: scalingChildren }),
    valNode(c('delete'), spec.valueAxisHidden ? '1' : '0'),
    valNode(c('axPos'), 'l'),
  ];
  if (spec.valueAxisMajorGridlines) {
    children.push(gridlinesElement('majorGridlines', spec.valueAxisMajorGridlineColor));
  }
  if (spec.valueAxisMinorGridlines) {
    children.push(gridlinesElement('minorGridlines', spec.valueAxisMinorGridlineColor));
  }
  if (spec.valueAxisTitle !== undefined) {
    children.push(
      titleElement(spec.valueAxisTitle, spec.valueAxisTitleStyle, spec.valueAxisTitleRotationDeg),
    );
  }
  if (spec.valueAxis?.numberFormat !== undefined) {
    children.push(
      elem(c('numFmt'), {
        attrs: [
          attr(qname('', 'formatCode', ''), spec.valueAxis.numberFormat),
          attr(qname('', 'sourceLinked', ''), '0'),
        ],
      }),
    );
  }
  if (spec.valueAxisMajorTickMark !== undefined) {
    children.push(valNode(c('majorTickMark'), spec.valueAxisMajorTickMark));
  }
  if (spec.valueAxisMinorTickMark !== undefined) {
    children.push(valNode(c('minorTickMark'), spec.valueAxisMinorTickMark));
  }
  if (spec.valueAxisLineColor !== undefined) {
    children.push(spPrChildren(undefined, spec.valueAxisLineColor));
  }
  const valTxPr = axisTxPrElement(spec.valueAxisLabelStyle, spec.valueAxisLabelRotationDeg);
  if (valTxPr) children.push(valTxPr);
  children.push(valNode(c('crossAx'), CAT_AX_ID));
  // <c:crosses val>/`crossesAt val>` — mutually exclusive per the
  // schema. Object form `{ at: N }` → crossesAt; string form → crosses.
  const xross = spec.valueAxisCrosses;
  if (xross !== undefined) {
    if (typeof xross === 'string') {
      children.push(valNode(c('crosses'), xross));
    } else {
      children.push(valNode(c('crossesAt'), String(xross.at)));
    }
  }
  if (spec.valueAxisCrossBetween !== undefined) {
    children.push(valNode(c('crossBetween'), spec.valueAxisCrossBetween));
  }
  if (spec.valueAxis?.majorUnit !== undefined) {
    children.push(valNode(c('majorUnit'), spec.valueAxis.majorUnit));
  }
  if (spec.valueAxis?.minorUnit !== undefined) {
    children.push(valNode(c('minorUnit'), spec.valueAxis.minorUnit));
  }
  if (spec.valueAxis?.displayUnits !== undefined) {
    children.push(
      elem(c('dispUnits'), {
        children: [valNode(c('builtInUnit'), spec.valueAxis.displayUnits)],
      }),
    );
  }
  return elem(c('valAx'), { children });
};

// Build `<c:dLbls>` from a ChartDataLabels (showVal / showCatName /
// showSerName / showPercent toggles plus optional numFmt, position,
// separator). Returns `null` when no dataLabels were authored so
// callers know to skip the element entirely.
const buildDLblsFromLabels = (dl: ChartSpec['dataLabels'] | undefined): XmlElement | null => {
  if (!dl) return null;
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
  return elem(c('dLbls'), { children });
};

const dLblsElement = (spec: ChartSpec): XmlElement | null => buildDLblsFromLabels(spec.dataLabels);

const buildBarChart = (spec: ChartSpec, sheet: string, direction: 'col' | 'bar'): XmlElement => {
  const ser = spec.series.map((_, i) => seriesElement(spec, i, sheet));
  const dl = dLblsElement(spec);
  const grouping = spec.grouping ?? 'clustered';
  const children: XmlElement[] = [
    valNode(c('barDir'), direction),
    valNode(c('grouping'), grouping),
    valNode(c('varyColors'), spec.varyColors ? '1' : '0'),
    ...ser,
    ...(dl ? [dl] : []),
  ];
  if (spec.gapWidthPct !== undefined) children.push(valNode(c('gapWidth'), spec.gapWidthPct));
  if (spec.overlapPct !== undefined) children.push(valNode(c('overlap'), spec.overlapPct));
  children.push(valNode(c('axId'), CAT_AX_ID), valNode(c('axId'), VAL_AX_ID));
  return elem(c(direction === 'col' ? 'barChart' : 'barChart'), { children });
};

const buildLineChart = (spec: ChartSpec, sheet: string): XmlElement => {
  const ser = spec.series.map((_, i) => seriesElement(spec, i, sheet));
  const dl = dLblsElement(spec);
  const children: XmlElement[] = [
    valNode(c('grouping'), spec.grouping ?? 'standard'),
    valNode(c('varyColors'), spec.varyColors ? '1' : '0'),
    ...ser,
    ...(dl ? [dl] : []),
  ];
  if (spec.dropLines) children.push(elem(c('dropLines')));
  if (spec.hiLowLines) children.push(elem(c('hiLowLines')));
  children.push(
    valNode(c('marker'), '1'),
    valNode(c('axId'), CAT_AX_ID),
    valNode(c('axId'), VAL_AX_ID),
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
    children.push(valNode(c('firstSliceAng'), Math.round(spec.firstSliceAngleDeg)));
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
    children.push(valNode(c('firstSliceAng'), Math.round(spec.firstSliceAngleDeg)));
  }
  children.push(valNode(c('holeSize'), spec.holeSizePct ?? 50));
  return elem(c('doughnutChart'), { children });
};

const buildAreaChart = (spec: ChartSpec, sheet: string): XmlElement => {
  const ser = spec.series.map((_, i) => seriesElement(spec, i, sheet));
  const dl = dLblsElement(spec);
  return elem(c('areaChart'), {
    children: [
      valNode(c('grouping'), spec.grouping ?? 'standard'),
      valNode(c('varyColors'), spec.varyColors ? '1' : '0'),
      ...ser,
      ...(dl ? [dl] : []),
      valNode(c('axId'), CAT_AX_ID),
      valNode(c('axId'), VAL_AX_ID),
    ],
  });
};

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
  return { attrs, children };
};

const titleElement = (title: string, style?: ChartTextStyle, rotationDeg?: number): XmlElement => {
  // defRPr stays 1400 (14pt) as the legacy default; the run-level
  // <a:rPr> carries authored overrides so renderers honor them.
  const rPr = elem(a('defRPr'), { attrs: [attr(qname('', 'sz', ''), '1400')] });
  const pPr = elem(a('pPr'), { children: [rPr] });
  const { attrs: runAttrs, children: runChildren } = rPrAttrsFromStyle(style);
  const runRPr = elem(a('rPr'), {
    attrs: [attr(qname('', 'lang', ''), 'en-US'), ...runAttrs],
    children: runChildren,
  });
  const tRun = elem(a('r'), {
    children: [runRPr, elem(a('t'), { children: [text(title)] })],
  });
  const para = elem(a('p'), { children: [pPr, tRun] });
  const rotStr =
    rotationDeg !== undefined && rotationDeg !== 0 ? String(Math.round(rotationDeg * 60000)) : '0';
  const rich = elem(c('rich'), {
    children: [
      elem(a('bodyPr'), {
        attrs: [
          attr(qname('', 'rot', ''), rotStr),
          attr(qname('', 'spcFirstLastPara', ''), '1'),
          attr(qname('', 'vertOverflow', ''), 'ellipsis'),
          attr(qname('', 'vert', ''), 'horz'),
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

  let plotted: XmlElement;
  switch (spec.kind) {
    case 'column':
      plotted = buildBarChart(spec, sheet, 'col');
      break;
    case 'bar':
      plotted = buildBarChart(spec, sheet, 'bar');
      break;
    case 'line':
      plotted = buildLineChart(spec, sheet);
      break;
    case 'pie':
      plotted = buildPieChart(spec, sheet);
      break;
    case 'doughnut':
      plotted = buildDoughnutChart(spec, sheet);
      break;
    case 'area':
      plotted = buildAreaChart(spec, sheet);
      break;
    default: {
      const exhaustive: never = spec.kind;
      throw new Error(`unsupported chart kind: ${String(exhaustive)}`);
    }
  }

  const axisless = spec.kind === 'pie' || spec.kind === 'doughnut';
  const plotAreaChildren: XmlElement[] = [elem(c('layout')), plotted];
  if (!axisless) {
    plotAreaChildren.push(catAxis(spec), valAxis(spec));
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
  // CT_ChartSpace schema order: roundedCorners → style → chart.
  const rootChildren: XmlElement[] = [];
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

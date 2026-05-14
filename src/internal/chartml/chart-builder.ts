// Chart XML builder.
//
// Produces a complete `<c:chartSpace>` for one of the supported chart
// kinds (bar / column / line / pie). The chart references an embedded
// xlsx via `<c:externalData r:id="rId1">`; the calling layer is
// responsible for wiring that rel and writing the xlsx bytes. Inline
// `<c:strCache>` / `<c:numCache>` blocks carry the values so PowerPoint
// can render the chart without ever opening the workbook.

import { NS, type XmlDocument, type XmlElement, attr, elem, qname, text } from '../xml/index.ts';
import type { ChartSpec } from './types.ts';

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
    children: [
      valNode(c('ptCount'), points.length),
      ...points.map((p, i) => ptNode(i, p)),
    ],
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

// Default theme accent palette (matches Office 2013+ default theme).
const DEFAULT_ACCENT_COLORS = [
  '4472C4', // accent1
  'ED7D31', // accent2
  'A5A5A5', // accent3
  'FFC000', // accent4
  '5B9BD5', // accent5
  '70AD47', // accent6
];

const seriesElement = (
  spec: ChartSpec,
  seriesIdx: number,
  sheet: string,
): XmlElement => {
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

  return elem(c('ser'), {
    children: [
      valNode(c('idx'), seriesIdx),
      valNode(c('order'), seriesIdx),
      elem(c('tx'), { children: [strRef(headerCellFormula, [series.name])] }),
      solidFillSpPr(color),
      elem(c('cat'), { children: [strRef(catRange, spec.categories)] }),
      elem(c('val'), { children: [numRef(valRange, paddedValues)] }),
    ],
  });
};

// Axis ids — arbitrary distinct positive 32-bit integers PowerPoint just
// needs them stable within the chart for the `<c:crossAx>` back-pointer.
const CAT_AX_ID = 111111111;
const VAL_AX_ID = 222222222;

const catAxis = (): XmlElement =>
  elem(c('catAx'), {
    children: [
      valNode(c('axId'), CAT_AX_ID),
      elem(c('scaling'), { children: [valNode(c('orientation'), 'minMax')] }),
      valNode(c('delete'), '0'),
      valNode(c('axPos'), 'b'),
      valNode(c('crossAx'), VAL_AX_ID),
    ],
  });

const valAxis = (): XmlElement =>
  elem(c('valAx'), {
    children: [
      valNode(c('axId'), VAL_AX_ID),
      elem(c('scaling'), { children: [valNode(c('orientation'), 'minMax')] }),
      valNode(c('delete'), '0'),
      valNode(c('axPos'), 'l'),
      valNode(c('crossAx'), CAT_AX_ID),
    ],
  });

const buildBarChart = (spec: ChartSpec, sheet: string, direction: 'col' | 'bar'): XmlElement => {
  const ser = spec.series.map((_, i) => seriesElement(spec, i, sheet));
  return elem(c(direction === 'col' ? 'barChart' : 'barChart'), {
    children: [
      valNode(c('barDir'), direction),
      valNode(c('grouping'), 'clustered'),
      valNode(c('varyColors'), '0'),
      ...ser,
      valNode(c('axId'), CAT_AX_ID),
      valNode(c('axId'), VAL_AX_ID),
    ],
  });
};

const buildLineChart = (spec: ChartSpec, sheet: string): XmlElement => {
  const ser = spec.series.map((_, i) => seriesElement(spec, i, sheet));
  return elem(c('lineChart'), {
    children: [
      valNode(c('grouping'), 'standard'),
      valNode(c('varyColors'), '0'),
      ...ser,
      valNode(c('marker'), '1'),
      valNode(c('axId'), CAT_AX_ID),
      valNode(c('axId'), VAL_AX_ID),
    ],
  });
};

const buildPieChart = (spec: ChartSpec, sheet: string): XmlElement => {
  if (spec.series.length !== 1) {
    throw new Error('pie chart requires exactly one series');
  }
  const ser = seriesElement(spec, 0, sheet);
  return elem(c('pieChart'), {
    children: [valNode(c('varyColors'), '1'), ser],
  });
};

const buildDoughnutChart = (spec: ChartSpec, sheet: string): XmlElement => {
  if (spec.series.length !== 1) {
    throw new Error('doughnut chart requires exactly one series');
  }
  const ser = seriesElement(spec, 0, sheet);
  return elem(c('doughnutChart'), {
    children: [
      valNode(c('varyColors'), '1'),
      ser,
      // 50% hole — PowerPoint's default.
      valNode(c('holeSize'), '50'),
    ],
  });
};

const buildAreaChart = (spec: ChartSpec, sheet: string): XmlElement => {
  const ser = spec.series.map((_, i) => seriesElement(spec, i, sheet));
  return elem(c('areaChart'), {
    children: [
      valNode(c('grouping'), 'standard'),
      valNode(c('varyColors'), '0'),
      ...ser,
      valNode(c('axId'), CAT_AX_ID),
      valNode(c('axId'), VAL_AX_ID),
    ],
  });
};

const titleElement = (title: string): XmlElement => {
  const rPr = elem(a('defRPr'), { attrs: [attr(qname('', 'sz', ''), '1400')] });
  const pPr = elem(a('pPr'), { children: [rPr] });
  const tRun = elem(a('r'), {
    children: [
      elem(a('rPr'), { attrs: [attr(qname('', 'lang', ''), 'en-US')] }),
      elem(a('t'), { children: [text(title)] }),
    ],
  });
  const para = elem(a('p'), { children: [pPr, tRun] });
  const rich = elem(c('rich'), {
    children: [
      elem(a('bodyPr'), {
        attrs: [
          attr(qname('', 'rot', ''), '0'),
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
    children: [
      elem(c('tx'), { children: [rich] }),
      valNode(c('overlay'), '0'),
    ],
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
    plotAreaChildren.push(catAxis(), valAxis());
  }
  const plotArea = elem(c('plotArea'), { children: plotAreaChildren });

  const chartChildren: XmlElement[] = [];
  if (spec.title !== undefined) chartChildren.push(titleElement(spec.title));
  chartChildren.push(
    valNode(c('autoTitleDeleted'), spec.title !== undefined ? '0' : '1'),
    plotArea,
    valNode(c('plotVisOnly'), '1'),
    valNode(c('dispBlanksAs'), 'gap'),
  );
  const chart = elem(c('chart'), { children: chartChildren });

  const externalData = elem(c('externalData'), {
    attrs: [attr(r('id'), 'rId1')],
    children: [valNode(c('autoUpdate'), '0')],
  });

  const root = elem(c('chartSpace'), {
    prefixDecls: new Map([
      ['c', NS_C],
      ['a', NS_A],
      ['r', NS_R],
    ]),
    children: [chart, externalData],
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

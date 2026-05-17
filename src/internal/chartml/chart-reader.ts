// Chart spec reader — parses `<c:chartSpace>` back into a `ChartSpec`.
//
// Companion to `chart-builder.ts`. Lets callers introspect or edit
// chart data on existing templates without dropping to XML. The reader
// uses inline `<c:strCache>` / `<c:numCache>` blocks rather than the
// embedded workbook — the cache is what PowerPoint renders from
// anyway, and it's always present in charts pptx-kit emits.

import {
  NS,
  type XmlElement,
  allChildElements,
  firstChildElement,
  getAttrValue,
  qname,
} from '../xml/index.ts';
import type { ChartDataLabels, ChartKind, ChartSeries, ChartSpec } from './types.ts';

const NS_C = NS.chart;
const NS_A = NS.dml;

const NAME_CHART_SPACE = qname('c', 'chartSpace', NS_C);
const NAME_CHART = qname('c', 'chart', NS_C);
const NAME_PLOT_AREA = qname('c', 'plotArea', NS_C);
const NAME_SER = qname('c', 'ser', NS_C);
const NAME_TX = qname('c', 'tx', NS_C);
const NAME_CAT = qname('c', 'cat', NS_C);
const NAME_VAL = qname('c', 'val', NS_C);
const NAME_STR_REF = qname('c', 'strRef', NS_C);
const NAME_NUM_REF = qname('c', 'numRef', NS_C);
const NAME_STR_CACHE = qname('c', 'strCache', NS_C);
const NAME_NUM_CACHE = qname('c', 'numCache', NS_C);
const NAME_STR_LIT = qname('c', 'strLit', NS_C);
const NAME_NUM_LIT = qname('c', 'numLit', NS_C);
const NAME_PT = qname('c', 'pt', NS_C);
const NAME_V = qname('c', 'v', NS_C);
const NAME_TITLE = qname('c', 'title', NS_C);
const NAME_RICH = qname('c', 'rich', NS_C);
const NAME_T = qname('a', 't', NS_A);
const NAME_P_DML = qname('a', 'p', NS_A);
const NAME_R_DML = qname('a', 'r', NS_A);
const NAME_SP_PR_C = qname('c', 'spPr', NS_C);
const NAME_SOLID_FILL = qname('a', 'solidFill', NS_A);
const NAME_SRGB_CLR = qname('a', 'srgbClr', NS_A);

const ATTR_VAL = qname('', 'val', '');
const ATTR_IDX = qname('', 'idx', '');

interface PlottedKindMap {
  readonly localName: string;
  readonly kind: ChartKind;
}

const KIND_MAP: ReadonlyArray<PlottedKindMap> = [
  // `barChart` is overloaded; `<c:barDir val="bar"/>` vs `"col"` decides.
  { localName: 'barChart', kind: 'column' },
  // The 3D variants share the same `<c:ser>` schema as their flat
  // counterparts; we degrade to the flat kind so renderers don't have
  // to special-case them. PowerPoint's own embedded data view does the
  // same flattening when "Edit data" is opened.
  { localName: 'bar3DChart', kind: 'column' },
  { localName: 'lineChart', kind: 'line' },
  { localName: 'line3DChart', kind: 'line' },
  { localName: 'pieChart', kind: 'pie' },
  { localName: 'pie3DChart', kind: 'pie' },
  { localName: 'ofPieChart', kind: 'pie' },
  { localName: 'doughnutChart', kind: 'doughnut' },
  { localName: 'areaChart', kind: 'area' },
  { localName: 'area3DChart', kind: 'area' },
  // Scatter / bubble carry xy / xyz tuples per series rather than
  // numeric channels against shared categories. We degrade to line
  // so the y-axis numbers + connecting strokes show up; the x-axis
  // collapses to index positions. Better than the "unsupported kind"
  // placeholder by a wide margin for the common single-series case.
  { localName: 'scatterChart', kind: 'line' },
  { localName: 'bubbleChart', kind: 'line' },
  // Radar collapses to a vertical line chart — the legend + axis
  // numbers are still readable.
  { localName: 'radarChart', kind: 'line' },
  // Stock charts: open / high / low / close as a four-line plot.
  { localName: 'stockChart', kind: 'line' },
  // Surface degrades to a column chart so the data table is visible.
  { localName: 'surfaceChart', kind: 'column' },
  { localName: 'surface3DChart', kind: 'column' },
];

const findFirst = (parent: XmlElement, names: ReadonlyArray<string>): XmlElement | null => {
  for (const c of parent.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS_C) continue;
    if (names.includes(c.name.localName)) return c;
  }
  return null;
};

const readPtArray = (cache: XmlElement): string[] => {
  const out: string[] = [];
  for (const pt of allChildElements(cache, NAME_PT)) {
    const idxRaw = getAttrValue(pt, ATTR_IDX);
    if (idxRaw === null) continue;
    const idx = Number.parseInt(idxRaw, 10);
    if (!Number.isFinite(idx) || idx < 0) continue;
    const v = firstChildElement(pt, NAME_V);
    if (v === null) continue;
    let text = '';
    for (const child of v.children) {
      if (child.kind === 'text' || child.kind === 'cdata') text += child.data;
    }
    out[idx] = text;
  }
  return out;
};

// `<c:strRef>` (cell range) and `<c:strLit>` (literal array) both serialise
// their points the same way: `<c:pt idx="...">...<c:v>text</c:v></c:pt>`.
// PowerPoint authors usually emit `strRef` with a `strCache`; other writers
// (python-pptx, pptxgenjs's older paths, hand-edited XML) skip the cache or
// drop the workbook entirely and emit `strLit`. Either way the cached points
// are enough to render the chart, so we accept both.
const readStringChannel = (parent: XmlElement): string[] | null => {
  const ref = firstChildElement(parent, NAME_STR_REF);
  if (ref) {
    const cache = firstChildElement(ref, NAME_STR_CACHE);
    if (cache) return readPtArray(cache);
  }
  const lit = firstChildElement(parent, NAME_STR_LIT);
  if (lit) return readPtArray(lit);
  return null;
};

const readStringRef = (parent: XmlElement): string[] | null => readStringChannel(parent);

// Same dual handling for numeric channels — `<c:numRef>/<c:numCache>` for
// workbook-referenced values, `<c:numLit>` for literal arrays. Without this
// fallback, charts authored as inline literals come through with empty
// `values` arrays and the renderer has nothing to plot.
const readNumChannel = (parent: XmlElement): Array<number | null> | null => {
  const ref = firstChildElement(parent, NAME_NUM_REF);
  let raw: string[] | null = null;
  if (ref) {
    const cache = firstChildElement(ref, NAME_NUM_CACHE);
    if (cache) raw = readPtArray(cache);
  }
  if (!raw) {
    const lit = firstChildElement(parent, NAME_NUM_LIT);
    if (lit) raw = readPtArray(lit);
  }
  if (!raw) return null;
  return raw.map((s) => {
    if (s === undefined || s === '') return null;
    const n = Number.parseFloat(s);
    return Number.isFinite(n) ? n : null;
  });
};

const readNumRef = (parent: XmlElement): Array<number | null> | null => readNumChannel(parent);

const readSeriesName = (ser: XmlElement): string => {
  const tx = firstChildElement(ser, NAME_TX);
  if (!tx) return '';
  const strs = readStringRef(tx);
  return strs?.[0] ?? '';
};

const readSeriesColor = (ser: XmlElement): string | undefined => {
  const spPr = firstChildElement(ser, NAME_SP_PR_C);
  if (!spPr) return undefined;
  const solidFill = firstChildElement(spPr, NAME_SOLID_FILL);
  if (!solidFill) return undefined;
  const srgb = firstChildElement(solidFill, NAME_SRGB_CLR);
  if (!srgb) return undefined;
  const v = getAttrValue(srgb, ATTR_VAL);
  return v !== null ? `#${v.toUpperCase()}` : undefined;
};

const readTitle = (chart: XmlElement): string | undefined => {
  const title = firstChildElement(chart, NAME_TITLE);
  if (!title) return undefined;
  const tx = firstChildElement(title, NAME_TX);
  if (!tx) return undefined;
  const rich = firstChildElement(tx, NAME_RICH);
  if (!rich) return undefined;
  let acc = '';
  for (const p of allChildElements(rich, NAME_P_DML)) {
    for (const r of allChildElements(p, NAME_R_DML)) {
      const tEl = firstChildElement(r, NAME_T);
      if (!tEl) continue;
      for (const child of tEl.children) {
        if (child.kind === 'text' || child.kind === 'cdata') acc += child.data;
      }
    }
  }
  return acc.length > 0 ? acc : undefined;
};

/**
 * Parses a `<c:chartSpace>` element into a typed `ChartSpec`. Throws if
 * the root or any required child is missing. Returns `null` only when
 * the chart is structurally well-formed but uses a kind we don't model
 * (so callers can fall through to pass-through).
 */
export const readChartSpec = (root: XmlElement): ChartSpec | null => {
  if (root.name.namespaceURI !== NS_C || root.name.localName !== 'chartSpace') {
    throw new Error(
      `expected <c:chartSpace> root, got <${root.name.prefix}:${root.name.localName}>`,
    );
  }
  void NAME_CHART_SPACE;
  const chart = firstChildElement(root, NAME_CHART);
  if (!chart) throw new Error('<c:chartSpace> has no <c:chart>');
  const plotArea = firstChildElement(chart, NAME_PLOT_AREA);
  if (!plotArea) throw new Error('<c:chart> has no <c:plotArea>');

  // Find which "plotted" element the plotArea carries.
  let plotted: XmlElement | null = null;
  let kind: ChartKind | null = null;
  for (const candidate of KIND_MAP) {
    const found = findFirst(plotArea, [candidate.localName]);
    if (found) {
      plotted = found;
      kind = candidate.kind;
      // Resolve bar vs column on a `barChart` / `bar3DChart`.
      if (candidate.localName === 'barChart' || candidate.localName === 'bar3DChart') {
        const barDir = firstChildElement(found, qname('c', 'barDir', NS_C));
        const v = barDir !== null ? getAttrValue(barDir, ATTR_VAL) : null;
        kind = v === 'bar' ? 'bar' : 'column';
      }
      break;
    }
  }
  if (!plotted || !kind) return null;

  // Read every <c:ser> in order.
  const series: ChartSeries[] = [];
  let categoriesFromFirst: string[] | null = null;
  for (const ser of allChildElements(plotted, NAME_SER)) {
    const name = readSeriesName(ser);
    const cat = firstChildElement(ser, NAME_CAT);
    if (cat !== null && categoriesFromFirst === null) {
      categoriesFromFirst = readStringRef(cat) ?? null;
    }
    let valEl = firstChildElement(ser, NAME_VAL);
    // Scatter / bubble charts carry numeric data on <c:yVal> rather
    // than <c:val>; surface those so the line-chart degradation has
    // something to plot. Bubble's <c:bubbleSize> is ignored — we
    // don't have a per-point sizing channel yet.
    if (!valEl) {
      valEl = firstChildElement(ser, qname('c', 'yVal', NS_C));
    }
    const values = valEl !== null ? readNumRef(valEl) : null;
    const color = readSeriesColor(ser);
    series.push({
      name,
      values: values ?? [],
      ...(color !== undefined ? { color } : {}),
    });
  }

  const categories = categoriesFromFirst ?? [];
  const title = readTitle(chart);

  // <c:dLbls> can sit either on the plotted-kind element (`barChart`,
  // `lineChart`, …) for chart-level defaults or on each `<c:ser>` for
  // per-series overrides. Surface the plotted-kind defaults; per-series
  // toggles are deferred until renderers care.
  const dLbls = firstChildElement(plotted, qname('c', 'dLbls', NS_C));
  let dataLabels: ChartDataLabels | undefined;
  if (dLbls) {
    const readToggle = (local: string): boolean => {
      const el = firstChildElement(dLbls, qname('c', local, NS_C));
      if (!el) return false;
      const v = getAttrValue(el, ATTR_VAL);
      // Absent val attribute defaults to true per the schema's CT_Boolean.
      return v === null || v === '1' || v === 'true';
    };
    dataLabels = {
      showValue: readToggle('showVal'),
      showCategory: readToggle('showCatName'),
      showSeriesName: readToggle('showSerName'),
      showPercent: readToggle('showPercent'),
    };
  }

  return {
    kind,
    categories,
    series,
    ...(title !== undefined ? { title } : {}),
    ...(dataLabels !== undefined ? { dataLabels } : {}),
  };
};

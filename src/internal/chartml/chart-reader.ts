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
import type { ChartKind, ChartSeries, ChartSpec } from './types.ts';

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
  { localName: 'lineChart', kind: 'line' },
  { localName: 'pieChart', kind: 'pie' },
  { localName: 'doughnutChart', kind: 'doughnut' },
  { localName: 'areaChart', kind: 'area' },
];

const findFirst = (parent: XmlElement, names: ReadonlyArray<string>): XmlElement | null => {
  for (const c of parent.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS_C) continue;
    if (names.includes(c.name.localName)) return c;
  }
  return null;
};

const readPtArray = (
  cache: XmlElement,
): string[] => {
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

const readStringRef = (parent: XmlElement): string[] | null => {
  const ref = firstChildElement(parent, NAME_STR_REF);
  if (!ref) return null;
  const cache = firstChildElement(ref, NAME_STR_CACHE);
  if (!cache) return null;
  return readPtArray(cache);
};

const readNumRef = (parent: XmlElement): Array<number | null> | null => {
  const ref = firstChildElement(parent, NAME_NUM_REF);
  if (!ref) return null;
  const cache = firstChildElement(ref, NAME_NUM_CACHE);
  if (!cache) return null;
  const raw = readPtArray(cache);
  return raw.map((s) => {
    if (s === undefined || s === '') return null;
    const n = Number.parseFloat(s);
    return Number.isFinite(n) ? n : null;
  });
};

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
      // Resolve bar vs column on a `barChart`.
      if (candidate.localName === 'barChart') {
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
    const valEl = firstChildElement(ser, NAME_VAL);
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

  return {
    kind,
    categories,
    series,
    ...(title !== undefined ? { title } : {}),
  };
};

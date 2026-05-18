// Charts (read + write).

import type { Emu } from '../units.ts';
import {
  type PartName,
  emptyRels,
  nextRelId,
  partName,
  resolveTarget,
} from '../../internal/opc/index.ts';
import type { OpcPackage } from '../../internal/parts/index.ts';
import { REL_TYPES } from '../../internal/presentationml/index.ts';
import {
  type ChartKind,
  type ChartSeries,
  type ChartSpec,
  buildChartSpaceDoc,
  buildEmbeddedXlsx,
  readChartSpec,
} from '../../internal/chartml/index.ts';
import {
  NS,
  type XmlElement,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  parseXml,
  qname,
  serializeXml,
  text as textNode,
} from '../../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  SHAPE_SNAPSHOT,
  SLIDE_PART_NAME,
  SLIDE_SHAPES,
  type SlideData,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import { appendAndReturnNewShape, decode, encode, nextShapeId, setOpcDefault } from './_helpers.ts';
// ---------------------------------------------------------------------------
// Charts.
//
// Authoring path for ChartML (`/ppt/charts/chart{N}.xml`) + the embedded
// `/ppt/embeddings/Microsoft_Excel_Worksheet{N}.xlsx` workbook that
// PowerPoint requires for the "Edit data" action to work. See plan §P9
// and §Risks for the scope constraints.
//
// Public surface is intentionally narrow: one `addSlideChart` entry point
// that takes a typed `ChartSpec`. The internal layer handles the chart
// XML, the embedded xlsx ZIP, and all the relationship wiring.

const CHART_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const EMBEDDED_XLSX_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const allocateChartIndex = (pkg: OpcPackage): number => {
  let next = 1;
  const re = /^\/ppt\/charts\/chart(\d+)\.xml$/;
  for (const p of pkg.parts) {
    const m = p.name.match(re);
    if (m?.[1] !== undefined) {
      const n = Number.parseInt(m[1], 10);
      if (Number.isFinite(n) && n >= next) next = n + 1;
    }
  }
  return next;
};

const NAME_GRAPHIC_FRAME = qname('p', 'graphicFrame', NS.pml);
const NAME_NV_GRAPHIC_FRAME_PR = qname('p', 'nvGraphicFramePr', NS.pml);
const NAME_C_NV_PR_FN = qname('p', 'cNvPr', NS.pml);
const NAME_C_NV_GRAPHIC_FRAME_PR = qname('p', 'cNvGraphicFramePr', NS.pml);
const NAME_NV_PR = qname('p', 'nvPr', NS.pml);
const NAME_XFRM = qname('p', 'xfrm', NS.pml);
const NAME_OFF = qname('a', 'off', NS.dml);
const NAME_EXT = qname('a', 'ext', NS.dml);
const NAME_GRAPHIC = qname('a', 'graphic', NS.dml);
const NAME_GRAPHIC_DATA = qname('a', 'graphicData', NS.dml);
const NAME_C_CHART = qname('c', 'chart', NS.chart);

const buildChartGraphicFrame = (opts: {
  id: number;
  name: string;
  x: Emu;
  y: Emu;
  w: Emu;
  h: Emu;
  rEmbed: string;
}): XmlElement => {
  const cNvPr = elem(NAME_C_NV_PR_FN, {
    attrs: [attr(qname('', 'id', ''), String(opts.id)), attr(qname('', 'name', ''), opts.name)],
  });
  const nvGraphicFramePr = elem(NAME_NV_GRAPHIC_FRAME_PR, {
    children: [cNvPr, elem(NAME_C_NV_GRAPHIC_FRAME_PR), elem(NAME_NV_PR)],
  });
  const off = elem(NAME_OFF, {
    attrs: [attr(qname('', 'x', ''), String(opts.x)), attr(qname('', 'y', ''), String(opts.y))],
  });
  const ext = elem(NAME_EXT, {
    attrs: [attr(qname('', 'cx', ''), String(opts.w)), attr(qname('', 'cy', ''), String(opts.h))],
  });
  const xfrm = elem(NAME_XFRM, { children: [off, ext] });
  const chartRef = elem(NAME_C_CHART, {
    prefixDecls: new Map([
      ['c', NS.chart],
      ['r', NS.officeDocRels],
    ]),
    attrs: [attr(qname('r', 'id', NS.officeDocRels), opts.rEmbed)],
  });
  const graphicData = elem(NAME_GRAPHIC_DATA, {
    attrs: [attr(qname('', 'uri', ''), NS.chart)],
    children: [chartRef],
  });
  const graphic = elem(NAME_GRAPHIC, { children: [graphicData] });
  return elem(NAME_GRAPHIC_FRAME, { children: [nvGraphicFramePr, xfrm, graphic] });
};

/**
 * Adds a chart to the slide. Returns the new shape handle (kind
 * `graphicFrame`). Supported chart kinds today: `bar`, `column`,
 * `line`, `pie` — see `ChartSpec.kind`.
 *
 * Side effects:
 *
 *   - Allocates `/ppt/charts/chart{N}.xml` for the chart definition.
 *   - Allocates `/ppt/embeddings/Microsoft_Excel_Worksheet{N}.xlsx` as
 *     a placeholder workbook (single sheet, header row + one row per
 *     category). PowerPoint reads the inline `<c:strCache>` /
 *     `<c:numCache>` so the workbook is for "Edit data" only.
 *   - Slide → chart and chart → workbook rels are wired with fresh rIds.
 *   - `<a:graphicFrame>` is appended to the slide's `<p:spTree>`.
 *
 * Constraints:
 *
 *   - `pie` charts require exactly one series.
 *   - All series should have at most `categories.length` values; missing
 *     values are treated as blanks (gaps in the visualization).
 */
export const addSlideChart = (
  slide: SlideData,
  opts: {
    spec: ChartSpec;
    x: Emu;
    y: Emu;
    w: Emu;
    h: Emu;
    name?: string;
  },
): SlideShapeData => {
  const pkg = slide[INTERNAL_PACKAGE];
  const chartN = allocateChartIndex(pkg);
  const chartPartName = partName(`/ppt/charts/chart${chartN}.xml`);
  const xlsxPartName = partName(`/ppt/embeddings/Microsoft_Excel_Worksheet${chartN}.xlsx`);

  // Build the embedded xlsx bytes. Each row in the sheet corresponds to
  // one category; header row carries the series names.
  const xlsxRows = opts.spec.categories.map((label, i) => ({
    label,
    values: opts.spec.series.map((s) => s.values[i] ?? null),
  }));
  const xlsxBytes = buildEmbeddedXlsx(
    opts.spec.series.map((s) => s.name),
    xlsxRows,
  );

  // Build the chart XML and serialize.
  const chartDoc = buildChartSpaceDoc(opts.spec);
  const chartBytes = encode(serializeXml(chartDoc));

  // Add the chart part + its rel → embedded xlsx.
  pkg.addPart(chartPartName, CHART_CONTENT_TYPE, chartBytes);

  // The xlsx is a binary part; xlsx is already an OPC zip so we add a
  // Content_Types override (no Default, since `.xlsx` shouldn't override
  // unrelated archive entries even though there's only one such part
  // here in practice).
  pkg.addPart(xlsxPartName, EMBEDDED_XLSX_CONTENT_TYPE, xlsxBytes);

  // Make sure `.rels` is a recognized Default (it always is by the time
  // we get here, but be defensive for new packages).
  setOpcDefault(pkg, 'rels', 'application/vnd.openxmlformats-package.relationships+xml');

  const chartRels = emptyRels();
  chartRels.items.push({
    id: 'rId1',
    type: REL_TYPES.package,
    target: `../embeddings/Microsoft_Excel_Worksheet${chartN}.xlsx`,
    targetMode: 'Internal',
  });
  pkg.setRels(chartPartName, chartRels);

  // Slide → chart rel.
  const slideRels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
  const slideChartRId = nextRelId(slideRels.items.map((r) => r.id));
  slideRels.items.push({
    id: slideChartRId,
    type: REL_TYPES.chart,
    target: `../charts/chart${chartN}.xml`,
    targetMode: 'Internal',
  });
  pkg.setRels(slide[SLIDE_PART_NAME], slideRels);

  // Build and append the <p:graphicFrame> wrapper.
  const frame = buildChartGraphicFrame({
    id: nextShapeId(slide),
    name: opts.name ?? `Chart ${chartN}`,
    x: opts.x,
    y: opts.y,
    w: opts.w,
    h: opts.h,
    rEmbed: slideChartRId,
  });
  return appendAndReturnNewShape(slide, frame);
};

// Re-export chart types for consumers.
export type { ChartKind, ChartSeries, ChartSpec };

/**
 * A chart sitting on a slide. `shape` is the `<p:graphicFrame>`
 * wrapper; `spec` is the chart definition parsed from the linked
 * `/ppt/charts/chart{N}.xml` part. `null` `spec` means the chart uses
 * a kind we don't model (callers can fall through to pass-through).
 */
export interface SlideChartData {
  readonly shape: SlideShapeData;
  readonly spec: ChartSpec | null;
}

const NAME_A_GRAPHIC_FN = qname('a', 'graphic', NS.dml);
const NAME_A_GRAPHIC_DATA_FN = qname('a', 'graphicData', NS.dml);
const NAME_C_CHART_FN = qname('c', 'chart', NS.chart);

/**
 * Resolves the chart part backing a graphic-frame shape, or `null` if
 * the shape isn't a chart wrapper.
 */
export const resolveChartPartName = (
  slide: SlideData,
  shape: SlideShapeData,
): { partName: PartName; rId: string } | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'graphicFrame') return null;
  const graphic = firstChildElement(shape[SHAPE_ELEMENT], NAME_A_GRAPHIC_FN);
  if (!graphic) return null;
  const graphicData = firstChildElement(graphic, NAME_A_GRAPHIC_DATA_FN);
  if (!graphicData) return null;
  const chartRef = firstChildElement(graphicData, NAME_C_CHART_FN);
  if (!chartRef) return null;
  const rId = getAttrValue(chartRef, qname('r', 'id', NS.officeDocRels));
  if (rId === null) return null;
  const slideRels = slide[INTERNAL_PACKAGE].getRels(slide[SLIDE_PART_NAME]);
  if (!slideRels) return null;
  const rel = slideRels.items.find((r) => r.id === rId);
  if (!rel) return null;
  const partNameValue = rel.target.startsWith('/')
    ? partName(rel.target)
    : resolveTarget(slide[SLIDE_PART_NAME], rel.target);
  return { partName: partNameValue, rId };
};

/**
 * Replaces the chart definition on an existing graphic-frame chart
 * shape. Updates the inline `<c:strCache>` / `<c:numCache>` blocks so
 * PowerPoint renders the new data without opening the embedded
 * workbook. The shape's geometry (position / size / rotation) is
 * preserved verbatim.
 *
 * The embedded xlsx is re-written too — it's what the "Edit data"
 * affordance opens. The previous workbook is replaced wholesale (no
 * attempt to preserve styles a user added through Excel).
 *
 * Pass any `ChartSpec`, including a different `kind` from the
 * original; this acts as "change my column chart to a line chart with
 * fresh data."
 */
export const setChartSpec = (chart: SlideChartData, spec: ChartSpec): void => {
  const slide = chart.shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];
  const resolved = resolveChartPartName(slide, chart.shape);
  if (!resolved) {
    throw new Error('setChartSpec: shape is not a chart graphic frame');
  }

  // Rewrite the chart XML.
  const doc = buildChartSpaceDoc(spec);
  const chartBytes = encode(serializeXml(doc));
  const chartPart = pkg.getPart(resolved.partName);
  if (!chartPart) {
    throw new Error(`setChartSpec: chart part ${resolved.partName} not found`);
  }
  chartPart.data = chartBytes;

  // Rewrite the embedded xlsx (the part chart→package rel points at).
  // Reuse the existing rel; fall back to creating a fresh xlsx part if
  // the chart had no package rel (unusual for charts we authored).
  const chartRels = pkg.getRels(resolved.partName);
  if (chartRels) {
    const xlsxRel = chartRels.items.find((r) => r.type === REL_TYPES.package);
    if (xlsxRel) {
      const xlsxName = xlsxRel.target.startsWith('/')
        ? partName(xlsxRel.target)
        : resolveTarget(resolved.partName, xlsxRel.target);
      const xlsxPart = pkg.getPart(xlsxName);
      const rows = spec.categories.map((label, i) => ({
        label,
        values: spec.series.map((s) => s.values[i] ?? null),
      }));
      const xlsxBytes = buildEmbeddedXlsx(
        spec.series.map((s) => s.name),
        rows,
      );
      if (xlsxPart) {
        xlsxPart.data = xlsxBytes;
      }
    }
  }
};

/**
 * Returns every chart on the slide, with its `ChartSpec` parsed from
 * the linked chart part. Skips graphic frames that don't carry a
 * `<c:chart>` reference (e.g. tables, diagrams).
 */
/**
 * For a graphic-frame shape that wraps a chart, returns the parsed
 * `ChartSpec`. Returns `null` when the shape isn't a chart wrapper
 * or the chart uses a kind we don't model (e.g. surface, radar).
 *
 * Convenience over `getSlideCharts(...).find((c) => c.shape === shape)`
 * when the caller already has the shape in hand (e.g. iterating
 * `getSlideShapes`).
 */
/**
 * Convenience over `getShapeChartSpec(shape)?.kind ?? null`. Returns
 * the chart's `ChartKind` ('bar', 'line', 'pie', …) when the shape
 * is a chart wrapper, or `null` for non-charts and charts whose
 * kind isn't modeled yet.
 */
export const getShapeChartKind = (shape: SlideShapeData): ChartKind | null => {
  const spec = getShapeChartSpec(shape);
  return spec === null ? null : spec.kind;
};

/**
 * Returns the categories axis labels of a chart shape, or `null`
 * if the shape isn't a chart wrapper or its kind isn't modeled.
 * Convenience over `getShapeChartSpec(shape)?.categories ?? null`.
 */
export const getShapeChartCategories = (shape: SlideShapeData): ReadonlyArray<string> | null => {
  const spec = getShapeChartSpec(shape);
  return spec === null ? null : spec.categories;
};

/**
 * Returns the chart's series-name list (in spec order). `null`
 * when the shape isn't a chart wrapper or the kind isn't modeled.
 */
export const getShapeChartSeriesNames = (shape: SlideShapeData): ReadonlyArray<string> | null => {
  const spec = getShapeChartSpec(shape);
  return spec === null ? null : spec.series.map((s) => s.name);
};

/**
 * Returns the values for the named series on a chart shape, or
 * `null` when the shape isn't a chart, the kind isn't modeled, or
 * no series matches `seriesName`.
 */
export const getShapeChartSeriesValues = (
  shape: SlideShapeData,
  seriesName: string,
): ReadonlyArray<number | null> | null => {
  const spec = getShapeChartSpec(shape);
  if (spec === null) return null;
  const series = spec.series.find((s) => s.name === seriesName);
  return series ? series.values : null;
};

export const getShapeChartSpec = (shape: SlideShapeData): ChartSpec | null => {
  const slide = shape[SHAPE_SLIDE];
  const resolved = resolveChartPartName(slide, shape);
  if (!resolved) return null;
  const part = slide[INTERNAL_PACKAGE].getPart(resolved.partName);
  if (!part) return null;
  try {
    const root = parseXml(decode(part.data)).root;
    return readChartSpec(root);
  } catch {
    return null;
  }
};

/**
 * Returns every chart on the slide that carries a series whose name
 * equals `seriesName` exactly. Useful for "find the revenue chart"
 * patterns where chart kind alone isn't unique. Skips charts whose
 * kind isn't modeled.
 */
export const findChartsBySeriesName = (
  slide: SlideData,
  seriesName: string,
): ReadonlyArray<SlideChartData> => {
  const out: SlideChartData[] = [];
  for (const chart of getSlideCharts(slide)) {
    if (chart.spec === null) continue;
    if (chart.spec.series.some((s) => s.name === seriesName)) out.push(chart);
  }
  return out;
};

/**
 * Returns the first chart on the slide whose parsed `kind` matches
 * `kind` (e.g. `'bar'`, `'line'`, `'pie'`). Returns `null` when no
 * chart on the slide has that kind, or when every chart on the slide
 * uses a kind this version doesn't yet model.
 */
export const findChartByKind = (slide: SlideData, kind: ChartKind): SlideChartData | null => {
  for (const chart of getSlideCharts(slide)) {
    if (chart.spec !== null && chart.spec.kind === kind) return chart;
  }
  return null;
};

export const getSlideCharts = (slide: SlideData): ReadonlyArray<SlideChartData> => {
  const pkg = slide[INTERNAL_PACKAGE];
  const out: SlideChartData[] = [];
  for (const shape of slide[SLIDE_SHAPES]) {
    const resolved = resolveChartPartName(slide, shape);
    if (!resolved) continue;
    const chartPart = pkg.getPart(resolved.partName);
    if (!chartPart) continue;
    let spec: ChartSpec | null;
    try {
      const root = parseXml(decode(chartPart.data)).root;
      spec = readChartSpec(root);
    } catch {
      spec = null;
    }
    out.push({ shape, spec });
  }
  return out;
};

void textNode;

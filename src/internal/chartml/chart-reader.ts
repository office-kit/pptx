// Chart spec reader — parses `<c:chartSpace>` back into a `ChartSpec`.
//
// Companion to `chart-builder.ts`. Lets callers introspect or edit
// chart data on existing templates without dropping to XML. The reader
// uses inline `<c:strCache>` / `<c:numCache>` blocks rather than the
// embedded workbook — the cache is what PowerPoint renders from
// anyway, and it's always present in charts @office-kit/pptx emits.

import {
  NS,
  type XmlElement,
  allChildElements,
  firstChildElement,
  getAttrValue,
  qname,
} from '../xml/index.ts';
import type {
  ChartAxisScaling,
  ChartDataLabelPosition,
  ChartDataLabels,
  ChartGrouping,
  ChartKind,
  ChartSecondaryValueAxis,
  ChartSeries,
  ChartSpec,
  ChartTextStyle,
  ChartTrendline,
} from './types.ts';

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
const NAME_MULTI_LVL_STR_REF = qname('c', 'multiLvlStrRef', NS_C);
const NAME_MULTI_LVL_STR_CACHE = qname('c', 'multiLvlStrCache', NS_C);
const NAME_LVL = qname('c', 'lvl', NS_C);
const NAME_PT = qname('c', 'pt', NS_C);
const NAME_PT_COUNT = qname('c', 'ptCount', NS_C);
// Excel's row limit — the most points a chart cache can legitimately hold.
const MAX_PT_COUNT = 1_048_576;
const NAME_V = qname('c', 'v', NS_C);
const NAME_D_LBLS = qname('c', 'dLbls', NS_C);
const NAME_D_LBL = qname('c', 'dLbl', NS_C);
const NAME_VAL_AX = qname('c', 'valAx', NS_C);
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

/** Kinds a combo plot group can carry (mirrors `ChartSeries.chartKind`). */
const isComboSeriesKind = (kind: ChartKind): kind is 'bar' | 'column' | 'line' | 'area' =>
  kind === 'bar' || kind === 'column' || kind === 'line' || kind === 'area';

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
  // Scatter / bubble carry xy / xyz tuples per series (`<c:xVal>` /
  // `<c:yVal>` / `<c:bubbleSize>`) rather than numeric channels against
  // shared categories; radar uses cat+val like line. All three are
  // modeled as their own kind (read + render only — the builder rejects
  // them, see chart-builder.ts).
  { localName: 'scatterChart', kind: 'scatter' },
  { localName: 'bubbleChart', kind: 'bubble' },
  { localName: 'radarChart', kind: 'radar' },
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

// `<c:ptCount val>` of a cache / literal; NaN when absent or malformed.
const readPtCount = (host: XmlElement): number => {
  const ptCountEl = firstChildElement(host, NAME_PT_COUNT);
  const raw = ptCountEl !== null ? getAttrValue(ptCountEl, ATTR_VAL) : null;
  return raw !== null ? Number.parseInt(raw, 10) : Number.NaN;
};

// CT_Boolean: an absent `val` means true (schema default); PowerPoint writes
// "0" / "1", other writers spell "true" / "false".
const readXmlBool = (el: XmlElement): boolean => {
  const v = getAttrValue(el, ATTR_VAL);
  return v === null || v === '1' || v === 'true';
};

// A point index (`idx`) from untrusted input: an unsigned integer below
// `limit`, else null, so a huge idx never allocates a huge sparse array.
const parsePointIndex = (raw: string | null, limit: number): number | null => {
  if (raw === null || !/^\d+$/.test(raw)) return null;
  const idx = Number.parseInt(raw, 10);
  return idx < limit ? idx : null;
};

// `<c:idx val>` of a `<c:dPt>` / `<c:dLbl>`, bounded by the series' point count.
const readPointIndex = (host: XmlElement, pointCount: number): number | null => {
  const idxEl = firstChildElement(host, qname('c', 'idx', NS_C));
  return parsePointIndex(idxEl !== null ? getAttrValue(idxEl, ATTR_VAL) : null, pointCount);
};

// `countHost` holds the `<c:ptCount>` when it is not `cache` itself: a
// `<c:multiLvlStrCache>` counts once for every `<c:lvl>` under it.
const readPtArray = (cache: XmlElement, countHost: XmlElement = cache): string[] => {
  const out: string[] = [];
  const ptCount = readPtCount(countHost);
  // `<c:pt idx>` past the authored count (or the hard cap) is dropped.
  const idxLimit = Number.isFinite(ptCount) ? Math.min(ptCount, MAX_PT_COUNT) : MAX_PT_COUNT;
  for (const pt of allChildElements(cache, NAME_PT)) {
    const idx = parsePointIndex(getAttrValue(pt, ATTR_IDX), idxLimit);
    if (idx === null) continue;
    const v = firstChildElement(pt, NAME_V);
    if (v === null) continue;
    let text = '';
    for (const child of v.children) {
      if (child.kind === 'text' || child.kind === 'cdata') text += child.data;
    }
    out[idx] = text;
  }
  // Writers (this builder included) omit `<c:pt>` for empty cells, so the
  // authored point count — not the last present idx — is the array length.
  // Untrusted input: a huge ptCount must not allocate a huge array.
  if (Number.isFinite(ptCount) && ptCount <= MAX_PT_COUNT) {
    for (let i = 0; i < ptCount; i++) {
      if (out[i] === undefined) out[i] = '';
    }
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
  // `<c:multiLvlStrRef>` is what PowerPoint and PptxGenJS emit for categories
  // (it supports grouped axis labels). A flat chart has a single `<c:lvl>`;
  // read that level's points so single-level categories round-trip. (For a
  // genuinely multi-level axis we surface the innermost level, which is the
  // per-point label.)
  const multi = firstChildElement(parent, NAME_MULTI_LVL_STR_REF);
  if (multi) {
    const cache = firstChildElement(multi, NAME_MULTI_LVL_STR_CACHE);
    const lvl = cache ? firstChildElement(cache, NAME_LVL) : null;
    if (lvl && cache) return readPtArray(lvl, cache);
  }
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

// Walks `<c:ser><c:dPt>` overrides and returns sparse per-point arrays
// for color (`<c:spPr><a:solidFill><a:srgbClr>`) and explosion
// (`<c:explosion val="N"/>`). Returns `undefined` for each side when no
// dPt authors the corresponding attribute.
const readDataPointOverrides = (
  ser: XmlElement,
  pointCount: number,
): {
  readonly colors: ReadonlyArray<string | null> | undefined;
  readonly explosions: ReadonlyArray<number | null> | undefined;
} => {
  const colors: Array<string | null> = [];
  const explosions: Array<number | null> = [];
  let anyColor = false;
  let anyExplosion = false;
  for (const c of ser.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS_C || c.name.localName !== 'dPt')
      continue;
    const idx = readPointIndex(c, pointCount);
    if (idx === null) continue;
    const spPr = firstChildElement(c, NAME_SP_PR_C);
    if (spPr) {
      const solidFill = firstChildElement(spPr, NAME_SOLID_FILL);
      if (solidFill) {
        const srgb = firstChildElement(solidFill, NAME_SRGB_CLR);
        if (srgb) {
          const val = getAttrValue(srgb, ATTR_VAL);
          if (val !== null) {
            colors[idx] = `#${val.toUpperCase()}`;
            anyColor = true;
          }
        }
      }
    }
    const explEl = firstChildElement(c, qname('c', 'explosion', NS_C));
    if (explEl) {
      const v = getAttrValue(explEl, ATTR_VAL);
      if (v !== null) {
        const n = Number.parseInt(v, 10);
        if (Number.isFinite(n) && n > 0) {
          explosions[idx] = n;
          anyExplosion = true;
        }
      }
    }
  }
  return {
    colors: anyColor ? colors : undefined,
    explosions: anyExplosion ? explosions : undefined,
  };
};

// `<c:trendline>` is a sibling of `<c:val>` inside `<c:ser>`. Read the
// type token and (where relevant) the period / polynomial order. Color
// comes from `<c:trendline><c:spPr><a:ln><a:solidFill>` when authored.
const readTrendline = (ser: XmlElement): ChartTrendline | undefined => {
  const tl = firstChildElement(ser, qname('c', 'trendline', NS_C));
  if (!tl) return undefined;
  // <c:trendline><c:name>…</c:name> — optional custom label.
  let name: string | undefined;
  const nameEl = firstChildElement(tl, qname('c', 'name', NS_C));
  if (nameEl) {
    let acc = '';
    for (const child of nameEl.children) {
      if (child.kind === 'text' || child.kind === 'cdata') acc += child.data;
    }
    if (acc.length > 0) name = acc;
  }
  const typeEl = firstChildElement(tl, qname('c', 'trendlineType', NS_C));
  const tToken = typeEl ? getAttrValue(typeEl, ATTR_VAL) : null;
  let type: ChartTrendline['type'];
  switch (tToken) {
    case 'exp':
    case 'log':
    case 'poly':
    case 'power':
    case 'movingAvg':
    case 'linear':
      type = tToken;
      break;
    default:
      type = 'linear';
  }
  let period: number | undefined;
  if (type === 'movingAvg') {
    const pEl = firstChildElement(tl, qname('c', 'period', NS_C));
    if (pEl) {
      const v = getAttrValue(pEl, ATTR_VAL);
      if (v !== null) {
        const n = Number.parseInt(v, 10);
        if (Number.isFinite(n) && n > 0) period = n;
      }
    }
  }
  let order: number | undefined;
  if (type === 'poly') {
    const oEl = firstChildElement(tl, qname('c', 'order', NS_C));
    if (oEl) {
      const v = getAttrValue(oEl, ATTR_VAL);
      if (v !== null) {
        const n = Number.parseInt(v, 10);
        if (Number.isFinite(n) && n >= 2) order = n;
      }
    }
  }
  // Line color via <c:spPr><a:ln><a:solidFill><a:srgbClr>.
  let color: string | undefined;
  const spPr = firstChildElement(tl, NAME_SP_PR_C);
  if (spPr) {
    const ln = firstChildElement(spPr, qname('a', 'ln', NS_A));
    if (ln) {
      const solid = firstChildElement(ln, NAME_SOLID_FILL);
      if (solid) {
        const srgb = firstChildElement(solid, NAME_SRGB_CLR);
        if (srgb) {
          const v = getAttrValue(srgb, ATTR_VAL);
          if (v !== null) color = `#${v.toUpperCase()}`;
        }
      }
    }
  }
  // <c:forward val="N"/> / <c:backward val="N"/> extend the trendline
  // N periods past the last / before the first data point.
  const readExtension = (local: string): number | undefined => {
    const el = firstChildElement(tl, qname('c', local, NS_C));
    if (!el) return undefined;
    const raw = getAttrValue(el, ATTR_VAL);
    if (raw === null) return undefined;
    const n = Number.parseFloat(raw);
    return Number.isFinite(n) && n > 0 ? n : undefined;
  };
  const forward = readExtension('forward');
  const backward = readExtension('backward');
  // <c:dispEq val="1"/> and <c:dispRSqr val="1"/> show the regression
  // equation / R² coefficient next to the line. Default false.
  const readDispBool = (local: string): boolean | undefined => {
    const el = firstChildElement(tl, qname('c', local, NS_C));
    if (!el) return undefined;
    const v = getAttrValue(el, ATTR_VAL);
    return v === null || v === '1' || v === 'true' ? true : undefined;
  };
  const displayEquation = readDispBool('dispEq');
  const displayRSquared = readDispBool('dispRSqr');
  return {
    ...(name !== undefined ? { name } : {}),
    type,
    ...(period !== undefined ? { period } : {}),
    ...(order !== undefined ? { order } : {}),
    ...(color !== undefined ? { color } : {}),
    ...(forward !== undefined ? { forward } : {}),
    ...(backward !== undefined ? { backward } : {}),
    ...(displayEquation === true ? { displayEquation: true } : {}),
    ...(displayRSquared === true ? { displayRSquared: true } : {}),
  };
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

const readSolidSrgb = (host: XmlElement | null): string | undefined => {
  const fill = host !== null ? firstChildElement(host, NAME_SOLID_FILL) : null;
  const srgb = fill !== null ? firstChildElement(fill, NAME_SRGB_CLR) : null;
  const color = srgb !== null ? getAttrValue(srgb, ATTR_VAL) : null;
  return color !== null ? `#${color.toUpperCase()}` : undefined;
};

// Per-series marker symbol, size + fill / outline color from <c:ser><c:marker>.
const readSeriesMarker = (
  ser: XmlElement,
): {
  markerSymbol?: ChartSeries['markerSymbol'];
  markerSizePt?: number;
  markerColor?: string;
  markerLineColor?: string;
} => {
  const m = firstChildElement(ser, qname('c', 'marker', NS_C));
  if (!m) return {};
  const out: {
    markerSymbol?: ChartSeries['markerSymbol'];
    markerSizePt?: number;
    markerColor?: string;
    markerLineColor?: string;
  } = {};
  const spPr = firstChildElement(m, NAME_SP_PR_C);
  const color = readSolidSrgb(spPr);
  if (color !== undefined) out.markerColor = color;
  const lineColor = readSolidSrgb(
    spPr !== null ? firstChildElement(spPr, qname('a', 'ln', NS_A)) : null,
  );
  if (lineColor !== undefined) out.markerLineColor = lineColor;
  const symEl = firstChildElement(m, qname('c', 'symbol', NS_C));
  if (symEl) {
    const v = getAttrValue(symEl, ATTR_VAL);
    if (
      v === 'none' ||
      v === 'auto' ||
      v === 'circle' ||
      v === 'square' ||
      v === 'diamond' ||
      v === 'triangle' ||
      v === 'star' ||
      v === 'x' ||
      v === 'plus' ||
      v === 'dash' ||
      v === 'dot' ||
      v === 'picture'
    ) {
      out.markerSymbol = v;
    }
  }
  const sizeEl = firstChildElement(m, qname('c', 'size', NS_C));
  if (sizeEl) {
    const v = getAttrValue(sizeEl, ATTR_VAL);
    if (v !== null) {
      const n = Number.parseInt(v, 10);
      if (Number.isFinite(n) && n > 0) out.markerSizePt = n;
    }
  }
  return out;
};

// Per-series line stroke width, color + dash from <c:ser><c:spPr><a:ln>.
const readSeriesLineProps = (
  ser: XmlElement,
): { lineWidthEmu?: number; lineColor?: string; lineDash?: string } => {
  const spPr = firstChildElement(ser, NAME_SP_PR_C);
  if (!spPr) return {};
  const ln = firstChildElement(spPr, qname('a', 'ln', NS_A));
  if (!ln) return {};
  const out: { lineWidthEmu?: number; lineColor?: string; lineDash?: string } = {};
  const lnFill = firstChildElement(ln, NAME_SOLID_FILL);
  const lnSrgb = lnFill !== null ? firstChildElement(lnFill, NAME_SRGB_CLR) : null;
  const lnColor = lnSrgb !== null ? getAttrValue(lnSrgb, ATTR_VAL) : null;
  if (lnColor !== null) out.lineColor = `#${lnColor.toUpperCase()}`;
  const w = getAttrValue(ln, qname('', 'w', ''));
  if (w !== null) {
    const n = Number.parseInt(w, 10);
    // 0 is a valid ST_LineWidth and the builder writes it, so it reads back.
    if (Number.isFinite(n) && n >= 0) out.lineWidthEmu = n;
  }
  const prstDash = firstChildElement(ln, qname('a', 'prstDash', NS_A));
  if (prstDash) {
    const v = getAttrValue(prstDash, ATTR_VAL);
    if (v !== null) out.lineDash = v;
  }
  return out;
};

const readTitle = (chart: XmlElement): string | undefined => {
  const title = firstChildElement(chart, NAME_TITLE);
  if (!title) return undefined;
  const tx = firstChildElement(title, NAME_TX);
  if (!tx) return undefined;
  // `<c:tx>` may carry either `<c:rich>` (literal) or `<c:strRef>` (cell
  // reference, with its `<c:strCache>` holding the resolved string).
  // Authors who type the title directly emit rich; chart wizards that
  // wire the title to a cell emit strRef. Accept both.
  const rich = firstChildElement(tx, NAME_RICH);
  if (rich) {
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
    if (acc.length > 0) return acc;
  }
  const cells = readStringRef(tx);
  if (cells && cells.length > 0) {
    const text = cells.filter((s) => s.length > 0).join(' ');
    if (text.length > 0) return text;
  }
  return undefined;
};

const NAME_A_RPR = qname('a', 'rPr', NS_A);
const NAME_A_DEF_RPR = qname('a', 'defRPr', NS_A);
const NAME_A_PPR = qname('a', 'pPr', NS_A);
const NAME_A_SRGB = qname('a', 'srgbClr', NS_A);
const NAME_A_LATIN = qname('a', 'latin', NS_A);
const ATTR_TYPEFACE = qname('', 'typeface', '');

// Reads `<a:rPr>` / `<a:defRPr>` attributes (size in 100ths of a pt,
// bold / italic) plus the first solidFill color inside it. Returns an
// undefined-only style when nothing is authored — callers should drop
// the style entirely in that case.
const readRunStyle = (rPr: XmlElement): ChartTextStyle | undefined => {
  let font: string | undefined;
  let sizePt: number | undefined;
  let bold: boolean | undefined;
  let italic: boolean | undefined;
  let color: string | undefined;
  const szRaw = getAttrValue(rPr, qname('', 'sz', ''));
  if (szRaw !== null) {
    const n = Number.parseInt(szRaw, 10);
    if (Number.isFinite(n) && n > 0) sizePt = n / 100;
  }
  const bRaw = getAttrValue(rPr, qname('', 'b', ''));
  if (bRaw !== null) bold = bRaw === '1' || bRaw === 'true';
  const iRaw = getAttrValue(rPr, qname('', 'i', ''));
  if (iRaw !== null) italic = iRaw === '1' || iRaw === 'true';
  const solidFill = firstChildElement(rPr, NAME_SOLID_FILL);
  if (solidFill) {
    const srgb = firstChildElement(solidFill, NAME_A_SRGB);
    if (srgb) {
      const v = getAttrValue(srgb, ATTR_VAL);
      if (v !== null) color = `#${v.toUpperCase()}`;
    }
  }
  // The latin slot is the authoritative font face on the round-trip; the
  // ea slot carries the same value (written by the builder) so reading
  // latin alone recovers the authored `font`.
  const latin = firstChildElement(rPr, NAME_A_LATIN);
  if (latin) {
    const tf = getAttrValue(latin, ATTR_TYPEFACE);
    if (tf !== null && tf !== '') font = tf;
  }
  if (
    font === undefined &&
    sizePt === undefined &&
    bold === undefined &&
    italic === undefined &&
    color === undefined
  ) {
    return undefined;
  }
  return {
    ...(font !== undefined ? { font } : {}),
    ...(sizePt !== undefined ? { sizePt } : {}),
    ...(bold !== undefined ? { bold } : {}),
    ...(italic !== undefined ? { italic } : {}),
    ...(color !== undefined ? { color } : {}),
  };
};

// Reads the authored text style of a chart label's `<c:rich>` / `<c:txPr>`
// block: the first styled run's `<a:rPr>` over the paragraph's `<a:defRPr>`
// (a run attribute wins, the rest falls back to the paragraph defaults, as
// renderers resolve it). Returns `undefined` when nothing is authored.
const readLabelStyle = (richHost: XmlElement): ChartTextStyle | undefined => {
  for (const p of allChildElements(richHost, NAME_P_DML)) {
    let runStyle: ChartTextStyle | undefined;
    for (const r of allChildElements(p, NAME_R_DML)) {
      const rPr = firstChildElement(r, NAME_A_RPR);
      runStyle = rPr !== null ? readRunStyle(rPr) : undefined;
      if (runStyle !== undefined) break;
    }
    const pPr = firstChildElement(p, NAME_A_PPR);
    const defRPr = pPr !== null ? firstChildElement(pPr, NAME_A_DEF_RPR) : null;
    const defStyle = defRPr !== null ? readRunStyle(defRPr) : undefined;
    if (runStyle === undefined && defStyle === undefined) continue;
    return { ...defStyle, ...runStyle };
  }
  return undefined;
};

// Extracts the authored text style from a `<c:title>` element. Used for
// both the chart-level title and axis titles, which share the
// `<c:tx><c:rich>` shape.
const readTitleStyleOf = (titleEl: XmlElement): ChartTextStyle | undefined => {
  const tx = firstChildElement(titleEl, NAME_TX);
  if (!tx) return undefined;
  const rich = firstChildElement(tx, NAME_RICH);
  if (!rich) return undefined;
  return readLabelStyle(rich);
};

const readTitleStyle = (chart: XmlElement): ChartTextStyle | undefined => {
  const title = firstChildElement(chart, NAME_TITLE);
  if (!title) return undefined;
  return readTitleStyleOf(title);
};

// Reads `<c:title><c:tx><c:rich><a:bodyPr rot="N"/>` and converts from
// the OOXML 60000ths-of-a-degree unit to plain degrees. Returns
// `undefined` when the rotation is absent (the renderer's default
// applies).
const readTitleRotationDeg = (titleEl: XmlElement): number | undefined => {
  const tx = firstChildElement(titleEl, NAME_TX);
  if (!tx) return undefined;
  const rich = firstChildElement(tx, NAME_RICH);
  if (!rich) return undefined;
  const bodyPr = firstChildElement(rich, qname('a', 'bodyPr', NS_A));
  if (!bodyPr) return undefined;
  const v = getAttrValue(bodyPr, qname('', 'rot', ''));
  if (v === null) return undefined;
  const n = Number.parseInt(v, 10);
  return Number.isFinite(n) ? n / 60000 : undefined;
};

// `<c:dLbls><c:dLblPos val="…"/>` — the chart-kind-dependent enum that
// names where labels sit relative to their data point. Returns
// `undefined` for unknown tokens so callers fall back to their default.
const readDataLabelPosition = (dLbls: XmlElement): ChartDataLabelPosition | undefined => {
  const el = firstChildElement(dLbls, qname('c', 'dLblPos', NS_C));
  if (!el) return undefined;
  const v = getAttrValue(el, ATTR_VAL);
  switch (v) {
    case 'ctr':
    case 'inEnd':
    case 'outEnd':
    case 'inBase':
    case 't':
    case 'b':
    case 'l':
    case 'r':
    case 'bestFit':
      return v;
    default:
      return undefined;
  }
};

// `<c:dLbls><c:separator>…</c:separator>` — leaf text content; common
// values are `" "`, `", "`, `"\n"`, `"; "`.
const readDataLabelSeparator = (dLbls: XmlElement): string | undefined => {
  const el = firstChildElement(dLbls, qname('c', 'separator', NS_C));
  if (!el) return undefined;
  let acc = '';
  for (const c of el.children) {
    if (c.kind === 'text' || c.kind === 'cdata') acc += c.data;
  }
  return acc.length > 0 ? acc : undefined;
};

// The shared tail of CT_DLbls / CT_DLbl (Group_DLbls / Group_DLbl): the
// show* toggles plus optional numFmt, position, separator, txPr.
// A toggle the host leaves out takes the `inherited` group's value (a
// per-point `<c:dLbl>` inherits from its series `<c:dLbls>`), else false.
const readDataLabelsGroup = (
  host: XmlElement,
  inherited: ChartDataLabels | undefined,
): ChartDataLabels => {
  const readToggle = (local: string, fallback: boolean): boolean => {
    const el = firstChildElement(host, qname('c', local, NS_C));
    return el !== null ? readXmlBool(el) : fallback;
  };
  const nfEl = firstChildElement(host, qname('c', 'numFmt', NS_C));
  let numberFormat: string | undefined;
  if (nfEl) {
    const fc = getAttrValue(nfEl, qname('', 'formatCode', ''));
    if (fc !== null && fc.length > 0 && fc !== 'General') numberFormat = fc;
  }
  const position = readDataLabelPosition(host);
  const separator = readDataLabelSeparator(host);
  const txPrEl = firstChildElement(host, qname('c', 'txPr', NS_C));
  const textStyle = txPrEl ? readLabelStyle(txPrEl) : undefined;
  return {
    showValue: readToggle('showVal', inherited?.showValue ?? false),
    showCategory: readToggle('showCatName', inherited?.showCategory ?? false),
    showSeriesName: readToggle('showSerName', inherited?.showSeriesName ?? false),
    showPercent: readToggle('showPercent', inherited?.showPercent ?? false),
    ...(numberFormat !== undefined ? { numberFormat } : {}),
    ...(position !== undefined ? { position } : {}),
    ...(separator !== undefined ? { separator } : {}),
    ...(textStyle !== undefined ? { textStyle } : {}),
  };
};

// Series- / chart-level `<c:dLbls>`: the shared group plus
// `<c:showLeaderLines>`, which Group_DLbl (a per-point `<c:dLbl>`) lacks.
const readSeriesLevelDataLabels = (dLbls: XmlElement): ChartDataLabels => {
  const group = readDataLabelsGroup(dLbls, undefined);
  const leaderEl = firstChildElement(dLbls, qname('c', 'showLeaderLines', NS_C));
  return leaderEl !== null ? { ...group, showLeaderLines: readXmlBool(leaderEl) } : group;
};

// True when `<c:dLbls>` carries series-level settings (anything besides the
// per-point `<c:dLbl>` overrides and the extension list).
const hasDataLabelGroupChild = (dLbls: XmlElement): boolean =>
  dLbls.children.some(
    (ch) =>
      ch.kind === 'element' &&
      ch.name.namespaceURI === NS_C &&
      ch.name.localName !== 'dLbl' &&
      ch.name.localName !== 'extLst',
  );

// A deleted point label reads as an override with every toggle off — the
// same shape the builder writes back for it, so it round-trips stably.
const DELETED_POINT_LABEL: ChartDataLabels = {
  showValue: false,
  showCategory: false,
  showSeriesName: false,
  showPercent: false,
};

// `<c:dLbls><c:dLbl><c:idx val="N"/>…` per-point overrides, as a sparse
// array indexed by point (`null` where no override is authored). Toggles a
// point leaves out inherit the series-level group (`seriesLabels`).
const readPointDataLabels = (
  dLbls: XmlElement,
  pointCount: number,
  seriesLabels: ChartDataLabels | undefined,
): ReadonlyArray<ChartDataLabels | null> | undefined => {
  const out: Array<ChartDataLabels | null> = [];
  let any = false;
  for (const dLbl of allChildElements(dLbls, NAME_D_LBL)) {
    const idx = readPointIndex(dLbl, pointCount);
    if (idx === null) continue;
    const deleteEl = firstChildElement(dLbl, qname('c', 'delete', NS_C));
    out[idx] =
      deleteEl !== null && readXmlBool(deleteEl)
        ? DELETED_POINT_LABEL
        : readDataLabelsGroup(dLbl, seriesLabels);
    any = true;
  }
  if (!any) return undefined;
  for (let i = 0; i < out.length; i++) out[i] ??= null;
  return out;
};

// `<c:valAx>` scaling: <c:scaling><c:min/>/<c:max/> as the authored axis
// range, plus optional <c:majorUnit> / <c:minorUnit> tick spacing,
// <c:numFmt>, <c:dispUnits>. Shared by the primary and secondary axes.
const readValueAxisScaling = (valAx: XmlElement): ChartAxisScaling | undefined => {
  let min: number | undefined;
  let max: number | undefined;
  let majorUnit: number | undefined;
  let minorUnit: number | undefined;
  let logBase: number | undefined;
  let displayUnits: ChartAxisScaling['displayUnits'];
  const scaling = firstChildElement(valAx, qname('c', 'scaling', NS_C));
  const readNumOn = (parent: XmlElement, local: string): number | undefined => {
    const el = firstChildElement(parent, qname('c', local, NS_C));
    if (!el) return undefined;
    const v = getAttrValue(el, ATTR_VAL);
    if (v === null) return undefined;
    const n = Number.parseFloat(v);
    return Number.isFinite(n) ? n : undefined;
  };
  if (scaling) {
    min = readNumOn(scaling, 'min');
    max = readNumOn(scaling, 'max');
    // <c:logBase val="N"/> — PowerPoint requires N in [2, 1000].
    const lb = readNumOn(scaling, 'logBase');
    if (lb !== undefined && lb >= 2 && lb <= 1000) logBase = lb;
  }
  majorUnit = readNumOn(valAx, 'majorUnit');
  minorUnit = readNumOn(valAx, 'minorUnit');
  // <c:dispUnits><c:builtInUnit val="hundreds|thousands|…"/>
  const dispUnits = firstChildElement(valAx, qname('c', 'dispUnits', NS_C));
  if (dispUnits) {
    const builtIn = firstChildElement(dispUnits, qname('c', 'builtInUnit', NS_C));
    if (builtIn) {
      const v = getAttrValue(builtIn, ATTR_VAL);
      switch (v) {
        case 'hundreds':
        case 'thousands':
        case 'tenThousands':
        case 'hundredThousands':
        case 'millions':
        case 'tenMillions':
        case 'hundredMillions':
        case 'billions':
        case 'trillions':
          displayUnits = v;
          break;
      }
    }
  }
  // <c:numFmt formatCode="…" sourceLinked="0|1"/> sits directly under
  // <c:valAx>. We surface the formatCode for renderers; sourceLinked
  // (whether to inherit Excel cell format) isn't useful at our layer.
  let numberFormat: string | undefined;
  const nfEl = firstChildElement(valAx, qname('c', 'numFmt', NS_C));
  if (nfEl) {
    const fc = getAttrValue(nfEl, qname('', 'formatCode', ''));
    if (fc !== null && fc.length > 0 && fc !== 'General') {
      numberFormat = fc;
    }
  }
  if (
    min !== undefined ||
    max !== undefined ||
    majorUnit !== undefined ||
    minorUnit !== undefined ||
    numberFormat !== undefined ||
    logBase !== undefined ||
    displayUnits !== undefined
  ) {
    return {
      ...(min !== undefined ? { min } : {}),
      ...(max !== undefined ? { max } : {}),
      ...(majorUnit !== undefined ? { majorUnit } : {}),
      ...(minorUnit !== undefined ? { minorUnit } : {}),
      ...(numberFormat !== undefined ? { numberFormat } : {}),
      ...(logBase !== undefined ? { logBase } : {}),
      ...(displayUnits !== undefined ? { displayUnits } : {}),
    };
  }
  return undefined;
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

  // Collect every "plotted" plot-group element the plotArea carries, in
  // document order. Combo charts emit several (`<c:barChart>` +
  // `<c:lineChart>` …); single-kind charts emit one.
  interface PlotGroup {
    readonly element: XmlElement;
    readonly kind: ChartKind;
  }
  const kindByLocalName = new Map(KIND_MAP.map((entry) => [entry.localName, entry] as const));
  const plotGroups: PlotGroup[] = [];
  for (const child of plotArea.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS_C) continue;
    const mapped = kindByLocalName.get(child.name.localName);
    if (!mapped) continue;
    let groupKind: ChartKind = mapped.kind;
    // Resolve bar vs column on a `barChart` / `bar3DChart`.
    if (mapped.localName === 'barChart' || mapped.localName === 'bar3DChart') {
      const barDir = firstChildElement(child, qname('c', 'barDir', NS_C));
      const v = barDir !== null ? getAttrValue(barDir, ATTR_VAL) : null;
      groupKind = v === 'bar' ? 'bar' : 'column';
    }
    plotGroups.push({ element: child, kind: groupKind });
  }
  const firstGroup = plotGroups[0];
  if (!firstGroup) return null;
  const plotted = firstGroup.element;
  const kind = firstGroup.kind;

  // Secondary-axis detection: the axes the first plot group references are
  // the primary set; a `<c:valAx>` only later groups reference is the
  // secondary value axis. Axis position can't tell them apart — a scatter
  // group references two primary valAx (X at the bottom or top, Y at the
  // left), while pptxgenjs puts a secondary axis at `r` or `t`.
  const groupAxisIds = (group: XmlElement): string[] => {
    const ids: string[] = [];
    for (const axIdEl of allChildElements(group, qname('c', 'axId', NS_C))) {
      const id = getAttrValue(axIdEl, ATTR_VAL);
      if (id !== null) ids.push(id);
    }
    return ids;
  };
  const valAxisIds = new Set<string>();
  for (const el of allChildElements(plotArea, NAME_VAL_AX)) {
    const axIdEl = firstChildElement(el, qname('c', 'axId', NS_C));
    const id = axIdEl !== null ? getAttrValue(axIdEl, ATTR_VAL) : null;
    if (id !== null) valAxisIds.add(id);
  }
  const primaryAxisIds = new Set(groupAxisIds(plotted));
  const secondaryValAxisIds = new Set<string>();
  for (const group of plotGroups.slice(1)) {
    for (const id of groupAxisIds(group.element)) {
      if (valAxisIds.has(id) && !primaryAxisIds.has(id)) secondaryValAxisIds.add(id);
    }
  }
  const groupUsesSecondaryAxis = (group: XmlElement): boolean =>
    groupAxisIds(group).some((id) => secondaryValAxisIds.has(id));

  // Read every <c:ser> from every plot group, tagging series from
  // non-first groups with their group's kind / axis so the round-trip
  // preserves the combo layout.
  const series: ChartSeries[] = [];
  let categoriesFromFirst: string[] | null = null;
  const serEntries: { ser: XmlElement; groupKind: ChartKind; secondary: boolean }[] = [];
  for (const group of plotGroups) {
    const secondary = groupUsesSecondaryAxis(group.element);
    for (const ser of allChildElements(group.element, NAME_SER)) {
      serEntries.push({ ser, groupKind: group.kind, secondary });
    }
  }
  for (const { ser, groupKind, secondary } of serEntries) {
    const name = readSeriesName(ser);
    const cat = firstChildElement(ser, NAME_CAT);
    if (cat !== null && categoriesFromFirst === null) {
      // `<c:cat>` is usually `<c:strRef>` (text categories), but date /
      // numeric categories serialize as `<c:numRef>`. Fall back to the
      // numeric channel formatted as a string so date / number cats
      // still appear on the axis instead of disappearing entirely.
      categoriesFromFirst = readStringRef(cat) ?? null;
      if (categoriesFromFirst === null) {
        const nums = readNumRef(cat);
        if (nums !== null) {
          categoriesFromFirst = nums.map((n) =>
            n === null || !Number.isFinite(n) ? '' : String(n),
          );
        }
      }
    }
    let valEl = firstChildElement(ser, NAME_VAL);
    // Scatter / bubble carry their y-channel on <c:yVal> rather than
    // <c:val>; their x-channel is <c:xVal> and bubble's per-point size
    // is <c:bubbleSize>. Radar / line use <c:val> against shared cats.
    if (!valEl) {
      valEl = firstChildElement(ser, qname('c', 'yVal', NS_C));
    }
    const values = valEl !== null ? readNumRef(valEl) : null;
    const xValEl = firstChildElement(ser, qname('c', 'xVal', NS_C));
    const xValues = xValEl !== null ? readNumRef(xValEl) : null;
    const bubbleSizeEl = firstChildElement(ser, qname('c', 'bubbleSize', NS_C));
    const bubbleSizes = bubbleSizeEl !== null ? readNumRef(bubbleSizeEl) : null;
    const color = readSeriesColor(ser);
    const { lineWidthEmu, lineColor, lineDash } = readSeriesLineProps(ser);
    const { markerSymbol, markerSizePt, markerColor, markerLineColor } = readSeriesMarker(ser);
    const invertEl = firstChildElement(ser, qname('c', 'invertIfNegative', NS_C));
    const invertIfNegative = invertEl !== null && readXmlBool(invertEl) ? true : undefined;
    // Per-point overrides address points of the value channel; anything
    // past its end has nothing to override.
    const pointCount = values !== null ? values.length : 0;
    // <c:dPt> data-point overrides — sparse maps idx → color / explosion.
    const { colors: pointColors, explosions: pointExplosions } = readDataPointOverrides(
      ser,
      pointCount,
    );
    // <c:smooth val="1"/> — line / area / scatter only.
    const smoothEl = firstChildElement(ser, qname('c', 'smooth', NS_C));
    const smooth = smoothEl !== null && readXmlBool(smoothEl);
    const trendline = readTrendline(ser);
    // Per-series <c:dLbls> overrides the chart-level toggles for this
    // one series.
    // A <c:dLbls> holding only per-point <c:dLbl> overrides carries no
    // series-level defaults, so `dataLabels` stays absent in that case.
    const serDLblsEl = firstChildElement(ser, NAME_D_LBLS);
    const serDataLabels =
      serDLblsEl !== null && hasDataLabelGroupChild(serDLblsEl)
        ? readSeriesLevelDataLabels(serDLblsEl)
        : undefined;
    const pointDataLabels =
      serDLblsEl !== null ? readPointDataLabels(serDLblsEl, pointCount, serDataLabels) : undefined;
    series.push({
      name,
      values: values ?? [],
      ...(groupKind !== kind && isComboSeriesKind(groupKind) ? { chartKind: groupKind } : {}),
      ...(secondary ? { secondaryAxis: true } : {}),
      ...(xValues !== null ? { xValues } : {}),
      ...(bubbleSizes !== null ? { bubbleSizes } : {}),
      ...(color !== undefined ? { color } : {}),
      ...(lineWidthEmu !== undefined ? { lineWidthEmu } : {}),
      ...(lineColor !== undefined ? { lineColor } : {}),
      ...(lineDash !== undefined ? { lineDash } : {}),
      ...(markerSymbol !== undefined ? { markerSymbol } : {}),
      ...(markerSizePt !== undefined ? { markerSizePt } : {}),
      ...(markerColor !== undefined ? { markerColor } : {}),
      ...(markerLineColor !== undefined ? { markerLineColor } : {}),
      ...(invertIfNegative !== undefined ? { invertIfNegative } : {}),
      ...(pointColors !== undefined ? { pointColors } : {}),
      ...(pointExplosions !== undefined ? { pointExplosions } : {}),
      ...(pointDataLabels !== undefined ? { pointDataLabels } : {}),
      ...(smoothEl !== null ? { smooth } : {}),
      ...(trendline !== undefined ? { trendline } : {}),
      ...(serDataLabels !== undefined ? { dataLabels: serDataLabels } : {}),
    });
  }

  const categories = categoriesFromFirst ?? [];
  const title = readTitle(chart);
  const titleStyle = readTitleStyle(chart);

  // <c:dLbls> can sit either on the plotted-kind element (`barChart`,
  // `lineChart`, …) for chart-level defaults or on each `<c:ser>` for
  // per-series overrides. Surface the plotted-kind defaults; per-series
  // toggles are deferred until renderers care.
  const dLbls = firstChildElement(plotted, NAME_D_LBLS);
  const dataLabels = dLbls !== null ? readSeriesLevelDataLabels(dLbls) : undefined;

  // <c:valAx> lives on the plotArea (not on the plotted-kind element). The
  // primary axis is the first one no secondary group references — writers
  // may list the secondary pair first.
  let valAx: XmlElement | null = null;
  let secValAx: XmlElement | null = null;
  for (const el of allChildElements(plotArea, NAME_VAL_AX)) {
    const axIdEl = firstChildElement(el, qname('c', 'axId', NS_C));
    const id = axIdEl !== null ? getAttrValue(axIdEl, ATTR_VAL) : null;
    if (id !== null && secondaryValAxisIds.has(id)) secValAx ??= el;
    else valAx ??= el;
  }
  if (valAx === null) {
    valAx = secValAx;
    secValAx = null;
  }
  const valueAxis = valAx !== null ? readValueAxisScaling(valAx) : undefined;

  // <c:grouping val="clustered|stacked|percentStacked|standard"/>
  // sits as a direct child of the plotted-kind element. Pie and line
  // kinds use different schemas (no grouping); restrict to the kinds
  // that actually carry it.
  let grouping: ChartGrouping | undefined;
  if (kind === 'column' || kind === 'bar' || kind === 'area' || kind === 'line') {
    const groupingEl = firstChildElement(plotted, qname('c', 'grouping', NS_C));
    if (groupingEl) {
      const g = getAttrValue(groupingEl, ATTR_VAL);
      if (g === 'clustered' || g === 'stacked' || g === 'percentStacked' || g === 'standard') {
        grouping = g;
      }
    }
  }

  // <c:dropLines> and <c:hiLowLines> on the plotted-kind element.
  // Both flags are pure booleans for our purposes — the line color /
  // style they author would require the full ln cascade and isn't
  // worth modeling at this layer.
  const dropLinesEl = firstChildElement(plotted, qname('c', 'dropLines', NS_C));
  const hiLowLinesEl = firstChildElement(plotted, qname('c', 'hiLowLines', NS_C));
  const dropLines = dropLinesEl !== null ? true : undefined;
  const hiLowLines = hiLowLinesEl !== null ? true : undefined;

  // <c:lineChart><c:marker val="1"/> toggles point markers for the whole
  // line chart: present (the "Line with Markers" subtype) shows them, absent
  // (the plain "Line" subtype) hides them. An empty <c:marker/> defaults to
  // true per CT_Boolean. Only meaningful for line charts.
  let lineMarkers: boolean | undefined;
  if (kind === 'line') {
    const markerEl = firstChildElement(plotted, qname('c', 'marker', NS_C));
    if (markerEl === null) {
      lineMarkers = false;
    } else {
      const v = getAttrValue(markerEl, ATTR_VAL);
      lineMarkers = v === null ? true : v === '1' || v === 'true';
    }
  }

  // <c:gapWidth> and <c:overlap> live on the plotted-kind element and
  // tune the bar / column spacing. PowerPoint defaults: gapWidth=150
  // (1.5× bar width gap), overlap=0 (clustered) or 100 (stacked).
  let gapWidthPct: number | undefined;
  let overlapPct: number | undefined;
  if (kind === 'column' || kind === 'bar') {
    const gwEl = firstChildElement(plotted, qname('c', 'gapWidth', NS_C));
    if (gwEl) {
      const v = getAttrValue(gwEl, ATTR_VAL);
      if (v !== null) {
        const n = Number.parseInt(v, 10);
        if (Number.isFinite(n)) gapWidthPct = n;
      }
    }
    const ovEl = firstChildElement(plotted, qname('c', 'overlap', NS_C));
    if (ovEl) {
      const v = getAttrValue(ovEl, ATTR_VAL);
      if (v !== null) {
        const n = Number.parseInt(v, 10);
        if (Number.isFinite(n)) overlapPct = n;
      }
    }
  }

  // Axis titles — <c:catAx><c:title> and <c:valAx><c:title>. Both
  // use the same rich-text container as the chart title, so reuse
  // readTitle's projection.
  let categoryAxisTitle: string | undefined;
  let categoryAxisTitleStyle: ChartTextStyle | undefined;
  let categoryAxisTitleRotationDeg: number | undefined;
  let categoryAxisLabelStyle: ChartTextStyle | undefined;
  let categoryAxisLabelRotationDeg: number | undefined;
  let valueAxisLabelRotationDeg: number | undefined;
  let valueAxisMajorTickMark: ChartSpec['valueAxisMajorTickMark'];
  let categoryAxisMajorTickMark: ChartSpec['categoryAxisMajorTickMark'];
  let valueAxisMinorTickMark: ChartSpec['valueAxisMinorTickMark'];
  let categoryAxisMinorTickMark: ChartSpec['categoryAxisMinorTickMark'];
  const readTickMarkLocal = (
    axis: XmlElement,
    local: 'majorTickMark' | 'minorTickMark',
  ): 'in' | 'out' | 'cross' | 'none' | undefined => {
    const el = firstChildElement(axis, qname('c', local, NS_C));
    if (!el) return undefined;
    const v = getAttrValue(el, ATTR_VAL);
    if (v === 'in' || v === 'out' || v === 'cross' || v === 'none') return v;
    return undefined;
  };
  const readTickMark = (axis: XmlElement): 'in' | 'out' | 'cross' | 'none' | undefined =>
    readTickMarkLocal(axis, 'majorTickMark');
  let valueAxisTitle: string | undefined;
  let valueAxisTitleStyle: ChartTextStyle | undefined;
  let valueAxisTitleRotationDeg: number | undefined;
  let valueAxisLabelStyle: ChartTextStyle | undefined;
  let categoryAxisHidden: boolean | undefined;
  let valueAxisHidden: boolean | undefined;
  let categoryAxisTickLabelSkip: number | undefined;
  let categoryAxisTickMarkSkip: number | undefined;
  let categoryAxisTickLabelPos: ChartSpec['categoryAxisTickLabelPos'];
  let categoryAxisLabelOffset: number | undefined;
  let categoryAxisLabelAlign: ChartSpec['categoryAxisLabelAlign'];
  let categoryAxisNumberFormat: string | undefined;
  let categoryAxisNoMultiLevelLabel: boolean | undefined;
  let categoryAxisLineColor: string | undefined;
  let valueAxisLineColor: string | undefined;
  let categoryAxisMajorGridlines: boolean | undefined;
  let categoryAxisMinorGridlines: boolean | undefined;
  let categoryAxisMajorGridlineColor: string | undefined;
  let categoryAxisMinorGridlineColor: string | undefined;
  let valueAxisMinorGridlineColor: string | undefined;
  let categoryAxisLineWidthEmu: number | undefined;
  let valueAxisLineWidthEmu: number | undefined;
  let categoryAxisMajorGridlineWidthEmu: number | undefined;
  let categoryAxisMinorGridlineWidthEmu: number | undefined;
  let valueAxisMajorGridlineWidthEmu: number | undefined;
  let valueAxisMinorGridlineWidthEmu: number | undefined;
  // <c:majorGridlines|minorGridlines><c:spPr><a:ln><a:solidFill><a:srgbClr val=…/>.
  const readGridlineColor = (gl: XmlElement): string | undefined => {
    const spPr = firstChildElement(gl, NAME_SP_PR_C);
    if (!spPr) return undefined;
    const ln = firstChildElement(spPr, qname('a', 'ln', NS_A));
    if (!ln) return undefined;
    const solid = firstChildElement(ln, NAME_SOLID_FILL);
    if (!solid) return undefined;
    const srgb = firstChildElement(solid, NAME_SRGB_CLR);
    if (!srgb) return undefined;
    const v = getAttrValue(srgb, ATTR_VAL);
    return v !== null ? `#${v.toUpperCase()}` : undefined;
  };
  // `w` of the <c:spPr><a:ln> directly under a gridlines or axis element.
  const readLineWidthEmu = (host: XmlElement): number | undefined => {
    const spPr = firstChildElement(host, NAME_SP_PR_C);
    const ln = spPr !== null ? firstChildElement(spPr, qname('a', 'ln', NS_A)) : null;
    const w = ln !== null ? getAttrValue(ln, qname('', 'w', '')) : null;
    if (w === null || !/^\d+$/.test(w)) return undefined;
    // `\d+` already excludes negatives; 0 is a valid ST_LineWidth.
    return Number.parseInt(w, 10);
  };
  // <c:catAx|valAx><c:spPr><a:ln><a:solidFill><a:srgbClr val=…/>.
  const readAxisLineColor = (axis: XmlElement): string | undefined => {
    const spPr = firstChildElement(axis, NAME_SP_PR_C);
    if (!spPr) return undefined;
    const ln = firstChildElement(spPr, qname('a', 'ln', NS_A));
    if (!ln) return undefined;
    const solid = firstChildElement(ln, NAME_SOLID_FILL);
    if (!solid) return undefined;
    const srgb = firstChildElement(solid, NAME_SRGB_CLR);
    if (!srgb) return undefined;
    const v = getAttrValue(srgb, ATTR_VAL);
    return v !== null ? `#${v.toUpperCase()}` : undefined;
  };
  const catAx = findFirst(plotArea, ['catAx', 'dateAx', 'serAx']);
  const isHidden = (axis: XmlElement): boolean | undefined => {
    const d = firstChildElement(axis, qname('c', 'delete', NS_C));
    if (!d) return undefined;
    const v = getAttrValue(d, ATTR_VAL);
    return v === null || v === '1' || v === 'true';
  };
  let categoryAxisOrientation: 'minMax' | 'maxMin' | undefined;
  let valueAxisOrientation: 'minMax' | 'maxMin' | undefined;
  let valueAxisCrosses: ChartSpec['valueAxisCrosses'];
  let valueAxisCrossBetween: ChartSpec['valueAxisCrossBetween'];
  let valueAxisTickLabelPos: ChartSpec['valueAxisTickLabelPos'];
  const readAxisOrientation = (axis: XmlElement): 'minMax' | 'maxMin' | undefined => {
    const scaling = firstChildElement(axis, qname('c', 'scaling', NS_C));
    if (!scaling) return undefined;
    const orientationEl = firstChildElement(scaling, qname('c', 'orientation', NS_C));
    if (!orientationEl) return undefined;
    const v = getAttrValue(orientationEl, ATTR_VAL);
    if (v === 'minMax' || v === 'maxMin') return v;
    return undefined;
  };
  if (catAx) {
    const t = readTitle(catAx);
    if (t !== undefined) categoryAxisTitle = t;
    const catTitleEl = firstChildElement(catAx, NAME_TITLE);
    if (catTitleEl) {
      categoryAxisTitleStyle = readTitleStyleOf(catTitleEl);
      categoryAxisTitleRotationDeg = readTitleRotationDeg(catTitleEl);
    }
    const catTxPr = firstChildElement(catAx, qname('c', 'txPr', NS_C));
    if (catTxPr) {
      categoryAxisLabelStyle = readLabelStyle(catTxPr);
      // <c:txPr><a:bodyPr rot="N"/> — N is in 60000ths of a degree.
      const bodyPr = firstChildElement(catTxPr, qname('a', 'bodyPr', NS_A));
      if (bodyPr) {
        const rotRaw = getAttrValue(bodyPr, qname('', 'rot', ''));
        if (rotRaw !== null) {
          const n = Number.parseInt(rotRaw, 10);
          if (Number.isFinite(n)) categoryAxisLabelRotationDeg = n / 60000;
        }
      }
    }
    categoryAxisHidden = isHidden(catAx);
    categoryAxisLineColor = readAxisLineColor(catAx);
    categoryAxisLineWidthEmu = readLineWidthEmu(catAx);
    const catMajorGl = firstChildElement(catAx, qname('c', 'majorGridlines', NS_C));
    categoryAxisMajorGridlines = catMajorGl !== null;
    if (catMajorGl) {
      categoryAxisMajorGridlineColor = readGridlineColor(catMajorGl);
      categoryAxisMajorGridlineWidthEmu = readLineWidthEmu(catMajorGl);
    }
    const catMinorGl = firstChildElement(catAx, qname('c', 'minorGridlines', NS_C));
    categoryAxisMinorGridlines = catMinorGl !== null;
    if (catMinorGl) {
      categoryAxisMinorGridlineColor = readGridlineColor(catMinorGl);
      categoryAxisMinorGridlineWidthEmu = readLineWidthEmu(catMinorGl);
    }
    categoryAxisOrientation = readAxisOrientation(catAx);
    categoryAxisMajorTickMark = readTickMark(catAx);
    categoryAxisMinorTickMark = readTickMarkLocal(catAx, 'minorTickMark');
    const skipEl = firstChildElement(catAx, qname('c', 'tickLblSkip', NS_C));
    if (skipEl) {
      const v = getAttrValue(skipEl, ATTR_VAL);
      if (v !== null) {
        const n = Number.parseInt(v, 10);
        if (Number.isFinite(n) && n > 1) categoryAxisTickLabelSkip = n;
      }
    }
    const markSkipEl = firstChildElement(catAx, qname('c', 'tickMarkSkip', NS_C));
    if (markSkipEl) {
      const v = getAttrValue(markSkipEl, ATTR_VAL);
      if (v !== null) {
        const n = Number.parseInt(v, 10);
        if (Number.isFinite(n) && n > 1) categoryAxisTickMarkSkip = n;
      }
    }
    const posEl = firstChildElement(catAx, qname('c', 'tickLblPos', NS_C));
    if (posEl) {
      const v = getAttrValue(posEl, ATTR_VAL);
      if (v === 'none' || v === 'low' || v === 'high' || v === 'nextTo') {
        categoryAxisTickLabelPos = v;
      }
    }
    const lblOffsetEl = firstChildElement(catAx, qname('c', 'lblOffset', NS_C));
    if (lblOffsetEl) {
      const v = getAttrValue(lblOffsetEl, ATTR_VAL);
      if (v !== null) {
        const n = Number.parseInt(v, 10);
        if (Number.isFinite(n) && n !== 100) categoryAxisLabelOffset = n;
      }
    }
    const lblAlgnEl = firstChildElement(catAx, qname('c', 'lblAlgn', NS_C));
    if (lblAlgnEl) {
      const v = getAttrValue(lblAlgnEl, ATTR_VAL);
      if (v === 'ctr' || v === 'l' || v === 'r') categoryAxisLabelAlign = v;
    }
    const catNumFmtEl = firstChildElement(catAx, qname('c', 'numFmt', NS_C));
    if (catNumFmtEl) {
      const fc = getAttrValue(catNumFmtEl, qname('', 'formatCode', ''));
      if (fc !== null && fc.length > 0 && fc !== 'General') categoryAxisNumberFormat = fc;
    }
    const noMultiLvlEl = firstChildElement(catAx, qname('c', 'noMultiLvlLbl', NS_C));
    if (noMultiLvlEl) {
      const v = getAttrValue(noMultiLvlEl, ATTR_VAL);
      if (v === '1' || v === 'true') categoryAxisNoMultiLevelLabel = true;
    }
  }
  // <c:majorGridlines> / <c:minorGridlines> presence governs visibility.
  // Surface as explicit booleans so renderers can branch on "absent
  // means hidden" (matches ECMA-376 §21.2.2.122).
  let valueAxisMajorGridlines: boolean | undefined;
  let valueAxisMinorGridlines: boolean | undefined;
  let valueAxisMajorGridlineColor: string | undefined;
  if (valAx) {
    const t = readTitle(valAx);
    if (t !== undefined) valueAxisTitle = t;
    const valTitleEl = firstChildElement(valAx, NAME_TITLE);
    if (valTitleEl) {
      valueAxisTitleStyle = readTitleStyleOf(valTitleEl);
      valueAxisTitleRotationDeg = readTitleRotationDeg(valTitleEl);
    }
    const valTxPr = firstChildElement(valAx, qname('c', 'txPr', NS_C));
    if (valTxPr) {
      valueAxisLabelStyle = readLabelStyle(valTxPr);
      const bodyPr = firstChildElement(valTxPr, qname('a', 'bodyPr', NS_A));
      if (bodyPr) {
        const rotRaw = getAttrValue(bodyPr, qname('', 'rot', ''));
        if (rotRaw !== null) {
          const n = Number.parseInt(rotRaw, 10);
          if (Number.isFinite(n)) valueAxisLabelRotationDeg = n / 60000;
        }
      }
    }
    valueAxisHidden = isHidden(valAx);
    valueAxisLineColor = readAxisLineColor(valAx);
    valueAxisLineWidthEmu = readLineWidthEmu(valAx);
    valueAxisOrientation = readAxisOrientation(valAx);
    valueAxisMajorTickMark = readTickMark(valAx);
    valueAxisMinorTickMark = readTickMarkLocal(valAx, 'minorTickMark');
    const valPosEl = firstChildElement(valAx, qname('c', 'tickLblPos', NS_C));
    if (valPosEl) {
      const v = getAttrValue(valPosEl, ATTR_VAL);
      if (v === 'none' || v === 'low' || v === 'high' || v === 'nextTo') {
        valueAxisTickLabelPos = v;
      }
    }
    // <c:crosses val> and <c:crossesAt val> are mutually exclusive per
    // the schema; crossesAt wins when both are emitted (matches the
    // PowerPoint priority).
    const crossesAtEl = firstChildElement(valAx, qname('c', 'crossesAt', NS_C));
    if (crossesAtEl) {
      const v = getAttrValue(crossesAtEl, ATTR_VAL);
      if (v !== null) {
        const n = Number.parseFloat(v);
        if (Number.isFinite(n)) valueAxisCrosses = { at: n };
      }
    } else {
      const crossesEl = firstChildElement(valAx, qname('c', 'crosses', NS_C));
      if (crossesEl) {
        const v = getAttrValue(crossesEl, ATTR_VAL);
        if (v === 'autoZero' || v === 'min' || v === 'max') valueAxisCrosses = v;
      }
    }
    const crossBetweenEl = firstChildElement(valAx, qname('c', 'crossBetween', NS_C));
    if (crossBetweenEl) {
      const v = getAttrValue(crossBetweenEl, ATTR_VAL);
      if (v === 'between' || v === 'midCat') valueAxisCrossBetween = v;
    }
    const majorGl = firstChildElement(valAx, qname('c', 'majorGridlines', NS_C));
    valueAxisMajorGridlines = majorGl !== null;
    if (majorGl) {
      valueAxisMajorGridlineColor = readGridlineColor(majorGl);
      valueAxisMajorGridlineWidthEmu = readLineWidthEmu(majorGl);
    }
    const minorGl = firstChildElement(valAx, qname('c', 'minorGridlines', NS_C));
    valueAxisMinorGridlines = minorGl !== null;
    if (minorGl) {
      valueAxisMinorGridlineColor = readGridlineColor(minorGl);
      valueAxisMinorGridlineWidthEmu = readLineWidthEmu(minorGl);
    }
  }

  // The secondary value axis is the one only the later plot groups
  // reference; read it with the same helpers as the primary.
  let secondaryValueAxis: ChartSecondaryValueAxis | undefined;
  if (secValAx !== null) {
    const scaling = readValueAxisScaling(secValAx);
    const title = readTitle(secValAx);
    const titleEl = firstChildElement(secValAx, NAME_TITLE);
    const titleStyle = titleEl !== null ? readTitleStyleOf(titleEl) : undefined;
    const txPr = firstChildElement(secValAx, qname('c', 'txPr', NS_C));
    const labelStyle = txPr !== null ? readLabelStyle(txPr) : undefined;
    const majorGl = firstChildElement(secValAx, qname('c', 'majorGridlines', NS_C));
    const majorGridlineColor = majorGl !== null ? readGridlineColor(majorGl) : undefined;
    const majorGridlineWidthEmu = majorGl !== null ? readLineWidthEmu(majorGl) : undefined;
    const lineColor = readAxisLineColor(secValAx);
    const lineWidthEmu = readLineWidthEmu(secValAx);
    const majorTickMark = readTickMark(secValAx);
    const minorTickMark = readTickMarkLocal(secValAx, 'minorTickMark');
    const secPosEl = firstChildElement(secValAx, qname('c', 'tickLblPos', NS_C));
    const secPos = secPosEl !== null ? getAttrValue(secPosEl, ATTR_VAL) : null;
    const tickLabelPos =
      secPos === 'none' || secPos === 'low' || secPos === 'high' || secPos === 'nextTo'
        ? secPos
        : undefined;
    const crossBetweenEl = firstChildElement(secValAx, qname('c', 'crossBetween', NS_C));
    const cb = crossBetweenEl !== null ? getAttrValue(crossBetweenEl, ATTR_VAL) : null;
    const crossBetween = cb === 'between' || cb === 'midCat' ? cb : undefined;
    const fields: ChartSecondaryValueAxis = {
      ...(scaling !== undefined ? { scaling } : {}),
      ...(title !== undefined ? { title } : {}),
      ...(titleStyle !== undefined ? { titleStyle } : {}),
      ...(labelStyle !== undefined ? { labelStyle } : {}),
      ...(majorGl !== null ? { majorGridlines: true } : {}),
      ...(majorGridlineColor !== undefined ? { majorGridlineColor } : {}),
      ...(majorGridlineWidthEmu !== undefined ? { majorGridlineWidthEmu } : {}),
      ...(lineColor !== undefined ? { lineColor } : {}),
      ...(lineWidthEmu !== undefined ? { lineWidthEmu } : {}),
      ...(majorTickMark !== undefined ? { majorTickMark } : {}),
      ...(minorTickMark !== undefined ? { minorTickMark } : {}),
      ...(tickLabelPos !== undefined ? { tickLabelPos } : {}),
      ...(crossBetween !== undefined ? { crossBetween } : {}),
    };
    if (Object.keys(fields).length > 0) secondaryValueAxis = fields;
  }

  // Plot area + chart area fills (`<c:spPr><a:solidFill><a:srgbClr/>`).
  const readSpPrFill = (parent: XmlElement): string | undefined => {
    const spPr = firstChildElement(parent, NAME_SP_PR_C);
    if (!spPr) return undefined;
    const solid = firstChildElement(spPr, NAME_SOLID_FILL);
    if (!solid) return undefined;
    const srgb = firstChildElement(solid, NAME_SRGB_CLR);
    if (!srgb) return undefined;
    const v = getAttrValue(srgb, ATTR_VAL);
    return v !== null ? `#${v.toUpperCase()}` : undefined;
  };
  // Stroke color from `<c:spPr><a:ln><a:solidFill><a:srgbClr/>`.
  const readSpPrStrokeColor = (parent: XmlElement): string | undefined => {
    const spPr = firstChildElement(parent, NAME_SP_PR_C);
    if (!spPr) return undefined;
    const ln = firstChildElement(spPr, qname('a', 'ln', NS_A));
    if (!ln) return undefined;
    const solid = firstChildElement(ln, NAME_SOLID_FILL);
    if (!solid) return undefined;
    const srgb = firstChildElement(solid, NAME_SRGB_CLR);
    if (!srgb) return undefined;
    const v = getAttrValue(srgb, ATTR_VAL);
    return v !== null ? `#${v.toUpperCase()}` : undefined;
  };
  const plotAreaFill = readSpPrFill(plotArea);
  const plotAreaStrokeColor = readSpPrStrokeColor(plotArea);
  const chartAreaFill = readSpPrFill(root);
  const chartAreaStrokeColor = readSpPrStrokeColor(root);

  // <c:dispBlanksAs val="…"/> sits on the chart element. Controls how
  // null gaps in line / area series render: 'gap' (default), 'zero', or
  // 'span'.
  let dispBlanksAs: ChartSpec['dispBlanksAs'];
  const dbaEl = firstChildElement(chart, qname('c', 'dispBlanksAs', NS_C));
  if (dbaEl) {
    const v = getAttrValue(dbaEl, ATTR_VAL);
    if (v === 'gap' || v === 'zero' || v === 'span') dispBlanksAs = v;
  }
  // <c:plotVisOnly val="0|1"/> — PowerPoint default is 1. Surface as
  // `plotVisibleCellsOnly: false` only when explicitly 0 so the round-trip
  // doesn't add a redundant `false` for every default-shaped chart.
  let plotVisibleCellsOnly: boolean | undefined;
  const pvoEl = firstChildElement(chart, qname('c', 'plotVisOnly', NS_C));
  if (pvoEl) {
    const v = getAttrValue(pvoEl, ATTR_VAL);
    if (v === '0' || v === 'false') plotVisibleCellsOnly = false;
  }
  // <c:chartSpace><c:roundedCorners val="1"/> — surface only when
  // explicitly true so the round-trip doesn't carry a redundant false.
  let roundedCorners: boolean | undefined;
  const rcEl = firstChildElement(root, qname('c', 'roundedCorners', NS_C));
  if (rcEl) {
    const v = getAttrValue(rcEl, ATTR_VAL);
    if (v === '1' || v === 'true') roundedCorners = true;
  }
  // <c:chartSpace><c:style val="N"/> — PowerPoint chart-style preset
  // (1..48). Surface for round-trip parity; renderers don't act on it.
  let chartStyle: number | undefined;
  const styleEl = firstChildElement(root, qname('c', 'style', NS_C));
  if (styleEl) {
    const v = getAttrValue(styleEl, ATTR_VAL);
    if (v !== null) {
      const n = Number.parseInt(v, 10);
      if (Number.isFinite(n) && n >= 1 && n <= 48) chartStyle = n;
    }
  }
  // <c:chartSpace><c:lang val="…"/> + <c:date1904 val="…"/> — Office
  // metadata. Surface for round-trip parity.
  let language: string | undefined;
  const langEl = firstChildElement(root, qname('c', 'lang', NS_C));
  if (langEl) {
    const v = getAttrValue(langEl, ATTR_VAL);
    if (v !== null && v.length > 0) language = v;
  }
  let date1904: boolean | undefined;
  const dateEl = firstChildElement(root, qname('c', 'date1904', NS_C));
  if (dateEl) {
    const v = getAttrValue(dateEl, ATTR_VAL);
    if (v === '1' || v === 'true') date1904 = true;
  }

  // <c:legend> sits on the chart element (not the plotArea). Read the
  // position; PowerPoint defaults to 'r' (right) when the element is
  // present but has no legendPos. Absent legend element means renderers
  // fall back to whatever they show by default.
  let legend: ChartSpec['legend'];
  const legendEl = firstChildElement(chart, qname('c', 'legend', NS_C));
  if (legendEl) {
    const posEl = firstChildElement(legendEl, qname('c', 'legendPos', NS_C));
    const tok = posEl ? getAttrValue(posEl, ATTR_VAL) : null;
    const ovEl = firstChildElement(legendEl, qname('c', 'overlay', NS_C));
    const overlay = ovEl ? getAttrValue(ovEl, ATTR_VAL) !== '0' : false;
    const txPr = firstChildElement(legendEl, qname('c', 'txPr', NS_C));
    const textStyle = txPr ? readLabelStyle(txPr) : undefined;
    // <c:legendEntry><c:idx val="N"/><c:delete val="1"/> — series indices
    // the author wants suppressed from the legend (trendlines often hide).
    const hiddenIndices: number[] = [];
    for (const c of legendEl.children) {
      if (
        c.kind !== 'element' ||
        c.name.namespaceURI !== NS_C ||
        c.name.localName !== 'legendEntry'
      )
        continue;
      const idxEl = firstChildElement(c, qname('c', 'idx', NS_C));
      const delEl = firstChildElement(c, qname('c', 'delete', NS_C));
      if (!idxEl || !delEl) continue;
      const delV = getAttrValue(delEl, ATTR_VAL);
      const isDeleted = delV === null || delV === '1' || delV === 'true';
      if (!isDeleted) continue;
      const idxV = getAttrValue(idxEl, ATTR_VAL);
      if (idxV === null) continue;
      const n = Number.parseInt(idxV, 10);
      if (Number.isFinite(n) && n >= 0) hiddenIndices.push(n);
    }
    const position: 'r' | 't' | 'b' | 'l' | 'tr' =
      tok === 'r' || tok === 't' || tok === 'b' || tok === 'l' || tok === 'tr' ? tok : 'r';
    legend = {
      position,
      ...(overlay ? { overlay } : {}),
      ...(textStyle !== undefined ? { textStyle } : {}),
      ...(hiddenIndices.length > 0 ? { hiddenIndices } : {}),
    };
  }

  // <c:barChart><c:varyColors val="1"/> etc. — single-series charts use
  // it to color each data point uniquely.
  let varyColors: boolean | undefined;
  const vcEl = firstChildElement(plotted, qname('c', 'varyColors', NS_C));
  if (vcEl) {
    const v = getAttrValue(vcEl, ATTR_VAL);
    varyColors = v === null || v === '1' || v === 'true';
  }

  // <c:title><c:overlay val="1"/>
  let titleOverlay: boolean | undefined;
  const titleEl = firstChildElement(chart, NAME_TITLE);
  if (titleEl) {
    const ovEl = firstChildElement(titleEl, qname('c', 'overlay', NS_C));
    if (ovEl) {
      titleOverlay = getAttrValue(ovEl, ATTR_VAL) !== '0';
    }
  }

  // Pie-specific extras: starting angle + doughnut hole size.
  let firstSliceAngleDeg: number | undefined;
  let holeSizePct: number | undefined;
  if (kind === 'pie' || kind === 'doughnut') {
    const fsAng = firstChildElement(plotted, qname('c', 'firstSliceAng', NS_C));
    if (fsAng) {
      const v = getAttrValue(fsAng, ATTR_VAL);
      if (v !== null) {
        const n = Number.parseInt(v, 10);
        if (Number.isFinite(n)) firstSliceAngleDeg = ((n % 360) + 360) % 360;
      }
    }
    if (kind === 'doughnut') {
      const hs = firstChildElement(plotted, qname('c', 'holeSize', NS_C));
      if (hs) {
        const v = getAttrValue(hs, ATTR_VAL);
        if (v !== null) {
          const n = Number.parseInt(v, 10);
          if (Number.isFinite(n)) holeSizePct = Math.max(10, Math.min(90, n));
        }
      }
    }
  }

  // Scatter / radar / bubble sub-type + bubble sizing, all read from the
  // plotted-kind element.
  let scatterStyle: ChartSpec['scatterStyle'];
  let radarStyle: ChartSpec['radarStyle'];
  let bubbleScale: number | undefined;
  let bubbleSizeRepresents: ChartSpec['bubbleSizeRepresents'];
  if (kind === 'scatter') {
    const el = firstChildElement(plotted, qname('c', 'scatterStyle', NS_C));
    if (el) {
      const v = getAttrValue(el, ATTR_VAL);
      if (
        v === 'none' ||
        v === 'line' ||
        v === 'lineMarker' ||
        v === 'marker' ||
        v === 'smooth' ||
        v === 'smoothMarker'
      ) {
        scatterStyle = v;
      }
    }
  } else if (kind === 'radar') {
    const el = firstChildElement(plotted, qname('c', 'radarStyle', NS_C));
    if (el) {
      const v = getAttrValue(el, ATTR_VAL);
      if (v === 'standard' || v === 'marker' || v === 'filled') radarStyle = v;
    }
  } else if (kind === 'bubble') {
    const bsEl = firstChildElement(plotted, qname('c', 'bubbleScale', NS_C));
    if (bsEl) {
      const v = getAttrValue(bsEl, ATTR_VAL);
      if (v !== null) {
        const n = Number.parseInt(v, 10);
        // ST_BubbleScale spans 0..300; clamp out-of-range author values.
        if (Number.isFinite(n) && n >= 0) bubbleScale = Math.min(300, n);
      }
    }
    const srEl = firstChildElement(plotted, qname('c', 'sizeRepresents', NS_C));
    if (srEl) {
      const v = getAttrValue(srEl, ATTR_VAL);
      // ST_SizeRepresents tokens are 'area' and 'w'; surface 'w' as 'width'.
      if (v === 'area') bubbleSizeRepresents = 'area';
      else if (v === 'w') bubbleSizeRepresents = 'width';
    }
  }

  return {
    kind,
    categories,
    series,
    ...(title !== undefined ? { title } : {}),
    ...(titleStyle !== undefined ? { titleStyle } : {}),
    ...(dataLabels !== undefined ? { dataLabels } : {}),
    ...(valueAxis !== undefined ? { valueAxis } : {}),
    ...(secondaryValueAxis !== undefined ? { secondaryValueAxis } : {}),
    ...(grouping !== undefined ? { grouping } : {}),
    ...(dropLines !== undefined ? { dropLines } : {}),
    ...(hiLowLines !== undefined ? { hiLowLines } : {}),
    ...(lineMarkers !== undefined ? { lineMarkers } : {}),
    ...(gapWidthPct !== undefined ? { gapWidthPct } : {}),
    ...(overlapPct !== undefined ? { overlapPct } : {}),
    ...(legend !== undefined ? { legend } : {}),
    ...(titleOverlay !== undefined ? { titleOverlay } : {}),
    ...(varyColors !== undefined ? { varyColors } : {}),
    ...(dispBlanksAs !== undefined ? { dispBlanksAs } : {}),
    ...(plotVisibleCellsOnly === false ? { plotVisibleCellsOnly: false } : {}),
    ...(roundedCorners === true ? { roundedCorners: true } : {}),
    ...(chartStyle !== undefined ? { chartStyle } : {}),
    ...(language !== undefined ? { language } : {}),
    ...(date1904 === true ? { date1904: true } : {}),
    ...(plotAreaFill !== undefined ? { plotAreaFill } : {}),
    ...(plotAreaStrokeColor !== undefined ? { plotAreaStrokeColor } : {}),
    ...(chartAreaFill !== undefined ? { chartAreaFill } : {}),
    ...(chartAreaStrokeColor !== undefined ? { chartAreaStrokeColor } : {}),
    ...(categoryAxisTitle !== undefined ? { categoryAxisTitle } : {}),
    ...(categoryAxisTitleStyle !== undefined ? { categoryAxisTitleStyle } : {}),
    ...(categoryAxisTitleRotationDeg !== undefined ? { categoryAxisTitleRotationDeg } : {}),
    ...(categoryAxisLabelStyle !== undefined ? { categoryAxisLabelStyle } : {}),
    ...(categoryAxisLabelRotationDeg !== undefined ? { categoryAxisLabelRotationDeg } : {}),
    ...(valueAxisTitle !== undefined ? { valueAxisTitle } : {}),
    ...(valueAxisTitleStyle !== undefined ? { valueAxisTitleStyle } : {}),
    ...(valueAxisTitleRotationDeg !== undefined ? { valueAxisTitleRotationDeg } : {}),
    ...(valueAxisLabelStyle !== undefined ? { valueAxisLabelStyle } : {}),
    ...(categoryAxisHidden !== undefined ? { categoryAxisHidden } : {}),
    ...(valueAxisHidden !== undefined ? { valueAxisHidden } : {}),
    ...(valueAxisMajorGridlines !== undefined ? { valueAxisMajorGridlines } : {}),
    ...(valueAxisMajorGridlineColor !== undefined ? { valueAxisMajorGridlineColor } : {}),
    ...(valueAxisMajorGridlineWidthEmu !== undefined ? { valueAxisMajorGridlineWidthEmu } : {}),
    ...(valueAxisMajorTickMark !== undefined ? { valueAxisMajorTickMark } : {}),
    ...(valueAxisMinorTickMark !== undefined ? { valueAxisMinorTickMark } : {}),
    ...(valueAxisTickLabelPos !== undefined ? { valueAxisTickLabelPos } : {}),
    ...(valueAxisLabelRotationDeg !== undefined ? { valueAxisLabelRotationDeg } : {}),
    ...(categoryAxisMajorTickMark !== undefined ? { categoryAxisMajorTickMark } : {}),
    ...(categoryAxisMinorTickMark !== undefined ? { categoryAxisMinorTickMark } : {}),
    ...(valueAxisMinorGridlines !== undefined ? { valueAxisMinorGridlines } : {}),
    ...(categoryAxisTickLabelSkip !== undefined ? { categoryAxisTickLabelSkip } : {}),
    ...(categoryAxisTickMarkSkip !== undefined ? { categoryAxisTickMarkSkip } : {}),
    ...(categoryAxisTickLabelPos !== undefined ? { categoryAxisTickLabelPos } : {}),
    ...(categoryAxisLabelOffset !== undefined ? { categoryAxisLabelOffset } : {}),
    ...(categoryAxisLabelAlign !== undefined ? { categoryAxisLabelAlign } : {}),
    ...(categoryAxisNumberFormat !== undefined ? { categoryAxisNumberFormat } : {}),
    ...(categoryAxisNoMultiLevelLabel === true ? { categoryAxisNoMultiLevelLabel: true } : {}),
    ...(categoryAxisLineColor !== undefined ? { categoryAxisLineColor } : {}),
    ...(valueAxisLineColor !== undefined ? { valueAxisLineColor } : {}),
    ...(categoryAxisMajorGridlines === true ? { categoryAxisMajorGridlines: true } : {}),
    ...(categoryAxisMinorGridlines === true ? { categoryAxisMinorGridlines: true } : {}),
    ...(categoryAxisMajorGridlineColor !== undefined ? { categoryAxisMajorGridlineColor } : {}),
    ...(categoryAxisMinorGridlineColor !== undefined ? { categoryAxisMinorGridlineColor } : {}),
    ...(valueAxisMinorGridlineColor !== undefined ? { valueAxisMinorGridlineColor } : {}),
    ...(categoryAxisLineWidthEmu !== undefined ? { categoryAxisLineWidthEmu } : {}),
    ...(valueAxisLineWidthEmu !== undefined ? { valueAxisLineWidthEmu } : {}),
    ...(categoryAxisMajorGridlineWidthEmu !== undefined
      ? { categoryAxisMajorGridlineWidthEmu }
      : {}),
    ...(categoryAxisMinorGridlineWidthEmu !== undefined
      ? { categoryAxisMinorGridlineWidthEmu }
      : {}),
    ...(valueAxisMinorGridlineWidthEmu !== undefined ? { valueAxisMinorGridlineWidthEmu } : {}),
    ...(categoryAxisOrientation !== undefined ? { categoryAxisOrientation } : {}),
    ...(valueAxisOrientation !== undefined ? { valueAxisOrientation } : {}),
    ...(valueAxisCrosses !== undefined ? { valueAxisCrosses } : {}),
    ...(valueAxisCrossBetween !== undefined ? { valueAxisCrossBetween } : {}),
    ...(firstSliceAngleDeg !== undefined ? { firstSliceAngleDeg } : {}),
    ...(holeSizePct !== undefined ? { holeSizePct } : {}),
    ...(scatterStyle !== undefined ? { scatterStyle } : {}),
    ...(radarStyle !== undefined ? { radarStyle } : {}),
    ...(bubbleScale !== undefined ? { bubbleScale } : {}),
    ...(bubbleSizeRepresents !== undefined ? { bubbleSizeRepresents } : {}),
  };
};

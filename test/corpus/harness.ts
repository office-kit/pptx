// Builds both libraries' output for a corpus case and extracts the canonical
// slide drawing tree from each, so callers can diff or score them.

import { createRequire } from 'node:module';
import {
  addSlide,
  createPresentation,
  findSlideLayout,
  getShapeChartCategories,
  getShapeChartKind,
  getShapeChartSeriesNames,
  getShapeChartSeriesValues,
  getSlides,
  getSlideShapes,
  isChartShape,
  loadPresentation,
  savePresentation,
} from '../../src/api/index.ts';
import { partName } from '../../src/internal/opc/index.ts';
import { OpcPackage } from '../../src/internal/parts/index.ts';
import { canonicalSpTree } from './canonical.ts';
import type { CorpusCase } from './cases.ts';
import type { PgjsCtor } from './pptxgenjs-types.ts';

const require = createRequire(import.meta.url);

// PptxGenJS ships a CJS bundle and pulls in jszip; both live in the submodule's
// own node_modules (run `git submodule update --init references/PptxGenJS &&
// npm --prefix references/PptxGenJS install jszip` once). When the submodule
// isn't checked out — e.g. a clean CI job that skips submodules — the corpus
// suite skips itself rather than failing the build.
let Pptx: PgjsCtor | null = null;
let loadError = '';
try {
  Pptx = require('../../references/PptxGenJS/dist/pptxgen.cjs.js') as PgjsCtor;
} catch (e) {
  loadError = (e as Error).message;
}

export const pptxGenJsAvailable = (): boolean => Pptx !== null;
export const pptxGenJsLoadError = (): string => loadError;

const dec = (b: Uint8Array): string => new TextDecoder().decode(b);

const slideXmlOf = (bytes: Uint8Array): string => {
  const pkg = OpcPackage.load(bytes);
  const part = pkg.getPart(partName('/ppt/slides/slide1.xml'));
  if (!part) throw new Error('no slide1.xml');
  return dec(part.data);
};

// The first chart part, if the deck has one. The part name is NOT hard-coded:
// PptxGenJS keeps a process-global chart counter, so the second chart built in
// the same test run lands at `chart2.xml`, the third at `chart3.xml`, etc.
const CHART_PART = /^\/ppt\/charts\/chart\d+\.xml$/;
const chartXmlOf = (bytes: Uint8Array): string | null => {
  const pkg = OpcPackage.load(bytes);
  const part = pkg.parts.find((p) => CHART_PART.test(p.name));
  return part ? dec(part.data) : null;
};

/**
 * The *semantic* content of a deck's first chart — type plus each series'
 * name and values — read back with @office-kit/pptx's own reader. Charts can't be
 * compared as raw XML: PptxGenJS stamps dozens of opinionated chrome defaults
 * (data-label blocks, gridlines, its own palette, `multiLvlStrRef` categories)
 * that PowerPoint treats as optional and @office-kit/pptx leaves to inheritance. What
 * must match is the data and the chart type, which is what this extracts —
 * from *either* library's output, since @office-kit/pptx can read any PPTX.
 */
export interface ChartSemantics {
  kind: string | null;
  categories: ReadonlyArray<string> | null;
  series: Array<{ name: string | null; values: ReadonlyArray<number | null> | null }>;
}

export const chartSemanticsOf = async (bytes: Uint8Array): Promise<ChartSemantics | null> => {
  const pres = await loadPresentation(bytes);
  for (const slide of getSlides(pres)) {
    for (const shape of getSlideShapes(slide)) {
      if (!isChartShape(shape)) continue;
      const names = getShapeChartSeriesNames(shape) ?? [];
      return {
        kind: getShapeChartKind(shape),
        categories: getShapeChartCategories(shape),
        series: names.map((name) => ({ name, values: getShapeChartSeriesValues(shape, name) })),
      };
    }
  }
  return null;
};

export interface CaseResult {
  id: string;
  kitXml: string;
  pgjsXml: string;
  kitCanonical: string;
  pgjsCanonical: string;
  /** Present only for chart cases: the chart part XML and read-back content. */
  kitChartXml: string | null;
  kitChartSemantics: ChartSemantics | null;
  pgjsChartSemantics: ChartSemantics | null;
}

export const runCase = async (c: CorpusCase): Promise<CaseResult> => {
  if (!Pptx) throw new Error(`PptxGenJS unavailable: ${loadError}`);
  // PptxGenJS side.
  const pptx = new Pptx();
  const pgjsSlide = pptx.addSlide();
  c.pgjs(pgjsSlide, pptx);
  const pgjsBytes = new Uint8Array(await pptx.write('nodebuffer'));

  // @office-kit/pptx side — blank layout so no placeholder scaffolding pollutes the diff.
  const pres = createPresentation({ size: '16:9' });
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('no Blank layout in createPresentation deck');
  const slide = addSlide(pres, { layout });
  c.kit(pres, slide);
  const kitBytes = await savePresentation(pres);

  const kitXml = slideXmlOf(kitBytes);
  const pgjsXml = slideXmlOf(pgjsBytes);
  const kitChartXml = chartXmlOf(kitBytes);

  return {
    id: c.id,
    kitXml,
    pgjsXml,
    kitCanonical: canonicalSpTree(kitXml),
    pgjsCanonical: canonicalSpTree(pgjsXml),
    kitChartXml,
    // Resolve the chart via relationships (name-independent), so a PptxGenJS
    // chart at chart2.xml/chart3.xml is still read.
    kitChartSemantics: kitChartXml ? await chartSemanticsOf(kitBytes) : null,
    pgjsChartSemantics: chartXmlOf(pgjsBytes) ? await chartSemanticsOf(pgjsBytes) : null,
  };
};

// Builds both libraries' output for a corpus case and extracts the canonical
// slide drawing tree from each, so callers can diff or score them.

import { createRequire } from 'node:module';
import {
  addSlide,
  createPresentation,
  findSlideLayout,
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

export interface CaseResult {
  id: string;
  kitXml: string;
  pgjsXml: string;
  kitCanonical: string;
  pgjsCanonical: string;
}

export const runCase = async (c: CorpusCase): Promise<CaseResult> => {
  if (!Pptx) throw new Error(`PptxGenJS unavailable: ${loadError}`);
  // PptxGenJS side.
  const pptx = new Pptx();
  const pgjsSlide = pptx.addSlide();
  c.pgjs(pgjsSlide, pptx);
  const pgjsBytes = new Uint8Array(await pptx.write('nodebuffer'));

  // pptx-kit side — blank layout so no placeholder scaffolding pollutes the diff.
  const pres = createPresentation({ size: '16:9' });
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('no Blank layout in createPresentation deck');
  const slide = addSlide(pres, { layout });
  c.kit(pres, slide);
  const kitBytes = await savePresentation(pres);

  const kitXml = slideXmlOf(kitBytes);
  const pgjsXml = slideXmlOf(pgjsBytes);
  return {
    id: c.id,
    kitXml,
    pgjsXml,
    kitCanonical: canonicalSpTree(kitXml),
    pgjsCanonical: canonicalSpTree(pgjsXml),
  };
};

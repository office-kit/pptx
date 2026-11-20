// Decks whose slide 1 holds hand-written OMML equations, for tests that read
// or render `<a14:m>` before the library can author it.

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  addBlankSlide,
  addSlideTable,
  addSlideTextBox,
  createPresentation,
  inches,
  loadPresentation,
  savePresentation,
  type PresentationData,
} from '../../src/api/index.ts';

const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const A14 = 'http://schemas.microsoft.com/office/drawing/2010/main';
const MARKER = 'MATH_MARKER';
const MARKER_RUN = /<a:r>(?:(?!<\/a:r>).)*?MATH_MARKER<\/a:t><\/a:r>/;

/** `<m:r>` holding `text`. */
export const mathRun = (text: string): string => `<m:r><m:t>${text}</m:t></m:r>`;

/** An inline `<a14:m>` around `inner`. */
export const mathZone = (inner: string): string =>
  `<a14:m xmlns:a14="${A14}"><m:oMath xmlns:m="${M}">${inner}</m:oMath></a14:m>`;

/** A display `<a14:m>` holding one `<m:oMath>` per entry. */
export const displayMathZone = (...equations: string[]): string =>
  `<a14:m xmlns:a14="${A14}"><m:oMathPara xmlns:m="${M}"><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr>${equations.map((e) => `<m:oMath>${e}</m:oMath>`).join('')}</m:oMathPara></a14:m>`;

/**
 * A one-slide deck whose text box (or 1×1 table) paragraph content is
 * `paragraphXml`, in place of the single run the box was created with.
 */
export const presentationWithParagraph = async (
  paragraphXml: string,
  { table = false } = {},
): Promise<PresentationData> => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const box = { x: inches(1), y: inches(1), w: inches(6), h: inches(1.5) };
  if (table) addSlideTable(slide, { ...box, rows: [[MARKER]] });
  else addSlideTextBox(slide, { ...box, text: MARKER });
  const zip = unzipSync(await savePresentation(pres));
  const part = 'ppt/slides/slide1.xml';
  const source = strFromU8(zip[part]!);
  const patched = source.replace(MARKER_RUN, paragraphXml);
  if (patched === source) throw new Error('marker run not found');
  zip[part] = strToU8(patched);
  return loadPresentation(zipSync(zip));
};

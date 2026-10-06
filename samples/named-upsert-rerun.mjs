// A save script that is safe to run repeatedly, written only against the public
// API. Adding shapes is not idempotent, so the script finds its own shape by
// name and updates it; a second run must not add a copy.
//
//   pnpm build
//   node samples/named-upsert-rerun.mjs samples/out/upsert.pptx Q1
//   node samples/named-upsert-rerun.mjs samples/out/upsert.pptx Q2
//
// Each run prints the reopened deck's `kpi` count, text, shape IDs and part
// count; only the text should change between runs.
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname } from 'node:path';
import {
  addSlideTextBox,
  addTitleSlide,
  createPresentation,
  emu,
  findShapeByName,
  getShapeId,
  getShapeName,
  getShapeText,
  getSlideShapes,
  getSlides,
  listPackageParts,
  loadPresentation,
  savePresentation,
  setShapeText,
} from '@office-kit/pptx';

const [file, kpi] = process.argv.slice(2);
if (!file || !kpi) {
  console.error('Usage: node samples/named-upsert-rerun.mjs <deck.pptx> <kpi text>');
  process.exit(2);
}

const readDeck = async () => {
  try {
    return await loadPresentation(await readFile(file));
  } catch (cause) {
    if (cause?.code !== 'ENOENT') throw cause;
    const pres = createPresentation();
    addTitleSlide(pres, 'Quarterly review');
    return pres;
  }
};

const pres = await readDeck();
const slide = getSlides(pres)[0];
const box = findShapeByName(slide, 'kpi');
if (box) setShapeText(box, kpi);
else
  addSlideTextBox(slide, {
    x: emu(457200),
    y: emu(457200),
    w: emu(3657600),
    h: emu(457200),
    text: kpi,
    name: 'kpi',
  });
await mkdir(dirname(file), { recursive: true });
await writeFile(file, await savePresentation(pres));

const reopened = await loadPresentation(await readFile(file));
const shapes = getSlideShapes(getSlides(reopened)[0]);
console.log(
  JSON.stringify({
    kpiCount: shapes.filter((shape) => getShapeName(shape) === 'kpi').length,
    kpiText: getShapeText(shapes.find((shape) => getShapeName(shape) === 'kpi')),
    ids: shapes.map(getShapeId),
    parts: listPackageParts(reopened).length,
  }),
);

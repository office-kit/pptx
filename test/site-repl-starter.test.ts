// The REPL's starter deck is plain script text with every library function as
// a free identifier, so the type checker never sees it. Run it the way the
// REPL page does, so an API rename fails here instead of on the live site.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import * as kit from '../src/api/index.ts';
import { partName } from '../src/internal/opc/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const starterPath = fileURLToPath(new URL('../site/repl/default-deck.js', import.meta.url));
const SLIDE_COUNT = 6;
const CHART_COUNT = 2;

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;
const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

const buildStarter = async (): Promise<kit.PresentationData> => {
  const code = await readFile(starterPath, 'utf8');
  const entries = Object.entries(kit);
  const run = new Function(
    ...entries.map(([name]) => name),
    'pres',
    `'use strict';\nreturn (async () => {\n${code}\n})();`,
  );
  const pres = kit.createPresentation();
  await run(...entries.map(([, value]) => value), pres);
  return pres;
};

describe('REPL starter deck', () => {
  it('builds, renders, and survives a save and reload', async () => {
    const pres = await buildStarter();
    const slides = kit.getSlides(pres);
    expect(slides).toHaveLength(SLIDE_COUNT);
    for (const slide of slides) {
      expect(renderSlideToSvg(pres, slide)).toContain('<svg');
    }
    const reloaded = await kit.loadPresentation(await kit.savePresentation(pres));
    expect(kit.getSlides(reloaded)).toHaveLength(SLIDE_COUNT);
  });

  skipIfNoXmllint('writes schema-valid slides and charts', async () => {
    const pres = await buildStarter();
    const pkg = kit._internalPackageOf(
      await kit.loadPresentation(await kit.savePresentation(pres)),
    );
    for (let n = 1; n <= SLIDE_COUNT; n++) {
      expectSchemaValid(decode(pkg.getPart(partName(`/ppt/slides/slide${n}.xml`))!.data), 'pml');
    }
    for (let n = 1; n <= CHART_COUNT; n++) {
      expectSchemaValid(decode(pkg.getPart(partName(`/ppt/charts/chart${n}.xml`))!.data), 'chart');
    }
  });
});

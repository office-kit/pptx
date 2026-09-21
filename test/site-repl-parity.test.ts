// The REPL's two tabs exist to compare two ways of writing one deck, which
// only works while they really are one deck. Build both starters and hold
// them to the same parts: every slide and every chart, byte for byte.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import * as kit from '../src/api/index.ts';
import * as dsl from '../packages/dsl/src/index.ts';
import * as jsxRuntime from '../packages/dsl/src/jsx-runtime.ts';
import { asyncBody, evaluateTsx } from '../site/repl/evaluate.ts';

const read = (relative: string): Promise<string> =>
  readFile(fileURLToPath(new URL(relative, import.meta.url)), 'utf8');

const buildFunctions = async (): Promise<kit.PresentationData> => {
  const entries = Object.entries(kit);
  const pres = kit.createPresentation();
  const names = entries.map(([name]) => name);
  const code = await read('../site/repl/default-deck.js');
  await new Function(...names, 'pres', asyncBody(code))(...entries.map(([, v]) => v), pres);
  return pres;
};

const buildTsx = async (): Promise<kit.PresentationData> => {
  const root = await evaluateTsx(await read('../packages/dsl/examples/board-deck.tsx'), {
    '@office-kit/pptx': kit,
    '@office-kit/pptx-dsl': dsl,
    '@office-kit/pptx-dsl/jsx-runtime': jsxRuntime,
  });
  return dsl.compile(root as dsl.Node);
};

// Slides and charts carry everything either starter authors; the rest of the
// package (theme, masters, layouts) comes from the same blank deck.
const authoredParts = async (pres: kit.PresentationData): Promise<Map<string, string>> => {
  const reloaded = await kit.loadPresentation(await kit.savePresentation(pres));
  const parts = new Map<string, string>();
  for (const { name } of kit.listPackageParts(reloaded)) {
    if (!/^\/ppt\/(slides|charts)\/[^/]+\.xml$/.test(name)) continue;
    const bytes = kit.readPackagePart(reloaded, name);
    if (bytes === null) throw new Error(`listed part ${name} has no bytes`);
    // One element per line, so a mismatch reads as a diff instead of two walls of XML.
    parts.set(name, new TextDecoder().decode(bytes).replaceAll('><', '>\n<'));
  }
  return parts;
};

describe('REPL starters', () => {
  it('build the same deck from Functions and from TSX', async () => {
    const functions = await authoredParts(await buildFunctions());
    const tsx = await authoredParts(await buildTsx());
    // Six slides and two charts; an empty comparison must not pass.
    expect(functions.size).toBe(8);
    expect([...tsx.keys()].sort()).toEqual([...functions.keys()].sort());
    for (const [name, xml] of functions) {
      expect(tsx.get(name), name).toBe(xml);
    }
  });
});

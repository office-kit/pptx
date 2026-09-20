// The REPL's TSX tab evaluates editor text in the browser. Run the same
// evaluator here against the library source, so a DSL rename or a change in
// how Sucrase emits modules fails in CI instead of on the live site.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import * as kit from '../src/api/index.ts';
import * as dsl from '../packages/dsl/src/index.ts';
import * as jsxRuntime from '../packages/dsl/src/jsx-runtime.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import { WRAPPER_LINES, evaluateTsx } from '../site/repl/evaluate.ts';

const starterPath = fileURLToPath(new URL('../packages/dsl/examples/review.tsx', import.meta.url));
const MODULES = {
  '@office-kit/pptx': kit,
  '@office-kit/pptx-dsl': dsl,
  '@office-kit/pptx-dsl/jsx-runtime': jsxRuntime,
};

const build = async (source: string): Promise<kit.PresentationData> => {
  const root = await evaluateTsx(source, MODULES);
  return dsl.compile(root as dsl.Node);
};

describe('REPL TSX mode', () => {
  it('builds and renders the starter deck', async () => {
    const pres = await build(await readFile(starterPath, 'utf8'));
    const slides = kit.getSlides(pres);
    expect(slides.length).toBeGreaterThan(1);
    for (const slide of slides) {
      expect(renderSlideToSvg(pres, slide)).toContain('<svg');
    }
    const reloaded = await kit.loadPresentation(await kit.savePresentation(pres));
    expect(kit.getSlides(reloaded)).toHaveLength(slides.length);
  });

  it('runs type annotations, core-API imports and top-level await', async () => {
    const pres = await build(`
      import { inches } from '@office-kit/pptx';
      import { Presentation, Slide, Text } from '@office-kit/pptx-dsl';
      const titles: string[] = await Promise.resolve(['One', 'Two']);
      const width: number = inches(1) > 0 ? 9 : 0;
      export default (
        <Presentation>
          {titles.map((title) => (
            <Slide>
              <Text x={1} y={1} width={width} height={1}>{title}</Text>
            </Slide>
          ))}
        </Presentation>
      );
    `);
    expect(kit.getSlides(pres)).toHaveLength(2);
  });

  it('names the modules it provides when an import is unknown', async () => {
    await expect(build(`import React from 'react';\nexport default React;`)).rejects.toThrow(
      /Cannot import "react".*@office-kit\/pptx-dsl/,
    );
  });

  it('asks for a default export', async () => {
    await expect(build(`export const deck = 1;`)).rejects.toThrow(/default-export a Presentation/);
  });

  it('reports a parse error with its position', async () => {
    await expect(build(`const a = 1;\nconst b = );\nconst c = 2;\n`)).rejects.toThrow(/\(2:\d+\)/);
  });

  it('keeps editor line numbers in a runtime error', async () => {
    const error: unknown = await build(`const a = 1;\n\nnull.boom;\n`).catch((e: unknown) => e);
    expect(error).toBeInstanceOf(Error);
    const frame = /<anonymous>:(\d+):\d+/.exec((error as Error).stack ?? '');
    expect(Number(frame?.[1]) - WRAPPER_LINES).toBe(3);
  });
});

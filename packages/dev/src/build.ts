import { build, transform } from 'esbuild';
import { readFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { compile, type Node } from '@office-kit/pptx-dsl';
import {
  getSlideSize,
  getSlides,
  loadPresentation,
  savePresentation,
  validatePresentation,
} from '@office-kit/pptx';
import { renderSlideToSvg } from '@office-kit/pptx-preview';

const resolvePackage = (name: string) => fileURLToPath(import.meta.resolve(name));
export interface BuildResult {
  bytes: Uint8Array;
  slides: string[];
  aspectRatio: number;
  dependencies: string[];
  diagnostics: ReturnType<typeof validatePresentation>;
}
/** Evaluates trusted local TSX. This is code execution, not a sandbox. */
export async function buildDeck(entry: string, directory: string): Promise<BuildResult> {
  const output = join(directory, 'deck.mjs');
  const result = await build({
    entryPoints: [resolve(entry)],
    outfile: output,
    bundle: true,
    platform: 'node',
    format: 'esm',
    target: 'node22',
    jsx: 'automatic',
    jsxImportSource: '@office-kit/pptx-dsl',
    sourcemap: 'inline',
    metafile: true,
    logLevel: 'silent',
    plugins: [
      {
        name: 'source-location',
        setup(builder) {
          // A bundled module lives in a temporary directory. Keep source-relative
          // asset URLs pointing at each original module, including imported files.
          builder.onLoad({ filter: /\.[cm]?[jt]sx?$/ }, async ({ path }) => {
            const source = await readFile(path, 'utf8');
            const loader = path.endsWith('.tsx')
              ? 'tsx'
              : path.endsWith('.ts') || path.endsWith('.mts') || path.endsWith('.cts')
                ? 'ts'
                : path.endsWith('.jsx')
                  ? 'jsx'
                  : 'js';
            const transformed = await transform(source, {
              loader,
              sourcefile: path,
              sourcemap: 'inline',
              jsx: 'automatic',
              jsxDev: true,
              jsxImportSource: '@office-kit/pptx-dsl',
              define: { 'import.meta.url': JSON.stringify(pathToFileURL(path).href) },
            });
            return { contents: transformed.code, loader: 'js' };
          });
        },
      },
      {
        name: 'shared-office-kit',
        setup(builder) {
          // Core uses symbol-backed handles; all consumers must share its instance.
          builder.onResolve({ filter: /^@office-kit\/pptx(?:-dsl)?(?:\/.*)?$/ }, ({ path }) => ({
            path: resolvePackage(path),
            external: true,
          }));
        },
      },
    ],
  });
  const module: { default?: Node } = await import(pathToFileURL(output).href);
  if (!module.default) throw new Error('The TSX file must default-export a Presentation.');
  const presentation = await compile(module.default);
  const diagnostics = validatePresentation(presentation);
  const errors = diagnostics.filter((issue) => issue.severity === 'error');
  if (errors.length) throw new Error(`Invalid presentation: ${JSON.stringify(errors)}`);
  const bytes = await savePresentation(presentation);
  // Preview serialized output too, so persistence defects are visible during authoring.
  const saved = await loadPresentation(bytes);
  const size = getSlideSize(saved);
  return {
    bytes,
    aspectRatio: size ? size.width / size.height : 16 / 9,
    slides: getSlides(saved).map((slide) => renderSlideToSvg(saved, slide)),
    dependencies: Object.keys(result.metafile.inputs).map((path) => resolve(path)),
    diagnostics,
  };
}

import { context, transform } from 'esbuild';
import { readFile, writeFile } from 'node:fs/promises';
import { resolve, sep } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const resolvePackage = (name: string) => fileURLToPath(import.meta.resolve(name));

export async function createDeckCompiler(entry: string) {
  const transformedSources = new Map<string, { source: string; code: string }>();
  const compiler = await context({
    entryPoints: [resolve(entry)],
    outfile: 'deck.mjs',
    write: false,
    bundle: true,
    platform: 'node',
    format: 'esm',
    target: 'node22',
    jsx: 'automatic',
    jsxImportSource: '@office-kit/pptx-dsl',
    sourcemap: 'inline',
    sourceRoot: pathToFileURL(process.cwd() + sep).href,
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
            const cached = transformedSources.get(path);
            if (cached?.source === source) return { contents: cached.code, loader: 'js' };
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
            transformedSources.set(path, { source, code: transformed.code });
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
  return {
    async build(output: string) {
      const result = await compiler.rebuild();
      await writeFile(output, result.outputFiles[0]!.contents);
      const dependencies = Object.keys(result.metafile!.inputs).map((path) => resolve(path));
      const used = new Set(dependencies);
      for (const path of transformedSources.keys())
        if (!used.has(path)) transformedSources.delete(path);
      return dependencies;
    },
    cancel: () => compiler.cancel(),
    close: () => compiler.dispose(),
  };
}

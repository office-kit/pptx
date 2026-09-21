import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { build, transform } from 'esbuild';
import { compile, compileModule } from 'svelte/compiler';

// Bundle the site's canonical editor into the published CLI; no site server is needed.
await build({
  entryPoints: [fileURLToPath(new URL('../../site/src/lib/editor/dev/main.ts', import.meta.url))],
  outfile: fileURLToPath(new URL('./dist/editor.js', import.meta.url)),
  bundle: true,
  minify: true,
  format: 'esm',
  platform: 'browser',
  target: 'es2022',
  conditions: ['browser'],
  plugins: [
    {
      name: 'svelte-editor',
      setup(builder) {
        builder.onLoad({ filter: /\.svelte$/ }, async ({ path }) => ({
          contents: compile(await readFile(path, 'utf8'), {
            filename: path,
            generate: 'client',
            css: 'injected',
          }).js.code,
        }));
        builder.onLoad({ filter: /\.svelte\.ts$/ }, async ({ path }) => {
          const source = await transform(await readFile(path, 'utf8'), { loader: 'ts' });
          return {
            contents: compileModule(source.code, { filename: path, generate: 'client' }).js.code,
          };
        });
      },
    },
  ],
});

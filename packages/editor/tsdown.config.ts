import { readFile } from 'node:fs/promises';
import { transform } from 'esbuild';
import { compile, compileModule } from 'svelte/compiler';
import { defineConfig, type Rolldown } from 'tsdown';

// The editor is written in Svelte but ships as plain JavaScript: components are
// compiled here and the Svelte runtime is bundled (it is a devDependency), so a
// host page needs no framework. Component styles are injected at mount into the
// editor's shadow root (`css: 'injected'`).
const svelte: Rolldown.Plugin = {
  name: 'svelte',
  async load(id) {
    if (id.endsWith('.svelte')) {
      const source = await readFile(id, 'utf8');
      const { js } = compile(source, {
        filename: id,
        generate: 'client',
        css: 'injected',
        // No `window.__svelte` version registry on the host's page.
        discloseVersion: false,
      });
      return { code: js.code, moduleType: 'js' };
    }
    if (id.endsWith('.svelte.ts')) {
      const source = await transform(await readFile(id, 'utf8'), { loader: 'ts' });
      const { js } = compileModule(source.code, {
        filename: id,
        generate: 'client',
        discloseVersion: false,
      });
      return { code: js.code, moduleType: 'js' };
    }
    if (id.endsWith('.css?inline')) {
      const css = await readFile(id.slice(0, -'?inline'.length), 'utf8');
      return { code: `export default ${JSON.stringify(css)};`, moduleType: 'js' };
    }
    return null;
  },
};

export default defineConfig({
  entry: {
    index: 'src/index.ts',
    merge: 'src/merge.ts',
    internal: 'src/internal.ts',
    'internal/animation-player': 'src/internal-animation-player.ts',
  },
  format: ['esm'],
  target: 'es2022',
  platform: 'browser',
  inputOptions: { resolve: { conditionNames: ['browser', 'import', 'default'] } },
  plugins: [svelte],
  deps: { neverBundle: [/^@office-kit\//], onlyBundle: ['svelte', 'esm-env', 'clsx'] },
  dts: { oxc: true },
  sourcemap: true,
  clean: true,
});

import { defineConfig } from 'tsdown';
export default defineConfig({
  entry: { index: 'src/index.ts', cli: 'src/cli.ts', worker: 'src/worker.ts' },
  format: ['esm'],
  target: 'es2022',
  platform: 'node',
  // The editor is bundled like the browser pages bundle it, so installing
  // pptx-dev never needs @office-kit/pptx-editor from the registry.
  deps: { neverBundle: [/^@office-kit\/(?!pptx-editor)/, 'esbuild', /^node:/] },
  dts: true,
  sourcemap: true,
  clean: true,
});

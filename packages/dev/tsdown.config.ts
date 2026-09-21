import { defineConfig } from 'tsdown';
export default defineConfig({
  entry: { index: 'src/index.ts', cli: 'src/cli.ts', worker: 'src/worker.ts' },
  format: ['esm'],
  target: 'es2022',
  platform: 'node',
  deps: { neverBundle: [/^@office-kit\//, 'esbuild', /^node:/] },
  dts: true,
  sourcemap: true,
  clean: true,
});

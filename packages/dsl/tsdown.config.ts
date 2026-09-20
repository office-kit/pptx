import { defineConfig } from 'tsdown';
export default defineConfig({
  entry: { index: 'src/index.ts', 'jsx-runtime': 'src/jsx-runtime.ts' },
  format: ['esm'],
  target: 'es2022',
  platform: 'neutral',
  deps: { neverBundle: ['@office-kit/pptx'] },
  dts: true,
  sourcemap: true,
  clean: true,
});

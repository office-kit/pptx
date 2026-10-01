import { defineConfig } from 'tsdown';
export default defineConfig({
  entry: {
    index: 'src/index.ts',
    'jsx-runtime': 'src/jsx-runtime.ts',
    'source-edit': 'src/source-edit.ts',
  },
  format: ['esm'],
  target: 'es2022',
  platform: 'neutral',
  deps: { neverBundle: ['@office-kit/pptx', 'typescript'] },
  dts: true,
  sourcemap: true,
  clean: true,
});

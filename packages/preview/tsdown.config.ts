import { defineConfig } from 'tsdown';

export default defineConfig({
  entry: {
    index: 'src/index.ts',
    node: 'src/node.ts',
  },
  format: ['esm'],
  target: 'es2022',
  // `index` is browser-safe; `node` pulls in resvg + fontkit + node:fs. Never
  // bundle the peer/runtime deps so this package stays a thin layer over them.
  // Node built-ins are kept unbundled too, so the neutral build leaves them as
  // bare `node:*` imports rather than warning it can't resolve them.
  deps: {
    neverBundle: [
      '@office-kit/pptx',
      '@office-kit/pptx/node',
      '@resvg/resvg-js',
      'fontkit',
      /^node:/,
    ],
  },
  platform: 'neutral',
  dts: true,
  sourcemap: true,
  clean: true,
  treeshake: true,
});

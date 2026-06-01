import { defineConfig } from 'tsup';

export default defineConfig({
  entry: {
    index: 'src/index.ts',
    node: 'src/node.ts',
  },
  format: ['esm'],
  target: 'es2022',
  // `index` is browser-safe; `node` pulls in resvg + fontkit + node:fs. Keep
  // peer/runtime deps external so this package stays a thin layer over them.
  external: ['pptx-kit', 'pptx-kit/node', '@resvg/resvg-js', 'fontkit'],
  // The browser entry must never statically reach the node-only modules.
  // Code-splitting would hoist shared chunks across the two entries and could
  // drag `measure.ts` (node:fs) into the browser bundle — keep them separate.
  splitting: false,
  platform: 'neutral',
  dts: true,
  sourcemap: true,
  clean: true,
  treeshake: true,
});

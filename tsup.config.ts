import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { defineConfig } from 'tsup';

const pkg = JSON.parse(
  readFileSync(fileURLToPath(new URL('./package.json', import.meta.url)), 'utf8'),
) as { version: string };

export default defineConfig({
  entry: {
    index: 'src/index.ts',
    node: 'src/node.ts',
  },
  format: ['esm'],
  target: 'es2022',
  platform: 'neutral',
  dts: true,
  sourcemap: true,
  clean: true,
  treeshake: true,
  splitting: false,
  // Replace `__PPTX_KIT_VERSION__` in the source with the literal version
  // string from package.json. Avoids hand-syncing src/api/index.ts on
  // every release.
  define: {
    __PPTX_KIT_VERSION__: JSON.stringify(pkg.version),
  },
});

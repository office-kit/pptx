import { build } from 'esbuild';
import { fileURLToPath } from 'node:url';

// The merge bundles the core's package layer, whose TypeScript uses syntax
// Node's type stripping cannot run, so the tests load it the way tsdown ships it.
const { outputFiles } = await build({
  entryPoints: [fileURLToPath(new URL('../../src/deck-merge.ts', import.meta.url))],
  bundle: true,
  write: false,
  format: 'esm',
  platform: 'node',
  target: 'es2022',
});

export const { mergeDecks, describeConflict } = await import(
  `data:text/javascript;base64,${Buffer.from(outputFiles[0].contents).toString('base64')}`
);

import { copyFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { build } from 'esbuild';
await build({
  entryPoints: ['src/terminal-client.ts'],
  outfile: 'dist/terminal-client.js',
  bundle: true,
  minify: true,
  format: 'esm',
  platform: 'browser',
  target: 'es2022',
});

await copyFile(
  createRequire(import.meta.url).resolve('@xterm/xterm/css/xterm.css'),
  'dist/terminal-client.css',
);

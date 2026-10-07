import { rm } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';

await build({
  entryPoints: [fileURLToPath(new URL('./src/media-player.ts', import.meta.url))],
  outfile: fileURLToPath(new URL('./dist/media-player.js', import.meta.url)),
  bundle: true,
  minify: true,
  format: 'esm',
  platform: 'browser',
  target: 'es2022',
});

const resolveDir = fileURLToPath(new URL('.', import.meta.url));
const browserBundle = {
  bundle: true,
  minify: true,
  format: 'esm',
  platform: 'browser',
  target: 'es2022',
  conditions: ['browser'],
};

// The animation player is shared by three surfaces: the preview's presentation
// mode, the presenter window and the editor's animation panel. The panel gets
// it through the editor bundle below; the two pages load this build of it over
// HTTP, so all three run the same code.
await build({
  ...browserBundle,
  stdin: {
    contents: "export * from '@office-kit/pptx-editor/internal/animation-player';",
    resolveDir,
  },
  outfile: fileURLToPath(new URL('./dist/animation-player.js', import.meta.url)),
});

// Bundle the editor page into the published CLI; no site server is needed.
// Split so that what the editor loads on demand (the agent tool schemas) stays
// out of the page's first download; the server serves `editor-chunks/`.
const chunks = fileURLToPath(new URL('./dist/editor-chunks', import.meta.url));
await rm(chunks, { recursive: true, force: true });
await build({
  ...browserBundle,
  entryPoints: { editor: fileURLToPath(new URL('./src/editor-host.ts', import.meta.url)) },
  outdir: fileURLToPath(new URL('./dist', import.meta.url)),
  splitting: true,
  chunkNames: 'editor-chunks/[name]-[hash]',
});

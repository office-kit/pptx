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

// Bundle the editor into the published CLI; no site server is needed.
await build({
  ...browserBundle,
  stdin: {
    contents:
      "import { mountDevEditor } from '@office-kit/pptx-editor/internal'; mountDevEditor(document.body);",
    resolveDir,
  },
  outfile: fileURLToPath(new URL('./dist/editor.js', import.meta.url)),
});

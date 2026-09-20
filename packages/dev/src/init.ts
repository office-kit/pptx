import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';

const starter = `import { Presentation, Slide, Text, Shape } from '@office-kit/pptx-dsl';

export default (
  <Presentation>
    <Slide background="#15171C">
      <Shape preset="rect" x={0.9} y={2.55} width={0.14} height={1.75} fill="#E5481F" stroke={false} />
      <Text x={1.25} y={2.4} width={10} height={1.1} size={48} bold color="#FFFFFF">
        Your next presentation
      </Text>
      <Text x={1.25} y={3.55} width={10} height={0.6} size={22} color="#B4B9C4">
        Edit this TSX and save to update the preview.
      </Text>
    </Slide>
  </Presentation>
);
`;

const guide = `# Writing this presentation

Edit deck.tsx and local TypeScript components. The JSX runtime is
@office-kit/pptx-dsl, with no React or Vue dependency.

- Run npm run dev, then open its local URL to preview. Saving source files
  updates the preview. A failed build retains the last successful output.
- The preview is view-only: use the vertical thumbnails to select slides, zoom
  to inspect details, and Present to view full-screen (Escape exits). Make all
  content and layout changes in TSX; there are no canvas editing controls.
- Run npm run check for TypeScript diagnostics; run npm run build to export
  deck.pptx. Inspect every slide in the preview after changes.
- Coordinates and dimensions are inches; font sizes and stroke widths are points.
  The default 16:9 canvas is 13.333 × 7.5 inches.
- Use native Text, Shape, Image, Table and Chart elements. Use JS functions,
  arrays and map to compose elements; do not import React.
- Chart takes the core ChartSpec in its spec prop. Text accepts rich ParagraphSpec
  arrays through paragraphs. Raw is an escape hatch for public core APIs.
- For templates: readFile(new URL('./template.pptx', import.meta.url)), then use
  Presentation source={bytes} mode="edit" and Slide target={{index: 0}}.
  Fill target={{name: 'Title 1'}} changes an existing shape. Indices are zero-based.
  Use office-pptx inspect template.pptx to discover names, layouts and placeholders.
- Use mode="compose" and Slide from={{index: 0}} to build a new sequence from a
  source deck. This intentionally removes the original slide sequence.
- Keep assets inside this project so the dev server sees changes. Use source-relative
  new URL('./image.png', import.meta.url) when loading local assets.
- Preserve existing slides and formatting unless the request calls for changes.
  Never rasterize a chart/table as a shortcut. Raw-only features are not DSL coverage.
- Preview is a rendering aid, not a guarantee of identical PowerPoint rendering.
  Check the exported file in the target application for final delivery.

VSCode uses tsconfig.json for completion and diagnostics. Start the
"Preview presentation" task, then run "Simple Browser: Show" with the printed URL
and move that editor to a side group. Type errors appear in the Problems panel.
`;

/** Creates a new directory; existing projects are never overwritten. */
export async function initProject(directory: string): Promise<string> {
  const root = resolve(directory);
  const packageLocations = [
    import.meta.resolve('@office-kit/pptx/package.json'),
    import.meta.resolve('@office-kit/pptx-dsl/package.json'),
    new URL('../package.json', import.meta.url),
  ];
  const versions = await Promise.all(
    packageLocations.map(async (location) => {
      const manifest: { version: string } = JSON.parse(await readFile(new URL(location), 'utf8'));
      return `^${manifest.version}`;
    }),
  );
  await mkdir(root);
  const files: Record<string, string> = {
    'package.json':
      JSON.stringify(
        {
          name: 'my-presentation',
          private: true,
          type: 'module',
          scripts: {
            dev: 'office-pptx dev deck.tsx',
            build: 'office-pptx build deck.tsx --out deck.pptx',
            check: 'tsc --noEmit',
          },
          dependencies: { '@office-kit/pptx': versions[0]!, '@office-kit/pptx-dsl': versions[1]! },
          devDependencies: {
            '@office-kit/pptx-dev': versions[2]!,
            '@types/node': '^24.13.2',
            typescript: '^6.0.3',
          },
        },
        null,
        2,
      ) + '\n',
    'tsconfig.json':
      JSON.stringify(
        {
          compilerOptions: {
            target: 'ES2022',
            module: 'ESNext',
            moduleResolution: 'Bundler',
            jsx: 'react-jsx',
            jsxImportSource: '@office-kit/pptx-dsl',
            strict: true,
            exactOptionalPropertyTypes: true,
            noEmit: true,
            allowImportingTsExtensions: true,
            types: ['node'],
          },
          include: ['**/*.ts', '**/*.tsx'],
          exclude: ['node_modules'],
        },
        null,
        2,
      ) + '\n',
    'deck.tsx': starter,
    'CLAUDE.md': guide,
    '.gitignore': 'node_modules/\ndeck.pptx\n',
    '.vscode/tasks.json':
      JSON.stringify(
        {
          version: '2.0.0',
          tasks: [
            {
              label: 'Preview presentation',
              type: 'shell',
              command: 'npm run dev',
              isBackground: true,
              problemMatcher: {
                pattern: { regexp: '^NEVER_MATCH$', file: 1, message: 2 },
                background: {
                  activeOnStart: true,
                  beginsPattern: '^.*office-pptx dev',
                  endsPattern: '^Preview: ',
                },
              },
            },
            {
              label: 'Check presentation',
              type: 'shell',
              command: 'npm run check',
              problemMatcher: '$tsc',
            },
            {
              label: 'Export presentation',
              type: 'shell',
              command: 'npm run build',
              problemMatcher: [],
            },
          ],
        },
        null,
        2,
      ) + '\n',
  };
  await mkdir(join(root, '.vscode'));
  await Promise.all(
    Object.entries(files).map(([name, content]) =>
      writeFile(join(root, name), content, { flag: 'wx' }),
    ),
  );
  return root;
}

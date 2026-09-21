// /llms.txt — self-contained reference for AI assistants.
//
// The llmstxt.org proposal suggests a short index. We extend that: the file
// keeps the index shape (H1 title, blockquote summary, link sections at the
// end) but also inlines common core API guidance. For TSX and the long-form
// prose docs an agent can still follow the links at the bottom, append
// `.md` to any docs URL for raw Markdown, or fetch `/llms-full.txt` for
// the whole site concatenated into a single document.

import { base } from '$app/paths';
import { docSections } from '$lib/docs-nav';
import type { RequestHandler } from './$types';

export const prerender = true;

const HEADER = `# @office-kit/pptx

> Generate and edit PowerPoint \`.pptx\` (OOXML PresentationML) files from
> Node 22.18+ and modern browsers, with no PowerPoint, no Python, and no
> native runtime dependencies. Round-trip safety is the design center —
> unknown extensions are preserved verbatim on save.

This file is written for AI assistants. It contains a core API quick reference.
For the TSX authoring workflow, follow the authoring and TSX guides. The link
index at the bottom points at canonical docs — append \`.md\` to any docs
URL for raw Markdown (e.g. \`/docs/install.md\`), or fetch \`/llms-full.txt\`
for every page concatenated into a single document.

## TSX authoring and preview

The companion packages \`@office-kit/pptx-dsl\` and \`@office-kit/pptx-dev\`
provide typed TSX, project initialization, an interactive slide preview, and export.
They are published on npm. For Claude Code, install the plugin with:

\`\`\`sh
claude plugin marketplace add office-kit/skills
claude plugin install pptx@office-kit
\`\`\`

Invoke \`/pptx:office-kit-pptx\` with the presentation brief. The skill handles setup,
authoring, preview and export. Enable auto-update once in /plugin → Marketplaces
→ office-kit. Reload plugins or restart to load updates. Existing project npm
dependencies are not automatically upgraded. To initialize manually, run
\`npx --yes @office-kit/pptx-dev@latest init my-slides\`, then run
\`npm install\` inside the generated directory.

In an initialized project, edit the relevant \`slides/*.tsx\` file (read \`CLAUDE.md\` first).
Use \`deck.tsx\` for slide order and \`theme.ts\` for shared design values,
keep \`npm run dev\` running, and use \`npm run check\` plus \`npm run build\`
before delivery. The output is \`deck.pptx\`. The viewer has vertical thumbnails,
zoom and presentation mode; all content changes happen in TSX. The TSX guide
covers native elements, template references, and Raw callbacks for core APIs.

## Runtime

- Node \`>= 22.18\` (uses built-in Web Streams, Blob, fetch).
- Modern browsers, Bun, Deno, Cloudflare Workers, edge runtimes — anywhere
  with \`fetch\` + Uint8Array.
- ESM-only. \`"sideEffects": false\` — fully tree-shakable.
- The only runtime dependency is \`fflate\` (deflate / inflate). XML is
  parsed and serialized with a hand-rolled namespace-aware AST.

## Install

\`\`\`sh
pnpm add @office-kit/pptx
# or: npm install @office-kit/pptx / bun add @office-kit/pptx
\`\`\`

TypeScript: \`tsconfig.json\` should use \`"moduleResolution": "bundler"\`
(or \`"node16"\` / \`"nodenext"\`) so subpath entries resolve.

## One way to do one thing

@office-kit/pptx ships a **single fn-only public surface**. Every capability has
exactly one canonical function — no classes, no parallel APIs. The complete
list of public exports lives at \`/api\` on the site; the most common
entries are inlined below.

## Entries

- \`@office-kit/pptx\` — the full library. Runs in Node and the browser.
- \`@office-kit/pptx/node\` — adds \`loadPresentationFile\` / \`savePresentationToFile\`
  on top of the full library. Node-only.

## Units

OOXML measures positions in EMUs (1 inch = 914 400 EMU). @office-kit/pptx brands
\`Emu\` as a nominal type:

\`\`\`ts
import { inches, cm, mm, pt, emu, type Emu } from '@office-kit/pptx';

inches(1) // Emu (914 400)
cm(2.54)  // Emu (914 400)
pt(72)    // Emu (914 400)
emu(914400) // Emu — escape hatch
\`\`\`

\`TextFormat.size\` is in **points** (number), not EMU.

## Load + save round-trip

\`\`\`ts
import { loadPresentation, savePresentation } from '@office-kit/pptx';

const pres = await loadPresentation(bytes); // Uint8Array | ArrayBuffer | Blob
// ...mutate...
const out: Uint8Array = await savePresentation(pres);
\`\`\`

\`createPresentation()\` returns an empty deck with a slide master, the Office
theme, and three layouts (Blank, Title Slide, Title and Content). No template
file is needed to author a new deck.

## Template fill

\`\`\`ts
import { replaceTokensInPresentation } from '@office-kit/pptx';

replaceTokensInPresentation(pres, {
  name: 'Yamashita',
  event: 'Re:Invent',
});
\`\`\`

Walks every text run on every slide (and the notes parts). Tokens are
exact-match: \`{{name}}\` replaces only the literal \`{{name}}\` string.

## Slides

\`\`\`ts
import {
  addSlide,
  addBlankSlide,
  addTitleSlide,
  duplicateSlide,
  findSlideLayout,
  getSlides,
  moveSlide,
  removeSlide,
  setSlideTitle,
} from '@office-kit/pptx';

const layout = findSlideLayout(pres, 'Title and Content');
const slide = addSlide(pres, { layout });
setSlideTitle(slide, 'Q3 review');
\`\`\`

## Shapes

\`\`\`ts
import {
  addSlideShape,
  addSlideTextBox,
  addSlideLine,
  addSlideImage,
  addSlideTable,
  addSlideChart,
  inches,
  pt,
  setShapeFill,
  setShapeShadow,
} from '@office-kit/pptx';

const star = addSlideShape(slide, {
  preset: 'star5', // 180+ presets: rect, ellipse, triangle, star4..32, arrows...
  x: inches(1), y: inches(1), w: inches(2), h: inches(2),
  text: '★',
});
setShapeFill(star, '#FFD966');
setShapeShadow(star, {
  blurEmu: pt(8),
  offsetEmu: pt(4),
  angleDeg: 45,
  color: '#000000',
  opacity: 0.5,
});
\`\`\`

## Text formatting

Text is indexed by \`(paragraphIndex, runIndex)\`. Each \`\\n\` in the
\`text:\` argument starts a new paragraph.

\`\`\`ts
import { setShapeRunFormat, setParagraphAlignment, setParagraphBullet } from '@office-kit/pptx';

setShapeRunFormat(box, 1, 0, {
  bold: true,
  size: 24,            // points
  color: '#C00000',
  font: 'Calibri',
});
setParagraphAlignment(box, 1, 'ctr');
setParagraphBullet(box, 1, { char: '•' });
\`\`\`

## Tables

\`\`\`ts
import { addSlideTable, inches } from '@office-kit/pptx';

addSlideTable(slide, {
  x: inches(1), y: inches(1.5), w: inches(8), h: inches(3),
  rows: [
    ['Quarter', 'Revenue', 'Margin'],
    ['Q1', '$1.2M', '33%'],
    ['Q2', '$1.8M', '50%'],
  ],
  firstRow: true,
  bandRow: true,
});
\`\`\`

## Charts

\`\`\`ts
import { addSlideChart, inches } from '@office-kit/pptx';

addSlideChart(slide, {
  x: inches(1), y: inches(1.5), w: inches(8), h: inches(4.5),
  spec: {
    // bar | column | line | area | pie | doughnut | radar | stock | surface
    // (scatter | bubble take per-series xValues instead of categories)
    kind: 'column',
    categories: ['Q1', 'Q2', 'Q3', 'Q4'],
    series: [
      { name: 'Revenue', values: [120, 180, 240, 300] },
      { name: 'Cost',    values: [80,  90,  130, 160] },
    ],
    title: 'FY26',
  },
});
\`\`\`

\`addSlideChart\` generates the chart XML, the drawing rels, **and** the
embedded xlsx that PowerPoint needs for "Edit data".

Every ECMA-376 plot type is authorable. The kind names the data shape; modifiers
on the spec pick the variant:

\`\`\`ts
// Scatter / bubble: each series carries its own x channel.
{ kind: 'scatter', categories: [], scatterStyle: 'lineMarker',
  series: [{ name: 'Trial', xValues: [1, 2, 3], values: [2.5, 4.1, 6.2] }] }
{ kind: 'bubble', categories: [],
  series: [{ name: 'Markets', xValues: [10, 20], values: [5, 9], bubbleSizes: [100, 250] }] }

// 3-D bar / column / line / area / pie: add view3D.
{ kind: 'column', categories, series, view3D: { rotX: 20, rotY: 30 }, bar3DShape: 'cylinder' }

// Pie-of-pie / bar-of-pie.
{ kind: 'pie', categories, series: [one], ofPie: { type: 'bar', splitType: 'pos', splitPos: 3 } }

// Stock: series by position — high, low, close, or open first for candlesticks.
{ kind: 'stock', categories: days, series: [open, high, low, close] }

// Surface (surfaceContour: true for the top-down contour plot).
{ kind: 'surface', categories: xs, series: rows }
\`\`\`

Also on the spec: per-series \`errorBars\` / \`trendline\` / \`fillOpacity\`,
\`dataTable\`, \`categoryAxisDate\` (date axis), \`categoryGroupLevels\`
(multi-level categories), \`upDownBars\`, \`plotAreaLayout\`, combo charts via
per-series \`chartKind\` / \`secondaryAxis\`. A spec whose fields contradict
each other throws instead of writing a chart PowerPoint would repair.

## Images

\`\`\`ts
import { addSlideImage, setShapeImage, inches } from '@office-kit/pptx';

addSlideImage(slide, pngBytes, {
  x: inches(1), y: inches(1), w: inches(3), h: inches(2),
});

// Or swap an existing template image's bytes — geometry preserved:
setShapeImage(pictureShape, newBytes);
\`\`\`

Formats: PNG, JPEG, GIF, SVG, BMP, TIFF — detected from magic bytes; pass
\`options.format\` to override.

## Video, audio, online video

\`\`\`ts
import { addSlideMedia, getShapeMedia, inches } from '@office-kit/pptx';

const box = { x: inches(1), y: inches(1.5), w: inches(8), h: inches(4.5) };

// Embedded video / audio, from bytes. The container is detected from the bytes.
const clip = addSlideMedia(slide, { kind: 'video', data: mp4Bytes, poster: posterPngBytes, ...box });
addSlideMedia(slide, { kind: 'audio', data: mp3Bytes, x: inches(0.5), y: inches(6.5), w: inches(0.6), h: inches(0.6) });

// Online video. YouTube watch / youtu.be / shorts URLs become the embed URL.
addSlideMedia(slide2, { kind: 'online', url: 'https://www.youtube.com/watch?v=dQw4w9WgXcQ', ...box });

const media = getShapeMedia(clip);
// { kind: 'video', partName: '/ppt/media/media1.mp4', contentType: 'video/mp4', bytes }
// an online video reads back as { kind: 'online', url }
\`\`\`

The shape is a picture showing the poster frame (\`setShapeImage\` replaces it;
a play-button poster is used when \`poster\` is omitted). Containers detected
from magic bytes: mp4, m4v, mov, webm, avi, wmv, mp3, wav, m4a, ogg, wma; pass
\`format\` to override. Identical clip bytes are stored once per deck.

## Notes, comments, transitions, animations

\`\`\`ts
import {
  addSlideComment,
  cm,
  setShapeAnimation,
  setSlideNotes,
  setSlideTransition,
} from '@office-kit/pptx';

setSlideNotes(slide, 'Speaker note here.');
setSlideTransition(slide, { effect: 'fade', speed: 'med' });
setShapeAnimation(shape, { effect: 'fadeIn', durationMs: 700 });
addSlideComment(slide, {
  author: { name: 'Reviewer', initials: 'R' },
  text: 'Tighten the headline.',
  position: { x: cm(2), y: cm(2) },
});
\`\`\`

Animation presets: \`fadeIn\` / \`fadeOut\` / \`appear\` / \`disappear\`. Deeper
timing-tree authoring is post-1.0.

Transition effects: \`none\` / \`fade\` / \`push\` / \`cover\` / \`wipe\` /
\`split\` / \`cut\` / \`dissolve\` / \`checker\` / \`blinds\` / \`randomBar\` /
\`zoom\` / \`circle\` / \`diamond\` / \`plus\` / \`wedge\` / \`newsflash\`.

## Document metadata

\`\`\`ts
import { setCoreProperties, setExtendedProperties } from '@office-kit/pptx';

setCoreProperties(pres, {
  title: 'Q3 review',
  creator: 'Yamashita',
  subject: 'Revenue analysis',
  keywords: 'finance, planning',
});
\`\`\`

## Validation

\`\`\`ts
import { validatePresentation } from '@office-kit/pptx';

const issues = validatePresentation(pres);
for (const i of issues) console.error(i.severity, i.message);
\`\`\`

Catches: missing rels, dangling slide ids, layouts without masters,
off-spec ID ranges. Every emitted XML part is **also** schema-validated
against the ECMA-376 XSDs in the project's own CI.

## Node fs helpers

\`\`\`ts
import { loadPresentationFile, savePresentationToFile } from '@office-kit/pptx/node';

const pres = await loadPresentationFile('./template.pptx');
// ...mutate...
await savePresentationToFile(pres, './out.pptx');
\`\`\`

## What is **not** in 1.0

- Constructing new themes / masters / layouts from scratch (read-only).
- SmartArt authoring (preserved verbatim on round-trip).
- Complex animation timing-tree authoring.
- OLE / ActiveX authoring.
- Document encryption (read or write).

## Bundle budgets

- Minimal \`load → save\` bundle: ~61 KB unminified.
- Full fn-API bundle: ~122 KB unminified.

Tree-shaking is enforced by a CI test; \`"sideEffects": false\` is set in
the published \`package.json\`.

## Links
`;

const FOOTER = `

## Append-.md trick

Every docs page is available as raw Markdown by appending \`.md\` to the
URL. The same content powers this index, the long-form docs, and the
\`/llms-full.txt\` concatenation.

## Source

- [GitHub repository](https://github.com/office-kit/pptx)
- [npm package](https://www.npmjs.com/package/@office-kit/pptx)
`;

function buildBody(siteBase: string): string {
  const sections = docSections
    .map((section) => {
      const lines = [
        `### ${section.title}`,
        ...section.links.map((l) => `- [${l.title}](${siteBase}${l.href}.md): ${l.description}`),
      ];
      return lines.join('\n');
    })
    .join('\n\n');
  return `${HEADER}\n${sections}${FOOTER}`;
}

export const GET: RequestHandler = () => {
  return new Response(buildBody(base), {
    headers: {
      'Content-Type': 'text/plain; charset=utf-8',
      'Cache-Control': 'public, max-age=300',
    },
  });
};

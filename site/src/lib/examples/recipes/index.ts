// Recipe registry. Each entry pairs human-readable framing with the
// literal source of a real .ts file in this directory; the file is
// type-checked against `@office-kit/pptx` on every build, so the snippet shown
// to readers can never drift from the live API.

import addChart from './add-chart.ts?raw';
import addImage from './add-image.ts?raw';
import addTable from './add-table.ts?raw';
import buildFromScratch from './build-from-scratch.ts?raw';
import fillsAndEffects from './fills-and-effects.ts?raw';
import groupShapesRecipe from './group-shapes.ts?raw';
import hyperlinks from './hyperlinks.ts?raw';
import inspectParts from './inspect-parts.ts?raw';
import notesAndComments from './notes-and-comments.ts?raw';
import openAndIterate from './open-and-iterate.ts?raw';
import presetShapes from './preset-shapes.ts?raw';
import textFormatting from './text-formatting.ts?raw';
import themeBranding from './theme-branding.ts?raw';
import transitionsAnimations from './transitions-animations.ts?raw';
import validate from './validate.ts?raw';

import browserFetch from '../browser-fetch.ts?raw';
import imageReplace from '../image-replace.ts?raw';
import nodeFs from '../node-fs.ts?raw';
import templateFill from '../template-fill.ts?raw';

export type Recipe = {
  /** URL-safe slug, used as the anchor on /docs/recipes. */
  slug: string;
  title: string;
  /** One-line teaser shown above the code block. */
  teaser: string;
  /** Repo path (used as caption above the snippet). */
  path: string;
  /** Verbatim source. */
  source: string;
  /** Optional bullet points to add context after the snippet. */
  notes?: string[];
  /** Names of related public exports — link to /api anchors when available. */
  relatedApi?: string[];
};

export const recipeGroups: Array<{ title: string; recipes: Recipe[] }> = [
  {
    title: 'Basics',
    recipes: [
      {
        slug: 'open-and-iterate',
        title: 'Open a presentation and walk every slide',
        teaser: 'Load an existing .pptx, iterate slides, and read title + text content.',
        path: 'site/src/lib/examples/recipes/open-and-iterate.ts',
        source: openAndIterate,
        notes: [
          'getSlideTitle returns the title placeholder text or null if the slide has no title placeholder.',
          'getSlideText concatenates every text body on the slide; getSlideOutline gives a structured outline if you need paragraph hierarchy.',
        ],
        relatedApi: ['loadPresentation', 'getSlides', 'getSlideTitle', 'getSlideText'],
      },
      {
        slug: 'template-fill',
        title: 'Token-based template fill',
        teaser:
          'The canonical L2 workflow: open a designer-authored template, swap `{{tokens}}`, save.',
        path: 'site/src/lib/examples/template-fill.ts',
        source: templateFill,
        notes: [
          'Tokens are exact-match: {{name}} replaces only the literal `{{name}}` string in any text run.',
          'replaceTokensInPresentation also walks speaker notes; use replaceTokensInSlide for a single slide.',
        ],
        relatedApi: [
          'loadPresentation',
          'replaceTokensInPresentation',
          'replaceTokensInSlide',
          'savePresentation',
        ],
      },
      {
        slug: 'build-from-scratch',
        title: 'Build a deck on a blank template',
        teaser: 'Start from a tiny blank.pptx, add slides on its layouts, and save.',
        path: 'site/src/lib/examples/recipes/build-from-scratch.ts',
        source: buildFromScratch,
        relatedApi: ['addSlide', 'findSlideLayout', 'findSlidePlaceholder', 'setShapeText'],
      },
      {
        slug: 'theme-branding',
        title: 'Brand a deck’s color scheme + fonts',
        teaser:
          'setPresentationTheme / setPresentationFonts patch the theme’s color and font scheme — no template required.',
        path: 'site/src/lib/examples/recipes/theme-branding.ts',
        source: themeBranding,
        notes: [
          'Only the slots you pass are overwritten; every other color/typeface keeps its default.',
          'Colors are #RRGGBB strings; every theme slot is always a plain srgbClr, never a scheme-color reference.',
        ],
        relatedApi: [
          'setPresentationTheme',
          'setPresentationFonts',
          'getPresentationTheme',
          'getPresentationFonts',
        ],
      },
      {
        slug: 'node-fs',
        title: 'Direct fs helpers (Node)',
        teaser: 'loadPresentationFile / savePresentationToFile skip the manual fs glue.',
        path: 'site/src/lib/examples/node-fs.ts',
        source: nodeFs,
        relatedApi: ['loadPresentationFile', 'savePresentationToFile'],
      },
    ],
  },
  {
    title: 'Authoring shapes',
    recipes: [
      {
        slug: 'text-formatting',
        title: 'Per-run text formatting',
        teaser:
          'Address text by (paragraph, run) indices and apply font, size, color, bold / italic / underline.',
        path: 'site/src/lib/examples/recipes/text-formatting.ts',
        source: textFormatting,
        notes: [
          'TextFormat.size is in points (number), not EMU. Color accepts #RRGGBB, RRGGBB, or a scheme token like "accent1".',
          'setShapeTextFormat (no index) applies the same format to every run in the shape.',
        ],
        relatedApi: ['setShapeRunFormat', 'setShapeTextFormat', 'setParagraphAlignment'],
      },
      {
        slug: 'preset-shapes',
        title: 'Preset shapes (180+ presets)',
        teaser: 'addSlideShape ships every ECMA-376 preset geometry — pass the token as `preset`.',
        path: 'site/src/lib/examples/recipes/preset-shapes.ts',
        source: presetShapes,
        notes: [
          'Use PresetShape from the public types for autocompletion. Unknown strings pass through to <a:prstGeom prst="..."/> verbatim.',
        ],
        relatedApi: ['addSlideShape', 'PresetShape'],
      },
      {
        slug: 'fills-and-effects',
        title: 'Fills, shadows, glows',
        teaser: 'Solid / gradient / pattern fills plus the two effect helpers.',
        path: 'site/src/lib/examples/recipes/fills-and-effects.ts',
        source: fillsAndEffects,
        relatedApi: [
          'setShapeFill',
          'setShapeGradientFill',
          'setShapePatternFill',
          'setShapeShadow',
          'setShapeGlow',
        ],
      },
      {
        slug: 'group-shapes',
        title: 'Group shapes into one component',
        teaser: 'Compose a rectangle + label into a "KPI card"; move/resize the group as one unit.',
        path: 'site/src/lib/examples/recipes/group-shapes.ts',
        source: groupShapesRecipe,
        notes: [
          "groupShapes needs every member to have an explicit position/size — placeholders that inherit geometry from the layout can't be grouped directly.",
          'ungroupShapes reverses it, rescaling each member if the group itself was moved/resized in between.',
        ],
        relatedApi: ['groupShapes', 'ungroupShapes', 'getGroupChildren', 'getGroupTransform'],
      },
    ],
  },
  {
    title: 'Tables, charts, images',
    recipes: [
      {
        slug: 'add-table',
        title: 'Add a table',
        teaser: 'A 2D array of cell strings, plus firstRow / bandRow flags for header styling.',
        path: 'site/src/lib/examples/recipes/add-table.ts',
        source: addTable,
        notes: [
          'Per-cell control after the fact: getTableCell + setTableCellText / setTableCellFill / setTableCellAlignment.',
          'insertTableRow / removeTableRow / insertTableColumn / removeTableColumn for structural changes.',
        ],
        relatedApi: ['addSlideTable', 'getTableCell', 'setTableCellText'],
      },
      {
        slug: 'add-chart',
        title: 'Add a chart',
        teaser:
          'addSlideChart writes the chart XML, the drawing rels, and an embedded xlsx so "Edit data" works in PowerPoint.',
        path: 'site/src/lib/examples/recipes/add-chart.ts',
        source: addChart,
        notes: [
          'Chart kinds: bar, column, line, pie, doughnut, area. Each takes the same { categories, series } shape.',
          'setChartSpec updates an existing chart in place — useful for re-rendering a template chart with new data.',
        ],
        relatedApi: ['addSlideChart', 'setChartSpec', 'getShapeChartSpec'],
      },
      {
        slug: 'add-image',
        title: 'Insert an image',
        teaser:
          'Format auto-detected from magic bytes. Reuse setShapeImage to swap an existing picture.',
        path: 'site/src/lib/examples/recipes/add-image.ts',
        source: addImage,
        relatedApi: ['addSlideImage', 'setShapeImage'],
      },
      {
        slug: 'image-replace',
        title: 'Replace a template image in place',
        teaser:
          'Find a placeholder picture by name and swap its bytes — geometry / crop preserved.',
        path: 'site/src/lib/examples/image-replace.ts',
        source: imageReplace,
        relatedApi: ['setShapeImage', 'getShapeKind', 'getShapeName'],
      },
    ],
  },
  {
    title: 'L4: notes, comments, transitions, animations',
    recipes: [
      {
        slug: 'notes-and-comments',
        title: 'Speaker notes + review comments',
        teaser:
          'Notes live on the slide; comments are separate parts with author metadata and an EMU position.',
        path: 'site/src/lib/examples/recipes/notes-and-comments.ts',
        source: notesAndComments,
        relatedApi: ['setSlideNotes', 'getSlideNotes', 'addSlideComment', 'getSlideComments'],
      },
      {
        slug: 'transitions-animations',
        title: 'Slide transitions + shape animations',
        teaser: 'Per-slide transition effect plus entrance / exit animations on individual shapes.',
        path: 'site/src/lib/examples/recipes/transitions-animations.ts',
        source: transitionsAnimations,
        notes: [
          'Animation effects: fadeIn, fadeOut, appear, disappear. Deeper timing-tree authoring is post-1.0.',
        ],
        relatedApi: ['setSlideTransition', 'setShapeAnimation'],
      },
      {
        slug: 'hyperlinks',
        title: 'Hyperlinks + click-to-slide',
        teaser: 'External URL on a text run, and a click action that jumps to another slide.',
        path: 'site/src/lib/examples/recipes/hyperlinks.ts',
        source: hyperlinks,
        relatedApi: ['setShapeHyperlink', 'setShapeClickAction'],
      },
    ],
  },
  {
    title: 'Diagnostics',
    recipes: [
      {
        slug: 'validate',
        title: 'Validate the package',
        teaser:
          'validatePresentation returns invariant violations: dangling rels, missing parts, off-spec IDs.',
        path: 'site/src/lib/examples/recipes/validate.ts',
        source: validate,
        relatedApi: ['validatePresentation'],
      },
      {
        slug: 'inspect-parts',
        title: 'Inspect raw OPC parts',
        teaser:
          'listPackageParts + readPackagePart let you peek at the underlying XML without dropping to the internal class.',
        path: 'site/src/lib/examples/recipes/inspect-parts.ts',
        source: inspectParts,
        relatedApi: ['listPackageParts', 'readPackagePart', 'getMediaParts'],
      },
    ],
  },
  {
    title: 'Browser',
    recipes: [
      {
        slug: 'browser-fetch',
        title: 'Browser: load via fetch',
        teaser:
          'loadPresentation accepts ArrayBuffer / Uint8Array / Blob — pipe a fetch Response straight in.',
        path: 'site/src/lib/examples/browser-fetch.ts',
        source: browserFetch,
        relatedApi: ['loadPresentation', 'getSlides', 'getSlideTitle'],
      },
    ],
  },
];

export const allRecipes: Recipe[] = recipeGroups.flatMap((g) => g.recipes);

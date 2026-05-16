// Registry of every example file. The ?raw imports give us the on-disk
// source verbatim; the same files are also type-checked by svelte-check
// against the real library on every build — if an API renames, the docs
// build fails.

import browserFetch from './browser-fetch.ts?raw';
import buildDeck from './build-deck.ts?raw';
import imageReplace from './image-replace.ts?raw';
import nodeFs from './node-fs.ts?raw';
import templateFill from './template-fill.ts?raw';

export type Example = {
  /** Human title for the snippet (shown above the code block). */
  title: string;
  /** Repo-relative path, also used for the file-tab caption. */
  path: string;
  /** Verbatim source text (already type-checked by svelte-check / tsc). */
  source: string;
  /** Short description used in docs. */
  description: string;
};

export const examples = {
  templateFill: {
    title: 'Token-based template fill',
    path: 'site/src/lib/examples/template-fill.ts',
    description:
      'Replace `{{tokens}}` across every slide and the speaker notes — the canonical full-library round-trip.',
    source: templateFill,
  },
  buildDeck: {
    title: 'Build a deck from a blank template',
    path: 'site/src/lib/examples/build-deck.ts',
    description:
      'Add slides on a layout, drop in a text box, an image, and a chart, then save.',
    source: buildDeck,
  },
  imageReplace: {
    title: 'Replace an image in place',
    path: 'site/src/lib/examples/image-replace.ts',
    description:
      'Swap the bytes of a picture shape while preserving its crop, transform, and sizing.',
    source: imageReplace,
  },
  nodeFs: {
    title: 'Direct fs helpers',
    path: 'site/src/lib/examples/node-fs.ts',
    description:
      'loadPresentationFile / savePresentationToFile from pptx-kit/node skip the manual fs glue.',
    source: nodeFs,
  },
  browserFetch: {
    title: 'Fetch in the browser',
    path: 'site/src/lib/examples/browser-fetch.ts',
    description:
      'Pipe a fetch Response straight through arrayBuffer() into loadPresentation.',
    source: browserFetch,
  },
} as const satisfies Record<string, Example>;

export type ExampleKey = keyof typeof examples;

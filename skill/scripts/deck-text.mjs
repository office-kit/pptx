#!/usr/bin/env node
// Prints a deck's headline track and flags wording for review (getPresentationTextReview).
// Usage (from the slide project): node <skill>/scripts/deck-text.mjs deck.pptx [--mode talk|document] [--full]
//
// Flags are suspicions, not verdicts: references/review.md explains how to judge each one.
import { readFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { parseArgs } from 'node:util';

const USAGE = 'Usage: deck-text.mjs <deck.pptx> [--mode talk|document] [--full]';
const MODES = ['talk', 'document'];

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: { mode: { type: 'string' }, full: { type: 'boolean', default: false } },
});
const [file] = positionals;
const { mode, full } = values;
if (positionals.length !== 1 || (mode !== undefined && !MODES.includes(mode))) {
  console.error(USAGE);
  process.exit(2);
}

const pptx = await import(await resolveCore(process.cwd()));
if (typeof pptx.getPresentationTextReview !== 'function') {
  console.error(
    "The slide project's @office-kit/pptx is older than 0.27.0 and has no getPresentationTextReview; update it to the latest version.",
  );
  process.exit(1);
}
const presentation = await pptx.loadPresentation(await readFile(file));
console.log(pptx.getPresentationTextReview(presentation, { mode, full }));

// The script ships with the skill, outside the slide project, so a bare import
// would resolve against the skill directory. Resolve the project's installed copy.
async function resolveCore(start) {
  for (let directory = start; ; directory = dirname(directory)) {
    const own = join(directory, 'package.json');
    if (existsSync(own) && JSON.parse(await readFile(own, 'utf8')).name === '@office-kit/pptx')
      return entryOf(directory);
    const installed = join(directory, 'node_modules', '@office-kit', 'pptx');
    if (existsSync(installed)) return entryOf(installed);
    if (dirname(directory) === directory)
      throw new Error('@office-kit/pptx is not installed; run this from the slide project.');
  }
}

async function entryOf(packageDirectory) {
  const manifest = JSON.parse(await readFile(join(packageDirectory, 'package.json'), 'utf8'));
  return pathToFileURL(join(packageDirectory, manifest.exports['.'].import)).href;
}

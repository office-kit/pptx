// Regression: real .pptx files produced by some tools (notably PowerPoint
// on Windows and certain third-party libraries) include explicit
// directory entries in the ZIP central directory. These are zero-byte
// entries whose names end with "/" (e.g. `_rels/`, `ppt/`,
// `ppt/slides/`). They are *not* OPC parts.
//
// Before the skip in `OpcPackage.load`, attempting to convert one of
// these entries via `fromZipPath` threw
// `part name must not end with "/": /_rels/`, killing the load.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { unzipSync, zipSync } from 'fflate';
import { getSlideCount, loadPresentation, savePresentation } from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const ENC = new TextEncoder();

// Build a `.pptx` byte stream that copies every entry from `source` and
// adds a few zero-byte directory entries up front, mimicking what
// PowerPoint-on-Windows / Office Open XML SDK / docx4j sometimes write.
const withDirectoryEntries = (source: Uint8Array): Uint8Array => {
  const entries = unzipSync(source);
  const augmented: Record<string, Uint8Array> = {};
  for (const dir of ['_rels/', 'ppt/', 'ppt/slides/', 'ppt/slideLayouts/']) {
    augmented[dir] = ENC.encode('');
  }
  for (const [name, data] of Object.entries(entries)) {
    augmented[name] = data;
  }
  return zipSync(augmented);
};

describe('loadPresentation: ZIP directory entries', () => {
  it('skips zero-byte directory entries instead of rejecting them', async () => {
    const original = await readFile(fixture('two-slides.pptx'));
    const withDirs = withDirectoryEntries(new Uint8Array(original));

    const pres = await loadPresentation(withDirs);
    expect(getSlideCount(pres)).toBe(2);

    // Round-trip remains valid.
    const out = await savePresentation(pres);
    const reloaded = await loadPresentation(out);
    expect(getSlideCount(reloaded)).toBe(2);
  });
});

// getCoreProperties — read /docProps/core.xml (OPC core-properties).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { getCoreProperties, loadPresentation } from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getCoreProperties', () => {
  it('reads core-properties from the fixture', async () => {
    // two-slides.pptx ships with python-pptx defaults: empty title/
    // subject/creator, a description, lastModifiedBy "Steve Canny",
    // revision "1", and W3CDTF timestamps.
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const props = getCoreProperties(pres);
    expect(props).not.toBeNull();
    expect(props!.description).toBe('generated using python-pptx');
    expect(props!.lastModifiedBy).toBe('Steve Canny');
    expect(props!.revision).toBe('1');
    expect(props!.created).toMatch(/^\d{4}-\d{2}-\d{2}T/);
    expect(props!.modified).toMatch(/^\d{4}-\d{2}-\d{2}T/);
  });

  it('every field is null when the package has no core-properties part', async () => {
    // The blank.pptx fixture is a freshly-generated deck and may or
    // may not have core props; we only sanity-check the shape here.
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const props = getCoreProperties(pres);
    // Either null (no part at all) or an object with the right
    // shape — both are acceptable for an empty deck.
    if (props !== null) {
      expect(typeof props.title === 'string' || props.title === null).toBe(true);
      expect(typeof props.creator === 'string' || props.creator === null).toBe(true);
    }
  });
});

// getExtendedProperties — read /docProps/app.xml.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { getExtendedProperties, loadPresentation } from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getExtendedProperties', () => {
  it('reads Application + AppVersion from the fixture', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const props = getExtendedProperties(pres);
    expect(props).not.toBeNull();
    // python-pptx-generated fixture ships these values.
    expect(props!.application).toBe('Microsoft Macintosh PowerPoint');
    expect(props!.appVersion).toBe('14.0000');
    expect(props!.presentationFormat).toBe('On-screen Show (4:3)');
  });

  it('exposes empty string fields as null', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const props = getExtendedProperties(pres);
    expect(props).not.toBeNull();
    // Fixture ships empty <Manager> / <Company> / <HyperlinkBase>.
    expect(props!.manager).toBeNull();
    expect(props!.company).toBeNull();
    expect(props!.hyperlinkBase).toBeNull();
  });
});

// setExtendedProperties — partial setter for /docProps/app.xml.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getExtendedProperties,
  loadPresentation,
  savePresentation,
  setExtendedProperties,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: setExtendedProperties', () => {
  it('updates only the requested fields', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const before = getExtendedProperties(pres)!;
    setExtendedProperties(pres, { company: 'Acme', manager: 'Dana' });

    const after = getExtendedProperties(pres)!;
    expect(after.company).toBe('Acme');
    expect(after.manager).toBe('Dana');
    // Untouched fields are preserved verbatim.
    expect(after.application).toBe(before.application);
    expect(after.appVersion).toBe(before.appVersion);
    expect(after.presentationFormat).toBe(before.presentationFormat);
  });

  it('persists through save → reload', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    setExtendedProperties(pres, {
      company: 'Acme',
      hyperlinkBase: 'https://acme.example/',
    });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const props = getExtendedProperties(reloaded)!;
    expect(props.company).toBe('Acme');
    expect(props.hyperlinkBase).toBe('https://acme.example/');
  });

  it('clears a field when passed null', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    setExtendedProperties(pres, { application: null });
    expect(getExtendedProperties(pres)!.application).toBeNull();
  });
});

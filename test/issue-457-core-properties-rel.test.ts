import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  createPresentation,
  getCoreProperties,
  getExtendedProperties,
  loadPresentation,
  savePresentation,
  setCoreProperties,
  setExtendedProperties,
  validatePresentation,
} from '../src/api/index.ts';

const CORE_RELATIONSHIP = /relationships\/metadata\/core-properties"/g;

// OPC locates these parts through the package's root relationships, not by
// name, so a producer may store them anywhere. `target` is written into
// `_rels/.rels` as is, so it can be relative to the package root or absolute.
const movePart = async (from: string, to: string, target = to) => {
  const pres = createPresentation();
  setCoreProperties(pres, { title: 'Quarterly review' });
  setExtendedProperties(pres, { company: 'Example' });
  const files = unzipSync(await savePresentation(pres));
  files[to] = files[from]!;
  delete files[from];
  files['_rels/.rels'] = strToU8(
    strFromU8(files['_rels/.rels']!).replace(`"${from}"`, `"${target}"`),
  );
  files['[Content_Types].xml'] = strToU8(
    strFromU8(files['[Content_Types].xml']!).replace(`/${from}`, `/${to}`),
  );
  return zipSync(files);
};

describe('issue-457: properties parts are found through the root relationships', () => {
  const moved = 'package/services/metadata/core-properties/abc.psmdcp';

  it('reads core properties stored under another name', async () => {
    const pres = await loadPresentation(await movePart('docProps/core.xml', moved));
    expect(getCoreProperties(pres)?.title).toBe('Quarterly review');
  });

  it('follows an absolute relationship target', async () => {
    const pres = await loadPresentation(await movePart('docProps/core.xml', moved, `/${moved}`));
    expect(getCoreProperties(pres)?.title).toBe('Quarterly review');
  });

  it('updates the existing core-properties part instead of adding a second one', async () => {
    const pres = await loadPresentation(await movePart('docProps/core.xml', moved));
    setCoreProperties(pres, { title: 'Renamed' });
    const out = await savePresentation(pres);
    const files = unzipSync(out);
    expect(strFromU8(files['_rels/.rels']!).match(CORE_RELATIONSHIP)).toHaveLength(1);
    expect(files['docProps/core.xml']).toBeUndefined();
    expect(strFromU8(files[moved]!)).toContain('Renamed');
    const reloaded = await loadPresentation(out);
    expect(getCoreProperties(reloaded)?.title).toBe('Renamed');
    expect(validatePresentation(reloaded).filter((issue) => issue.severity === 'error')).toEqual(
      [],
    );
  });

  it('reads and writes extended properties stored under another name', async () => {
    const pres = await loadPresentation(await movePart('docProps/app.xml', 'meta/app-props.xml'));
    expect(getExtendedProperties(pres)?.company).toBe('Example');
    setExtendedProperties(pres, { company: 'Renamed' });
    const files = unzipSync(await savePresentation(pres));
    expect(files['docProps/app.xml']).toBeUndefined();
    expect(strFromU8(files['meta/app-props.xml']!)).toContain('<Company>Renamed</Company>');
  });

  it('ignores a /docProps/core.xml part that no relationship points at', async () => {
    const pres = createPresentation();
    setCoreProperties(pres, { title: 'Quarterly review' });
    const files = unzipSync(await savePresentation(pres));
    files['_rels/.rels'] = strToU8(
      strFromU8(files['_rels/.rels']!).replace(/<Relationship [^>]*core-properties"[^>]*\/>/, ''),
    );
    const unlinked = await loadPresentation(zipSync(files));
    expect(getCoreProperties(unlinked)).toBeNull();

    setCoreProperties(unlinked, { title: 'Linked' });
    const out = await savePresentation(unlinked);
    expect(strFromU8(unzipSync(out)['_rels/.rels']!).match(CORE_RELATIONSHIP)).toHaveLength(1);
    const reloaded = await loadPresentation(out);
    expect(getCoreProperties(reloaded)).toMatchObject({ title: 'Linked', creator: null });
    expect(validatePresentation(reloaded).filter((issue) => issue.severity === 'error')).toEqual(
      [],
    );
  });

  it('replaces a core-properties relationship whose part is missing', async () => {
    const pres = createPresentation();
    const files = unzipSync(await savePresentation(pres));
    delete files['docProps/core.xml'];
    const dangling = await loadPresentation(zipSync(files));
    setCoreProperties(dangling, { title: 'Restored' });
    const out = await savePresentation(dangling);
    expect(strFromU8(unzipSync(out)['_rels/.rels']!).match(CORE_RELATIONSHIP)).toHaveLength(1);
    expect(getCoreProperties(await loadPresentation(out))?.title).toBe('Restored');
  });
});

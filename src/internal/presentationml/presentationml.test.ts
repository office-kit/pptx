import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { partName } from '../opc/index.ts';
import { OpcPackage } from '../parts/index.ts';
import { NS, parseFragment, parseXml } from '../xml/index.ts';
import { readPresentationPart } from './index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`../../../test/fixtures/minimal/${name}`, import.meta.url));

const loadPresentationXml = async (pptxName: string) => {
  const pkg = OpcPackage.load(await readFile(fixture(pptxName)));
  const part = pkg.getPart(partName('/ppt/presentation.xml'));
  if (!part) throw new Error('presentation.xml not found');
  return parseXml(new TextDecoder().decode(part.data)).root;
};

describe('readPresentationPart', () => {
  it('parses the minimal blank.pptx presentation', async () => {
    const root = await loadPresentationXml('blank.pptx');
    const pres = readPresentationPart(root);

    // python-pptx default template ships with one slide master, no slides.
    expect(pres.slideMasters.length).toBe(1);
    expect(pres.slideMasters[0]?.rId).toMatch(/^rId\d+$/);
    expect(pres.slideMasters[0]?.id).toBeGreaterThanOrEqual(2147483648);

    expect(pres.slides).toEqual([]);

    // python-pptx's blank template omits `p:notesMasterIdLst`. The notes
    // master is reachable via the slide-master's rels graph instead.
    expect(pres.notesMaster).toBeNull();

    expect(pres.slideSize).not.toBeNull();
    expect(pres.slideSize?.cx).toBeGreaterThan(0);
    expect(pres.slideSize?.cy).toBeGreaterThan(0);
  });

  it('parses two-slides.pptx with two slides registered', async () => {
    const root = await loadPresentationXml('two-slides.pptx');
    const pres = readPresentationPart(root);

    expect(pres.slides.length).toBe(2);
    // sldId values must be ≥256 and unique within the list.
    const ids = pres.slides.map((s) => s.id);
    expect(ids.every((id) => id >= 256)).toBe(true);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it('throws on the wrong root element', () => {
    const wrong = parseFragment(`<p:bogus xmlns:p="${NS.pml}"/>`);
    expect(() => readPresentationPart(wrong)).toThrow();
  });

  it('throws when sldMasterId lacks the r:id attribute', () => {
    const wrong = parseFragment(
      `<p:presentation xmlns:p="${NS.pml}" xmlns:r="${NS.officeDocRels}"><p:sldMasterIdLst><p:sldMasterId id="2147483648"/></p:sldMasterIdLst></p:presentation>`,
    );
    expect(() => readPresentationPart(wrong)).toThrow(/r:id/);
  });
});

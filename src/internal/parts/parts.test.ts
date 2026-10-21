import { describe, expect, it } from 'vitest';
import {
  ONE_PIXEL_PNG_BYTES,
  SYNTHETIC_PACKAGE_PART_NAMES,
  SYNTHETIC_REL_TYPES,
  buildSyntheticPackageBytes,
} from '../../../test/lib/synthetic-package.ts';
import { partName, writeZip } from '../opc/index.ts';
import { OpcPackage } from './index.ts';

const decode = (b: Uint8Array): string => new TextDecoder().decode(b);

describe('OpcPackage.load', () => {
  it('returns independent relationship snapshots and observes replaced bytes', () => {
    const pkg = OpcPackage.load(buildSyntheticPackageBytes());
    const name = partName('/ppt/presentation.xml');
    const initial = pkg.getRels(name)!;
    const target = initial.items[0]!.target;
    initial.items[0]!.target = 'changed.xml';
    expect(pkg.getRels(name)!.items[0]!.target).toBe(target);
    pkg.setRels(name, initial);
    expect(pkg.getRels(name)!.items[0]!.target).toBe('changed.xml');
    const root = pkg.rootRels()!;
    const rootTarget = root.items[0]!.target;
    root.items[0]!.target = 'root.xml';
    expect(pkg.rootRels()!.items[0]!.target).toBe(rootTarget);
    pkg.setRootRels(root);
    expect(pkg.rootRels()!.items[0]!.target).toBe('root.xml');
  });

  it('loads a synthetic package and resolves every content type', () => {
    const bytes = buildSyntheticPackageBytes();
    const pkg = OpcPackage.load(bytes);
    expect(pkg.parts.map((p) => p.name)).toEqual(SYNTHETIC_PACKAGE_PART_NAMES);
    for (const part of pkg.parts) {
      expect(part.contentType).not.toBe('');
    }
  });

  it('exposes the part content type for an override entry', () => {
    const pkg = OpcPackage.load(buildSyntheticPackageBytes());
    const slide = pkg.getPart(partName('/ppt/slides/slide1.xml'));
    expect(slide?.contentType).toBe(
      'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'.replace(
        '.presentation.main+xml',
        '.slide+xml',
      ),
    );
  });

  it('exposes the part content type for a default entry (extension lookup)', () => {
    const pkg = OpcPackage.load(buildSyntheticPackageBytes());
    const image = pkg.getPart(partName('/ppt/media/image1.png'));
    expect(image?.contentType).toBe('image/png');
  });

  it('resolves case-insensitive Overrides, first match winning', () => {
    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="XML" ContentType="application/xml"/>
  <Override PartName="/PPT/A.xml" ContentType="first"/>
  <Override PartName="/ppt/a.xml" ContentType="second"/>
</Types>`;
    const pkg = OpcPackage.load(
      writeZip([
        { name: '[Content_Types].xml', data: new TextEncoder().encode(contentTypes) },
        { name: 'ppt/a.xml', data: new Uint8Array() },
        { name: 'ppt/b.Xml', data: new Uint8Array() },
      ]),
    );
    expect(pkg.getPart(partName('/ppt/a.xml'))?.contentType).toBe('first');
    expect(pkg.getPart(partName('/PPT/B.XML'))?.contentType).toBe('application/xml');
  });

  it('loads a package with tens of thousands of Overrides in linear time', () => {
    const build = (count: number): Uint8Array => {
      const overrides: string[] = [];
      const entries = [];
      for (let i = 0; i < count; i++) {
        overrides.push(`<Override PartName="/p/part${i}.bin" ContentType="application/x-test"/>`);
        entries.push({ name: `p/part${i}.bin`, data: new Uint8Array() });
      }
      const xml = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${overrides.join('')}</Types>`;
      return writeZip([
        { name: '[Content_Types].xml', data: new TextEncoder().encode(xml) },
        ...entries,
      ]);
    };
    const time = (bytes: Uint8Array): number => {
      const started = performance.now();
      const pkg = OpcPackage.load(bytes);
      const elapsed = performance.now() - started;
      expect(pkg.getPart(partName('/P/PART7.BIN'))?.contentType).toBe('application/x-test');
      return elapsed;
    };
    const small = build(3_000);
    const large = build(30_000);
    time(small);
    // Ten times the parts must not cost anywhere near the hundredfold a
    // per-part scan of every Override would.
    expect(time(large)).toBeLessThan(Math.max(time(small), 5) * 30);
  });

  it('throws when [Content_Types].xml is missing', () => {
    // Build a package without [Content_Types].xml by writing a raw ZIP.
    expect(() => OpcPackage.load(new Uint8Array([0]))).toThrow();
  });
});

describe('OpcPackage round-trip', () => {
  it('preserves entry order and binary media bytes', () => {
    const bytes = buildSyntheticPackageBytes();
    const pkg = OpcPackage.load(bytes);
    const written = pkg.save();
    const reread = OpcPackage.load(written);

    expect(reread.parts.map((p) => p.name)).toEqual(SYNTHETIC_PACKAGE_PART_NAMES);

    const image = reread.getPart(partName('/ppt/media/image1.png'));
    expect(image?.data).toEqual(ONE_PIXEL_PNG_BYTES);

    const slide = reread.getPart(partName('/ppt/slides/slide1.xml'));
    expect(decode(slide?.data ?? new Uint8Array())).toContain('<p:sld');
  });

  it('round-trips XML parts byte-equivalent through the parsers', () => {
    const bytes = buildSyntheticPackageBytes();
    const pkg1 = OpcPackage.load(bytes);
    const written = pkg1.save();
    const pkg2 = OpcPackage.load(written);
    for (let i = 0; i < pkg1.parts.length; i++) {
      const a = pkg1.parts[i];
      const b = pkg2.parts[i];
      expect(b?.name).toBe(a?.name);
      expect(b?.contentType).toBe(a?.contentType);
      expect(b?.data).toEqual(a?.data);
    }
  });

  it('exposes parsed relationships per part', () => {
    const pkg = OpcPackage.load(buildSyntheticPackageBytes());

    const root = pkg.rootRels();
    expect(root?.items[0]?.type).toBe(SYNTHETIC_REL_TYPES.officeDoc);

    const presRels = pkg.getRels(partName('/ppt/presentation.xml'));
    expect(presRels?.items[0]?.type).toBe(SYNTHETIC_REL_TYPES.slide);

    const slideRels = pkg.getRels(partName('/ppt/slides/slide1.xml'));
    expect(slideRels?.items.length).toBe(2);
    expect(slideRels?.items[0]?.targetMode).toBe('External');
    expect(slideRels?.items[1]?.target).toBe('../media/image1.png');
  });

  it('returns null rels for parts without a sibling .rels', () => {
    const pkg = OpcPackage.load(buildSyntheticPackageBytes());
    expect(pkg.getRels(partName('/ppt/media/image1.png'))).toBeNull();
  });

  it('persists rels changes through save/reload', () => {
    const pkg = OpcPackage.load(buildSyntheticPackageBytes());
    const slideRels = pkg.getRels(partName('/ppt/slides/slide1.xml'));
    if (!slideRels) throw new Error('expected slide rels');
    slideRels.items.push({
      id: 'rId3',
      type: 'http://example.com/test-rel-type',
      target: 'some-target',
      targetMode: 'Internal',
    });
    pkg.setRels(partName('/ppt/slides/slide1.xml'), slideRels);

    const reread = OpcPackage.load(pkg.save());
    const out = reread.getRels(partName('/ppt/slides/slide1.xml'));
    expect(out?.items.map((i) => i.id)).toEqual(['rId1', 'rId2', 'rId3']);
  });
});

describe('OpcPackage authoring', () => {
  it('adds a new part with an Override content type', () => {
    const pkg = OpcPackage.empty();
    const added = pkg.addPart(
      partName('/docProps/core.xml'),
      'application/vnd.openxmlformats-package.core-properties+xml',
      new TextEncoder().encode('<cp:coreProperties xmlns:cp="urn:cp"/>'),
    );
    expect(added.contentType).toBe('application/vnd.openxmlformats-package.core-properties+xml');
    expect(pkg.contentTypes.overrides.length).toBe(1);
  });

  it('refuses to add a duplicate part name', () => {
    const pkg = OpcPackage.empty();
    pkg.addPart(partName('/a.xml'), 'application/xml', new Uint8Array());
    expect(() => pkg.addPart(partName('/A.XML'), 'application/xml', new Uint8Array())).toThrow();
  });

  it('removes a part and its Override entry', () => {
    const pkg = OpcPackage.empty();
    pkg.addPart(
      partName('/ppt/slides/slide99.xml'),
      'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
      new TextEncoder().encode('<p:sld xmlns:p="urn:p"/>'),
    );
    expect(pkg.removePart(partName('/ppt/slides/slide99.xml'))).toBe(true);
    expect(pkg.parts.length).toBe(0);
    expect(pkg.contentTypes.overrides.length).toBe(0);
    expect(pkg.getPart(partName('/ppt/slides/slide99.xml'))).toBeNull();
  });

  it('finds added parts case-insensitively and forgets removed ones', () => {
    const pkg = OpcPackage.empty();
    const added = pkg.addPart(partName('/ppt/A.xml'), 'application/xml', new Uint8Array());
    expect(pkg.getPart(partName('/PPT/a.XML'))).toBe(added);
    expect(pkg.removePart(partName('/ppt/a.xml'))).toBe(true);
    expect(pkg.getPart(partName('/ppt/A.xml'))).toBeNull();
    const readded = pkg.addPart(partName('/ppt/a.xml'), 'application/xml', new Uint8Array());
    expect(pkg.getPart(partName('/ppt/A.xml'))).toBe(readded);
  });

  it('omits the Override when an existing default already matches', () => {
    const pkg = OpcPackage.empty();
    // The defaults already register `xml` → application/xml.
    pkg.addPart(partName('/random.xml'), 'application/xml', new TextEncoder().encode('<x/>'));
    expect(pkg.contentTypes.overrides.length).toBe(0);
  });
});

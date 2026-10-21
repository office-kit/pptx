// compactPackage — drop parts unreachable from the package root.

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  _internalPackageOf,
  addSlide,
  addSlideChart,
  addSlideImage,
  compactPackage,
  getMediaParts,
  getOrphanMediaPartNames,
  getSlideLayouts,
  getSlideNotes,
  getSlidePartName,
  getSlides,
  inches,
  listPackageParts,
  loadPresentation,
  type PresentationData,
  removeShape,
  removeSlide,
  savePresentation,
  setSlideNotes,
  type SlideData,
} from '../src/api/index.ts';
import { partName, relsPartNameFor, resolveTarget } from '../src/internal/opc/index.ts';
import { expectPackageValid } from './lib/expect-package-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const tinyPng = (): Uint8Array =>
  new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
    0x42, 0x60, 0x82,
  ]);

describe('fn API: compactPackage', () => {
  it('removes media parts not referenced by any rel', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const pic = addSlideImage(slide, tinyPng(), {
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      format: 'png',
    });
    // Save → reload → orphan the media by clearing slide rels.
    const reloaded = await loadPresentation(await savePresentation(pres));
    const pkg = _internalPackageOf(reloaded);
    pkg.setRels(partName('/ppt/slides/slide1.xml'), {
      items: [
        {
          id: 'rId1',
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
          target: '../slideLayouts/slideLayout1.xml',
          targetMode: 'Internal',
        },
      ],
    });

    const after = await loadPresentation(await savePresentation(reloaded));
    expect(getMediaParts(after).length).toBeGreaterThan(0);
    const { removed } = compactPackage(after);
    expect(removed.length).toBeGreaterThan(0);
    expect(getMediaParts(after).length).toBe(0);
    void pic;
  });

  it('does not touch media that is still referenced', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    addSlideImage(getSlides(pres)[0]!, tinyPng(), {
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      format: 'png',
    });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const before = getMediaParts(reloaded).length;
    expect(before).toBeGreaterThan(0);
    expect(compactPackage(reloaded).removed).toEqual([]);
    expect(getMediaParts(reloaded).length).toBe(before);
  });

  it('keeps media whose rel target case differs from the stored part name', async () => {
    // OPC part names compare case-insensitively (ECMA-376 Part 2 §9.1.1), so a
    // rel pointing at `../media/IMAGE1.PNG` still references the stored
    // `/ppt/media/image1.png`. A case-sensitive compaction would wrongly drop it.
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    addSlideImage(getSlides(pres)[0]!, tinyPng(), {
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      format: 'png',
    });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const pkg = _internalPackageOf(reloaded);
    const slideRels = pkg.getRels(partName('/ppt/slides/slide1.xml'))!;
    // Uppercase only the media rel's target; leave layout/other rels untouched.
    pkg.setRels(partName('/ppt/slides/slide1.xml'), {
      items: slideRels.items.map((r) =>
        r.target.includes('/media/') ? { ...r, target: r.target.toUpperCase() } : r,
      ),
    });
    const before = getMediaParts(reloaded).length;
    expect(before).toBeGreaterThan(0);

    expect(getOrphanMediaPartNames(reloaded)).toEqual([]);
    expect(compactPackage(reloaded).removed).toEqual([]);
    expect(getMediaParts(reloaded).length).toBe(before);
  });

  it('cleans up after a removeShape() that took the last image with it', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const pic = addSlideImage(getSlides(pres)[0]!, tinyPng(), {
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      format: 'png',
    });
    removeShape(pic);
    // The media is still in the package but no shape uses it now.
    // The rel is still there too (removeShape doesn't touch slide rels in v1).
    // compactPackage shouldn't drop the media because the rel still
    // references it; this is the documented behavior.
    expect(compactPackage(pres).removed).toEqual([]);
  });

  describe('unreachable parts', () => {
    const TAGS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags';
    const LAYOUT_REL =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';

    const addTags = (pres: PresentationData, slide: SlideData, n: number): void => {
      const pkg = _internalPackageOf(pres);
      pkg.addPart(
        partName(`/ppt/tags/tag${n}.xml`),
        'application/vnd.openxmlformats-officedocument.presentationml.tags+xml',
        new TextEncoder().encode(
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
            `<p:tag name="SOURCE" val="${n}"/></p:tagLst>`,
        ),
      );
      const source = partName(getSlidePartName(slide));
      const rels = pkg.getRels(source)!;
      rels.items.push({
        id: `rIdTags${n}`,
        type: TAGS_REL,
        target: `../tags/tag${n}.xml`,
        targetMode: 'Internal',
      });
      pkg.setRels(source, rels);
    };

    // Internal targets of `source`'s rels other than its layout.
    const targetsOf = (pres: PresentationData, source: string): string[] =>
      (_internalPackageOf(pres).getRels(partName(source))?.items ?? [])
        .filter((r) => r.targetMode !== 'External' && r.type !== LAYOUT_REL)
        .map((r) => resolveTarget(partName(source), r.target));

    // What a slide alone pulls in: its charts with their workbooks, its notes
    // slide, tags and pictures. Notes slides point back at the slide and on to
    // the shared notes master, so they are not followed.
    const slideDependencies = (pres: PresentationData, slide: SlideData): Set<string> => {
      const out = new Set<string>();
      for (const target of targetsOf(pres, getSlidePartName(slide))) {
        out.add(target);
        if (target.startsWith('/ppt/charts/')) for (const t of targetsOf(pres, target)) out.add(t);
      }
      return out;
    };

    const partNames = (pres: PresentationData): string[] =>
      listPackageParts(pres).map((p) => p.name);

    it('drops the charts, workbooks, notes slide and tags of a removed slide', async () => {
      // A pptxgenjs deck: a chart slide with its notes slide and a notes master.
      const files = unzipSync(
        await readFile(fileURLToPath(new URL('./fixtures/pptxgenjs/combo.pptx', import.meta.url))),
      );
      // pptxgenjs writes <p:notesMasterIdLst> after <p:sldIdLst>; the schema wants it before.
      const presentation = strFromU8(files['ppt/presentation.xml']!);
      const notesMasterIdLst = /<p:notesMasterIdLst>.*?<\/p:notesMasterIdLst>/.exec(
        presentation,
      )![0];
      files['ppt/presentation.xml'] = strToU8(
        presentation
          .replace(notesMasterIdLst, '')
          .replace('<p:sldIdLst>', `${notesMasterIdLst}<p:sldIdLst>`),
      );
      const pres = await loadPresentation(zipSync(files));
      addSlide(pres, { layout: getSlideLayouts(pres)[0]! });
      const slides = getSlides(pres);
      for (const [i, slide] of slides.entries()) {
        addSlideChart(slide, {
          x: inches(0.5),
          y: inches(0.5),
          w: inches(4),
          h: inches(3),
          spec: {
            kind: 'column',
            categories: ['Q1', 'Q2'],
            series: [{ name: `Series ${i + 1}`, values: [1, 2] }],
          },
        });
        addSlideImage(slide, tinyPng(), {
          x: inches(5),
          y: inches(0),
          w: inches(1),
          h: inches(1),
          format: 'png',
        });
        setSlideNotes(slide, `Notes ${i + 1}`);
        addTags(pres, slide, i + 1);
      }
      const [first, second] = slides as [SlideData, SlideData];
      const kept = slideDependencies(pres, second);
      const dropped = [...slideDependencies(pres, first)].filter((name) => !kept.has(name));
      // Each kind the removed slide used must be represented, or the test proves nothing.
      for (const prefix of ['/ppt/charts/', '/ppt/embeddings/', '/ppt/notesSlides/', '/ppt/tags/'])
        expect(dropped.some((name) => name.startsWith(prefix))).toBe(true);
      const droppedWithRels = new Set(
        dropped.flatMap((name) => [name, relsPartNameFor(partName(name))]),
      );

      const slideName = getSlidePartName(first);
      removeSlide(pres, first);
      const before = partNames(pres);
      const { removed } = compactPackage(pres);

      // Exactly the removed slide's own parts (and their rels) go; presentation.xml,
      // docProps, theme, master, every layout, notes master and the other slide stay.
      expect(new Set(removed)).toEqual(new Set(before.filter((name) => droppedWithRels.has(name))));
      expect(partNames(pres)).toEqual(before.filter((name) => !droppedWithRels.has(name)));
      for (const name of [
        '/ppt/presentation.xml',
        '/docProps/core.xml',
        '/docProps/app.xml',
        '/ppt/presProps.xml',
        '/ppt/theme/theme1.xml',
        '/ppt/slideMasters/slideMaster1.xml',
        '/ppt/slideLayouts/slideLayout1.xml',
        '/ppt/notesMasters/notesMaster1.xml',
      ])
        expect(partNames(pres)).toContain(name);
      for (const name of kept) expect(partNames(pres)).toContain(name);
      expect(partNames(pres)).not.toContain(slideName);
      const overrides = _internalPackageOf(pres).contentTypes.overrides.map((o) => o.partName);
      for (const name of removed) expect(overrides).not.toContain(name);

      const reloaded = await expectPackageValid(pres);
      expect(getSlides(reloaded)).toHaveLength(1);
      expect(getSlideNotes(getSlides(reloaded)[0]!)).toBe('Notes 2');
      expect(compactPackage(reloaded).removed).toEqual([]);
    });

    it('keeps everything in an untouched deck', async () => {
      const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
      const before = partNames(pres);
      expect(compactPackage(pres).removed).toEqual([]);
      expect(partNames(pres)).toEqual(before);
    });

    it('follows a percent-encoded target to a part stored under the decoded name', async () => {
      const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
      const pkg = _internalPackageOf(pres);
      const picture = partName('/ppt/media/logo 1.png');
      pkg.addPart(picture, 'image/png', tinyPng());
      const slide = partName('/ppt/slides/slide1.xml');
      const rels = pkg.getRels(slide)!;
      rels.items.push({
        id: 'rIdLogo',
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        target: '../media/logo%201.png',
        targetMode: 'Internal',
      });
      pkg.setRels(slide, rels);

      expect(compactPackage(pres).removed).toEqual([]);
      expect(partNames(pres)).toContain(picture);
    });

    it('drops a part nothing points at, with its override', async () => {
      const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
      const pkg = _internalPackageOf(pres);
      const stray = partName('/ppt/tags/tag9.xml');
      pkg.addPart(
        stray,
        'application/vnd.openxmlformats-officedocument.presentationml.tags+xml',
        new TextEncoder().encode('<p:tagLst/>'),
      );
      expect(pkg.contentTypes.overrides.map((o) => o.partName)).toContain(stray);

      expect(compactPackage(pres).removed).toEqual([stray]);
      expect(pkg.contentTypes.overrides.map((o) => o.partName)).not.toContain(stray);
      await expectPackageValid(pres);
    });

    it('keeps an override whose part was already missing', async () => {
      const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
      const pkg = _internalPackageOf(pres);
      const dangling = partName('/ppt/tags/tag9.xml');
      pkg.contentTypes.overrides.push({
        partName: dangling,
        contentType: 'application/vnd.openxmlformats-officedocument.presentationml.tags+xml',
      });

      expect(compactPackage(pres).removed).toEqual([]);
      expect(pkg.contentTypes.overrides.map((o) => o.partName)).toContain(dangling);
    });

    it('reports media referenced only by an unreachable part as orphaned', async () => {
      const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
      const pkg = _internalPackageOf(pres);
      const chart = partName('/ppt/charts/chart99.xml');
      const picture = partName('/ppt/media/image99.png');
      pkg.addPart(
        chart,
        'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
        new TextEncoder().encode('<c:chartSpace/>'),
      );
      pkg.addPart(picture, 'image/png', tinyPng());
      pkg.setRels(chart, {
        items: [
          {
            id: 'rId1',
            type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            target: '../media/image99.png',
            targetMode: 'Internal',
          },
        ],
      });

      expect(getOrphanMediaPartNames(pres)).toEqual([picture]);
      expect(new Set(compactPackage(pres).removed)).toEqual(
        new Set([chart, relsPartNameFor(chart), picture]),
      );
    });

    describe('without /_rels/.rels', () => {
      const withoutRootRels = async (): Promise<PresentationData> => {
        const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
        _internalPackageOf(pres).removePart(partName('/_rels/.rels'));
        return pres;
      };

      it('compactPackage throws and leaves the package unchanged', async () => {
        const pres = await withoutRootRels();
        const before = partNames(pres);
        const overrides = [..._internalPackageOf(pres).contentTypes.overrides];
        expect(() => compactPackage(pres)).toThrow('package has no /_rels/.rels');
        expect(partNames(pres)).toEqual(before);
        expect(_internalPackageOf(pres).contentTypes.overrides).toEqual(overrides);
      });

      it('getOrphanMediaPartNames reports every media part', async () => {
        const pres = await withoutRootRels();
        addSlideImage(getSlides(pres)[0]!, tinyPng(), {
          x: inches(0),
          y: inches(0),
          w: inches(1),
          h: inches(1),
          format: 'png',
        });
        const media = getMediaParts(pres).map((p) => p.name);
        expect(media.length).toBeGreaterThan(0);
        expect(getOrphanMediaPartNames(pres)).toEqual(media);
      });
    });
  });
});

// compactPackage — drop parts unreachable from the package root.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  _internalPackageOf,
  addSlideChart,
  addSlideImage,
  compactPackage,
  getMediaParts,
  getOrphanMediaPartNames,
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
  setNotesMasterPlaceholderIncluded,
  setSlideNotes,
  type SlideData,
} from '../src/api/index.ts';
import { partName, relsPartNameFor } from '../src/internal/opc/index.ts';
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

  it('stops resolving removed parts by name', async () => {
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

    const { removed } = compactPackage(reloaded);
    expect(removed.length).toBeGreaterThan(0);
    for (const name of removed) expect(pkg.getPart(partName(name))).toBeNull();
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

    const addImageRel = (pres: PresentationData, target: string): void => {
      const pkg = _internalPackageOf(pres);
      const slide = partName('/ppt/slides/slide1.xml');
      const rels = pkg.getRels(slide)!;
      rels.items.push({
        id: 'rIdExtra',
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        target,
        targetMode: 'Internal',
      });
      pkg.setRels(slide, rels);
    };

    const partNames = (pres: PresentationData): string[] =>
      listPackageParts(pres).map((p) => p.name);

    it('drops the charts, workbooks, notes slide and tags of a removed slide', async () => {
      const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
      for (const [i, slide] of getSlides(pres).entries()) {
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
      // Creates the notes master and links both notes slides to it.
      setNotesMasterPlaceholderIncluded(pres, 'hdr', true);

      const first = getSlides(pres)[0]!;
      removeSlide(pres, first);
      const before = partNames(pres);
      const { removed } = compactPackage(pres);

      expect([...removed].sort()).toEqual(
        [
          '/ppt/charts/chart1.xml',
          '/ppt/charts/_rels/chart1.xml.rels',
          '/ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx',
          '/ppt/media/image1.png',
          '/ppt/notesSlides/notesSlide1.xml',
          '/ppt/notesSlides/_rels/notesSlide1.xml.rels',
          '/ppt/tags/tag1.xml',
        ].sort(),
      );
      expect(partNames(pres)).toEqual(before.filter((name) => !removed.includes(name)));
      // presentation.xml, docProps, theme, master, every layout, the notes
      // master with its theme and the other slide's parts stay.
      for (const name of [
        '/ppt/presentation.xml',
        '/docProps/core.xml',
        '/docProps/app.xml',
        '/ppt/presProps.xml',
        '/ppt/theme/theme1.xml',
        '/ppt/slideMasters/slideMaster1.xml',
        '/ppt/slideLayouts/slideLayout1.xml',
        '/ppt/slideLayouts/slideLayout11.xml',
        '/ppt/notesMasters/notesMaster1.xml',
        '/ppt/theme/theme2.xml',
        '/ppt/charts/chart2.xml',
        '/ppt/embeddings/Microsoft_Excel_Worksheet2.xlsx',
        '/ppt/media/image2.png',
        '/ppt/notesSlides/notesSlide2.xml',
        '/ppt/tags/tag2.xml',
      ])
        expect(partNames(pres)).toContain(name);
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
      const picture = partName('/ppt/media/logo 1.png');
      _internalPackageOf(pres).addPart(picture, 'image/png', tinyPng());
      addImageRel(pres, '../media/logo%201.png');

      expect(compactPackage(pres).removed).toEqual([]);
      expect(partNames(pres)).toContain(picture);
    });

    it('follows a percent-encoded non-ASCII target whatever its case', async () => {
      const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
      const pkg = _internalPackageOf(pres);
      const picture = partName('/ppt/media/Émage.png');
      pkg.addPart(picture, 'image/png', tinyPng());
      addImageRel(pres, '../media/%C3%89mage.png');

      expect(getOrphanMediaPartNames(pres)).toEqual([]);
      expect(compactPackage(pres).removed).toEqual([]);
    });

    it('drops `..` segments above the root instead of losing the part', async () => {
      const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
      const picture = partName('/ppt/media/image99.png');
      _internalPackageOf(pres).addPart(picture, 'image/png', tinyPng());
      addImageRel(pres, '../../../ppt/media/image99.png');

      expect(getOrphanMediaPartNames(pres)).toEqual([]);
      expect(compactPackage(pres).removed).toEqual([]);
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

    describe('when /_rels/.rels does not reach presentation.xml', () => {
      const unreachablePresentation = async (): Promise<PresentationData> => {
        const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
        addSlideImage(getSlides(pres)[0]!, tinyPng(), {
          x: inches(0),
          y: inches(0),
          w: inches(1),
          h: inches(1),
          format: 'png',
        });
        const rootRels = _internalPackageOf(pres).getPart(partName('/_rels/.rels'))!;
        const xml = new TextDecoder().decode(rootRels.data);
        const withoutPresentation = xml.replace(/<Relationship [^>]*\/officeDocument"[^>]*\/>/, '');
        expect(withoutPresentation).not.toBe(xml);
        rootRels.data = new TextEncoder().encode(withoutPresentation);
        return pres;
      };

      it('compactPackage removes nothing', async () => {
        const pres = await unreachablePresentation();
        const before = partNames(pres);
        expect(compactPackage(pres).removed).toEqual([]);
        expect(partNames(pres)).toEqual(before);
      });

      it('getOrphanMediaPartNames reports nothing', async () => {
        expect(getOrphanMediaPartNames(await unreachablePresentation())).toEqual([]);
      });
    });
  });
});

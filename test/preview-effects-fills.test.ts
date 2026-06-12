// Renderer coverage for W5 effects & fills polish:
//   1. Reflection (`a:reflection`)        — flipped, gradient-masked copy.
//   2. Picture bullets (`a:buBlip`)       — inline <image> / <img>, not "■".
//   3. Inherited gradient fills           — real gradient def from the
//      layout/master cascade, never the old #FDBA74 orange tint.
//
// None of these has a public authoring API, so the OOXML is injected at the
// OPC zip layer — the same hook test/preview-custom-geometry.test.ts uses.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
// Import authoring from `pptx-kit` so the deck's `PresentationData` shares
// type identity with `renderSlideToSvg` (both import from the package).
import {
  addSlide,
  addSlideShape,
  addSlideTextBox,
  findSlideLayout,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeFill,
} from 'pptx-kit';
import { type ZipEntry, readZip, writeZip } from '../src/internal/opc/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

const dec = new TextDecoder();
const enc = new TextEncoder();

// 1×1 red PNG — enough for the image-bullet path to resolve real bytes.
const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  ),
  (c) => c.charCodeAt(0),
);

const editEntry = (entries: readonly ZipEntry[], name: string, edit: (xml: string) => string) =>
  entries.map((e) =>
    e.name === name ? { name: e.name, data: enc.encode(edit(dec.decode(e.data))) } : e,
  );

const slideEntryName = (entries: readonly ZipEntry[]): string => {
  const e = entries.find((x) => /slides\/slide\d+\.xml$/.test(x.name));
  if (!e) throw new Error('no slide part');
  return e.name;
};

describe('renderSlideToSvg: reflection effect', () => {
  // A standard PowerPoint "tight reflection": 50% near alpha, ~0% far alpha,
  // a full-height mirror (sy = -100%), pushed down by `dist`.
  const REFLECTION =
    '<a:effectLst><a:reflection blurRad="6350" stA="50000" stPos="0" endA="300" endPos="55000" dist="50800" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/></a:effectLst>';

  const renderWithReflection = async (): Promise<string> => {
    const pres = await loadPresentation(await readFile(fixturePath));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('Blank layout missing');
    const slide = addSlide(pres, { layout });
    const rect = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
    });
    setShapeFill(rect, '#3366CC');
    const { entries } = readZip(await savePresentation(pres));
    const name = slideEntryName(entries);
    const modified = editEntry(entries, name, (xml) =>
      xml.replace(
        '<a:solidFill><a:srgbClr val="3366CC"/></a:solidFill></p:spPr>',
        `<a:solidFill><a:srgbClr val="3366CC"/></a:solidFill>${REFLECTION}</p:spPr>`,
      ),
    );
    const reloaded = await loadPresentation(writeZip(modified));
    return renderSlideToSvg(reloaded, getSlides(reloaded).at(-1)!, { textLayout: 'svg' });
  };

  it('emits a vertically flipped, masked copy of the shape geometry', async () => {
    const svg = await renderWithReflection();
    expect(svg).toContain('data-pptx-reflection="1"');
    // sy = -100000 → full-height mirror about the bottom edge.
    expect(svg).toContain('scale(1 -1)');
    expect(svg).toContain('mask=');
    // The reflected copy carries the shape's own fill so it actually mirrors.
    expect(svg.toLowerCase()).toContain('#3366cc');
  });

  it('fades from the near-edge alpha (stA) to the far-edge alpha (endA)', async () => {
    const svg = await renderWithReflection();
    const grad = svg.match(/<linearGradient[\s\S]*?<\/linearGradient>/)?.[0] ?? '';
    // stA 50000 → 0.5 at the contact edge; endA 300 → ~0 at the far edge.
    expect(grad).toContain('stop-opacity="0.500"');
    expect(grad).toContain('stop-opacity="0.003"');
    // The contact-edge stop is the more opaque one (offset 0).
    expect(grad).toMatch(/offset="0"[^>]*stop-opacity="0\.500"/);
  });
});

describe('renderSlideToSvg: picture bullets', () => {
  // Builds a one-textbox deck, rewrites the paragraph to carry a `<a:buBlip>`
  // image bullet, and (when `withBytes`) wires up the media part + rel so the
  // bytes actually resolve. Returns both text-layout modes' SVG.
  const renderWithPictureBullet = async (
    withBytes: boolean,
  ): Promise<{ svg: string; fo: string }> => {
    const pres = await loadPresentation(await readFile(fixturePath));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('Blank layout missing');
    const slide = addSlide(pres, { layout });
    addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'Item',
    });
    const { entries } = readZip(await savePresentation(pres));
    const slideName = slideEntryName(entries);
    const relName = `ppt/slides/_rels/${slideName.split('/').pop()}.rels`;

    let out: ZipEntry[] = editEntry(entries, slideName, (xml) =>
      xml.replace(
        '<a:p><a:r>',
        '<a:p><a:pPr><a:buBlip><a:blip r:embed="rIdBullet"/></a:buBlip></a:pPr><a:r>',
      ),
    );
    if (withBytes) {
      out = editEntry(out, relName, (xml) =>
        xml.replace(
          '</Relationships>',
          '<Relationship Id="rIdBullet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/bullet.png"/></Relationships>',
        ),
      );
      out = editEntry(out, '[Content_Types].xml', (xml) =>
        xml.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>'),
      );
      out.push({ name: 'ppt/media/bullet.png', data: PNG });
    }
    const reloaded = await loadPresentation(writeZip(out));
    const s = getSlides(reloaded).at(-1)!;
    return {
      svg: renderSlideToSvg(reloaded, s, { textLayout: 'svg' }),
      fo: renderSlideToSvg(reloaded, s, { textLayout: 'foreignObject' }),
    };
  };

  it('renders the bullet as an inline <image> in svg mode (no "■")', async () => {
    const { svg } = await renderWithPictureBullet(true);
    expect(svg).toContain('<image');
    expect(svg).toContain('data:image/png;base64,');
    expect(svg).not.toContain('■');
  });

  it('renders the bullet as an <img> in foreignObject mode (no "■")', async () => {
    const { fo } = await renderWithPictureBullet(true);
    expect(fo).toContain('<img');
    expect(fo).toContain('data:image/png;base64,');
    expect(fo).not.toContain('■');
  });

  it('keeps the "■" fallback when the bullet bytes are unavailable', async () => {
    const { svg, fo } = await renderWithPictureBullet(false);
    expect(svg).toContain('■');
    expect(svg).not.toContain('data:image/png;base64,');
    expect(fo).toContain('■');
  });
});

describe('renderSlideToSvg: inherited gradient fills', () => {
  // Adds a gradient to the layout's ctrTitle placeholder, then renders a
  // slide whose ctrTitle inherits it (the slide shape carries no own fill).
  const GRAD =
    '<a:gradFill><a:gsLst>' +
    '<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>' +
    '<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>' +
    '</a:gsLst><a:lin ang="5400000"/></a:gradFill>';

  const renderWithInheritedGradient = async (): Promise<string> => {
    const pres = await loadPresentation(await readFile(fixturePath));
    const layout = findSlideLayout(pres, 'Title Slide');
    if (!layout) throw new Error('Title Slide layout missing');
    addSlide(pres, { layout });
    const { entries } = readZip(await savePresentation(pres));
    // slideLayout1.xml is the "Title Slide" layout; its ctrTitle placeholder
    // spPr ends with the xfrm below — inject the gradient right after it.
    const modified = editEntry(entries, 'ppt/slideLayouts/slideLayout1.xml', (xml) =>
      xml.replace(
        '<a:ext cx="7772400" cy="1470025"/></a:xfrm></p:spPr>',
        `<a:ext cx="7772400" cy="1470025"/></a:xfrm>${GRAD}</p:spPr>`,
      ),
    );
    const reloaded = await loadPresentation(writeZip(modified));
    return renderSlideToSvg(reloaded, getSlides(reloaded).at(-1)!, { textLayout: 'svg' });
  };

  it('emits a real gradient def resolved from the layout cascade', async () => {
    const svg = await renderWithInheritedGradient();
    expect(svg).toContain('<linearGradient');
    expect(svg.toUpperCase()).toContain('FF0000');
    expect(svg.toUpperCase()).toContain('0000FF');
  });

  it('never paints the old #FDBA74 orange-tint fallback', async () => {
    const svg = await renderWithInheritedGradient();
    expect(svg.toUpperCase()).not.toContain('FDBA74');
  });

  it('the #FDBA74 fallback constant is gone from the renderer source', async () => {
    const src = await readFile(
      fileURLToPath(new URL('../packages/preview/src/render-slide.ts', import.meta.url)),
      'utf8',
    );
    expect(src).not.toContain('FDBA74');
  });
});

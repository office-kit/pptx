// Builds a *syntactically* valid OPC package programmatically.
//
// The output is NOT a valid PPTX — there's no real PresentationML — but it
// has the structural shape that exercises every code path in the OPC layer:
//
//   - [Content_Types].xml with both Default and Override entries.
//   - Multiple parts at varying depths in the URI tree.
//   - A package-root rels file (`/_rels/.rels`) and a per-part rels file
//     (`/ppt/_rels/presentation.xml.rels`).
//   - At least one external relationship and one internal relationship.
//   - A binary media part to exercise the STORE compression path.
//
// Use from tests that want to confirm OPC behavior without depending on a
// real PPTX fixture. Real-PPTX fixtures live under `test/fixtures/`.

import {
  type ContentTypes,
  type Relationships,
  emptyContentTypes,
  partName,
  serializeContentTypes,
  serializeRels,
  toZipPath,
  writeZip,
} from '../../src/internal/opc/index.ts';

const encode = (s: string): Uint8Array => new TextEncoder().encode(s);

const REL_TYPE_OFFICE_DOC =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const REL_TYPE_SLIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
const REL_TYPE_HYPERLINK =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const REL_TYPE_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

const CT_PRESENTATION =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml';
const CT_SLIDE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const CT_PNG = 'image/png';

const placeholderPresentation =
  '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>';
const placeholderSlide =
  '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sld>';

// A 1x1 transparent PNG (real bytes, decodable by image software). Lifted by
// hand from the spec; checked in here to avoid runtime image generation.
// prettier-ignore
const ONE_PIXEL_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

const makeContentTypes = (): ContentTypes => {
  const ct = emptyContentTypes();
  // Default for PNG so the media part doesn't need an explicit Override.
  ct.defaults.push({ extension: 'png', contentType: CT_PNG });
  ct.overrides.push(
    { partName: partName('/ppt/presentation.xml'), contentType: CT_PRESENTATION },
    { partName: partName('/ppt/slides/slide1.xml'), contentType: CT_SLIDE },
  );
  return ct;
};

const makeRootRels = (): Relationships => ({
  items: [
    {
      id: 'rId1',
      type: REL_TYPE_OFFICE_DOC,
      target: 'ppt/presentation.xml',
      targetMode: 'Internal',
    },
  ],
});

const makePresRels = (): Relationships => ({
  items: [
    { id: 'rId1', type: REL_TYPE_SLIDE, target: 'slides/slide1.xml', targetMode: 'Internal' },
  ],
});

const makeSlideRels = (): Relationships => ({
  items: [
    {
      id: 'rId1',
      type: REL_TYPE_HYPERLINK,
      target: 'https://example.com',
      targetMode: 'External',
    },
    {
      id: 'rId2',
      type: REL_TYPE_IMAGE,
      target: '../media/image1.png',
      targetMode: 'Internal',
    },
  ],
});

/**
 * Returns the byte buffer of a synthetic OPC package. The entry list, in
 * write order:
 *
 *   1. [Content_Types].xml
 *   2. _rels/.rels
 *   3. ppt/presentation.xml
 *   4. ppt/_rels/presentation.xml.rels
 *   5. ppt/slides/slide1.xml
 *   6. ppt/slides/_rels/slide1.xml.rels
 *   7. ppt/media/image1.png
 */
export const buildSyntheticPackageBytes = (): Uint8Array => {
  const entries = [
    {
      name: '[Content_Types].xml',
      data: encode(serializeContentTypes(makeContentTypes())),
    },
    { name: toZipPath(partName('/_rels/.rels')), data: encode(serializeRels(makeRootRels())) },
    { name: toZipPath(partName('/ppt/presentation.xml')), data: encode(placeholderPresentation) },
    {
      name: toZipPath(partName('/ppt/_rels/presentation.xml.rels')),
      data: encode(serializeRels(makePresRels())),
    },
    { name: toZipPath(partName('/ppt/slides/slide1.xml')), data: encode(placeholderSlide) },
    {
      name: toZipPath(partName('/ppt/slides/_rels/slide1.xml.rels')),
      data: encode(serializeRels(makeSlideRels())),
    },
    { name: toZipPath(partName('/ppt/media/image1.png')), data: ONE_PIXEL_PNG },
  ];
  return writeZip(entries);
};

/** The expected part names in the synthetic package, in load order. */
export const SYNTHETIC_PACKAGE_PART_NAMES: ReadonlyArray<string> = [
  '/_rels/.rels',
  '/ppt/presentation.xml',
  '/ppt/_rels/presentation.xml.rels',
  '/ppt/slides/slide1.xml',
  '/ppt/slides/_rels/slide1.xml.rels',
  '/ppt/media/image1.png',
];

export const ONE_PIXEL_PNG_BYTES: Readonly<Uint8Array> = ONE_PIXEL_PNG;

// Re-export the placeholder XML so tests can assert exact byte equality on
// parts that were not modified.
export const PLACEHOLDER_PRESENTATION_XML = placeholderPresentation;
export const PLACEHOLDER_SLIDE_XML = placeholderSlide;

// Re-export the relationship type URIs so tests can refer to them by name
// rather than copy-pasting.
export const SYNTHETIC_REL_TYPES = {
  officeDoc: REL_TYPE_OFFICE_DOC,
  slide: REL_TYPE_SLIDE,
  hyperlink: REL_TYPE_HYPERLINK,
  image: REL_TYPE_IMAGE,
} as const;

// [Content_Types].xml — required per-package map of part-name → content type.
//
// ECMA-376 Part 2 §10.1. Structure:
//
//   <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
//     <Default Extension="xml" ContentType="application/xml"/>
//     <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
//     <Override PartName="/ppt/presentation.xml" ContentType="application/vnd...."/>
//     ...
//   </Types>
//
// A part's content type is determined by: first matching Override (by part
// name, case-insensitive); if none, the Default for the part's extension
// (case-insensitive). PPTX files always have at least the default for `xml`
// and `rels`, plus overrides for most parts.

import { NS, type XmlDocument, elem, parseXml, qname, serializeXml } from '../xml/index.ts';
import { type PartName, basename, partName, partNamesEqual } from './part-name.ts';

export interface ContentTypeDefault {
  /** Lowercase, no leading dot, e.g. `xml`. */
  extension: string;
  contentType: string;
}

export interface ContentTypeOverride {
  partName: PartName;
  contentType: string;
}

export interface ContentTypes {
  defaults: ContentTypeDefault[];
  overrides: ContentTypeOverride[];
}

const CT_NS = NS.contentTypes;
const NAME_TYPES = qname('', 'Types', CT_NS);
const NAME_DEFAULT = qname('', 'Default', CT_NS);
const NAME_OVERRIDE = qname('', 'Override', CT_NS);
const ATTR_EXTENSION = qname('', 'Extension', '');
const ATTR_CONTENT_TYPE = qname('', 'ContentType', '');
const ATTR_PART_NAME = qname('', 'PartName', '');

const getAttrValueRaw = (
  e: { attrs: ReadonlyArray<{ name: { localName: string }; value: string }> },
  local: string,
): string | null => {
  for (const a of e.attrs) {
    if (a.name.localName === local) return a.value;
  }
  return null;
};

export const parseContentTypes = (xml: string): ContentTypes => {
  const doc = parseXml(xml);
  if (doc.root.name.namespaceURI !== CT_NS || doc.root.name.localName !== 'Types') {
    throw new Error(`expected root <Types xmlns="${CT_NS}">, got <${doc.root.name.localName}>`);
  }
  const defaults: ContentTypeDefault[] = [];
  const overrides: ContentTypeOverride[] = [];
  for (const child of doc.root.children) {
    if (child.kind !== 'element') continue;
    if (child.name.namespaceURI !== CT_NS) continue;
    if (child.name.localName === 'Default') {
      const ext = getAttrValueRaw(child, 'Extension');
      const ct = getAttrValueRaw(child, 'ContentType');
      if (ext === null || ct === null) {
        throw new Error('Content_Types Default missing Extension or ContentType');
      }
      defaults.push({ extension: ext.toLowerCase(), contentType: ct });
    } else if (child.name.localName === 'Override') {
      const pn = getAttrValueRaw(child, 'PartName');
      const ct = getAttrValueRaw(child, 'ContentType');
      if (pn === null || ct === null) {
        throw new Error('Content_Types Override missing PartName or ContentType');
      }
      overrides.push({ partName: partName(pn), contentType: ct });
    }
  }
  return { defaults, overrides };
};

export const serializeContentTypes = (ct: ContentTypes): string => {
  const children = [
    ...ct.defaults.map((d) =>
      elem(NAME_DEFAULT, {
        attrs: [
          { name: ATTR_EXTENSION, value: d.extension },
          { name: ATTR_CONTENT_TYPE, value: d.contentType },
        ],
      }),
    ),
    ...ct.overrides.map((o) =>
      elem(NAME_OVERRIDE, {
        attrs: [
          { name: ATTR_PART_NAME, value: o.partName },
          { name: ATTR_CONTENT_TYPE, value: o.contentType },
        ],
      }),
    ),
  ];
  const root = elem(NAME_TYPES, {
    prefixDecls: new Map([['', CT_NS]]),
    children,
  });
  const doc: XmlDocument = {
    kind: 'document',
    decl: { version: '1.0', encoding: 'UTF-8', standalone: 'yes' },
    prolog: [],
    root,
    epilog: [],
  };
  return serializeXml(doc);
};

/**
 * Resolves the content type for a given part name.
 *
 *   1. The first Override whose PartName matches case-insensitively wins.
 *   2. Otherwise, the first Default whose Extension matches the part's
 *      extension (case-insensitive) wins.
 *   3. Returns `null` if neither rule matches — `[Content_Types].xml` is
 *      required to cover every part, so a null here indicates a malformed
 *      package.
 */
export const lookupContentType = (ct: ContentTypes, name: PartName): string | null => {
  for (const o of ct.overrides) {
    if (partNamesEqual(o.partName, name)) return o.contentType;
  }
  const file = basename(name);
  const dot = file.lastIndexOf('.');
  if (dot < 0) return null;
  const ext = file.slice(dot + 1).toLowerCase();
  for (const d of ct.defaults) {
    if (d.extension.toLowerCase() === ext) return d.contentType;
  }
  return null;
};

/**
 * Convenience: a fresh ContentTypes seeded with the two defaults every PPTX
 * has (`xml` and `rels`). Authoring code can layer overrides on top.
 */
export const emptyContentTypes = (): ContentTypes => ({
  defaults: [
    { extension: 'rels', contentType: 'application/vnd.openxmlformats-package.relationships+xml' },
    { extension: 'xml', contentType: 'application/xml' },
  ],
  overrides: [],
});

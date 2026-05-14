// Semantic equivalence comparator for OPC packages.
//
// "Semantic" means: two packages model the same logical content even if
// their byte representations differ. The cases we tolerate as equal:
//
//   - ZIP DEFLATE non-determinism: two emitters can produce different bytes
//     for the same uncompressed content.
//   - XML attribute spacing, the trailing newline inside the prolog, and
//     `<x/>` vs `<x></x>` for genuinely empty elements (we do not produce
//     this asymmetry, but real fixtures sometimes do).
//   - Order of `<Default>` and `<Override>` entries in `[Content_Types].xml`.
//   - Order of `<Relationship>` entries within a single `.rels` part.
//
// We do NOT tolerate:
//
//   - Different parts list (something added or removed).
//   - Different content types for a part.
//   - Different XML structure beyond the cases above.
//   - Different bytes for binary parts (media).

import type { ContentTypes, Relationships } from '../../src/internal/opc/index.ts';
import { parseRels } from '../../src/internal/opc/index.ts';
import type { OpcPackage } from '../../src/internal/parts/index.ts';
import { parseXml } from '../../src/internal/xml/index.ts';
import type { XmlAttr, XmlElement, XmlNode } from '../../src/internal/xml/index.ts';

const decode = (b: Uint8Array): string => new TextDecoder().decode(b);

const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

const sortRelsItems = (rels: Relationships): Relationships => ({
  items: [...rels.items].sort((a, b) => a.id.localeCompare(b.id)),
});

const sortContentTypes = (ct: ContentTypes): ContentTypes => ({
  defaults: [...ct.defaults].sort((a, b) => a.extension.localeCompare(b.extension)),
  overrides: [...ct.overrides].sort((a, b) => a.partName.localeCompare(b.partName)),
});

const compareUint8 = (a: Uint8Array, b: Uint8Array): boolean => {
  if (a.byteLength !== b.byteLength) return false;
  for (let i = 0; i < a.byteLength; i++) {
    if (a[i] !== b[i]) return false;
  }
  return true;
};

const compareAttrs = (a: ReadonlyArray<XmlAttr>, b: ReadonlyArray<XmlAttr>): boolean => {
  if (a.length !== b.length) return false;
  // Attribute order is significant for OOXML round-trip, but for SEMANTIC
  // equivalence we treat it as unordered.
  const ka = a.map((x) => `${x.name.namespaceURI}|${x.name.localName}=${x.value}`).sort();
  const kb = b.map((x) => `${x.name.namespaceURI}|${x.name.localName}=${x.value}`).sort();
  for (let i = 0; i < ka.length; i++) {
    if (ka[i] !== kb[i]) return false;
  }
  return true;
};

const compareElements = (a: XmlElement, b: XmlElement): boolean => {
  if (a.name.namespaceURI !== b.name.namespaceURI) return false;
  if (a.name.localName !== b.name.localName) return false;
  if (!compareAttrs(a.attrs, b.attrs)) return false;
  const ca = a.children.filter((c) => c.kind !== 'text' || c.data.trim().length > 0);
  const cb = b.children.filter((c) => c.kind !== 'text' || c.data.trim().length > 0);
  if (ca.length !== cb.length) return false;
  for (let i = 0; i < ca.length; i++) {
    if (!compareNodes(ca[i] as XmlNode, cb[i] as XmlNode)) return false;
  }
  return true;
};

const compareNodes = (a: XmlNode, b: XmlNode): boolean => {
  if (a.kind !== b.kind) return false;
  switch (a.kind) {
    case 'element':
      return compareElements(a, b as XmlElement);
    case 'text':
      return a.data.trim() === (b as { data: string }).data.trim();
    case 'cdata':
      return a.data === (b as { data: string }).data;
    case 'comment':
      return a.data === (b as { data: string }).data;
    case 'pi':
      return (
        a.target === (b as { target: string }).target && a.data === (b as { data: string }).data
      );
  }
};

const compareXmlBytes = (a: Uint8Array, b: Uint8Array): boolean => {
  const da = parseXml(decode(a));
  const db = parseXml(decode(b));
  return compareElements(da.root, db.root);
};

/**
 * Returns true if both packages model the same logical PPTX content.
 *
 * Throws on the first difference with a human-readable message identifying
 * the part and the kind of mismatch — useful in test failure messages.
 */
export const expectSemanticallyEqual = (left: OpcPackage, right: OpcPackage): void => {
  // Content_Types
  const leftCT = sortContentTypes(left.contentTypes);
  const rightCT = sortContentTypes(right.contentTypes);
  if (leftCT.defaults.length !== rightCT.defaults.length) {
    throw new Error(
      `Content_Types defaults count: left=${leftCT.defaults.length} right=${rightCT.defaults.length}`,
    );
  }
  for (let i = 0; i < leftCT.defaults.length; i++) {
    const a = leftCT.defaults[i];
    const b = rightCT.defaults[i];
    if (a?.extension !== b?.extension || a?.contentType !== b?.contentType) {
      throw new Error(
        `Content_Types Default[${i}] mismatch: ${JSON.stringify(a)} vs ${JSON.stringify(b)}`,
      );
    }
  }
  if (leftCT.overrides.length !== rightCT.overrides.length) {
    throw new Error(
      `Content_Types overrides count: left=${leftCT.overrides.length} right=${rightCT.overrides.length}`,
    );
  }
  for (let i = 0; i < leftCT.overrides.length; i++) {
    const a = leftCT.overrides[i];
    const b = rightCT.overrides[i];
    if (a?.partName !== b?.partName || a?.contentType !== b?.contentType) {
      throw new Error(
        `Content_Types Override[${i}] mismatch: ${JSON.stringify(a)} vs ${JSON.stringify(b)}`,
      );
    }
  }

  // Parts list (order-independent)
  const leftNames = new Set<string>(left.parts.map((p: { name: string }) => p.name));
  const rightNames = new Set<string>(right.parts.map((p: { name: string }) => p.name));
  if (leftNames.size !== rightNames.size) {
    throw new Error(`part count mismatch: left=${leftNames.size} right=${rightNames.size}`);
  }
  for (const name of leftNames) {
    if (!rightNames.has(name)) throw new Error(`part missing on right: ${name}`);
  }
  for (const name of rightNames) {
    if (!leftNames.has(name)) throw new Error(`part missing on left: ${name}`);
  }

  // Each part's bytes / XML
  for (const partLeft of left.parts) {
    const partRight = right.getPart(partLeft.name);
    if (!partRight) throw new Error(`unreachable: ${partLeft.name}`);
    if (partLeft.contentType !== partRight.contentType) {
      throw new Error(
        `${partLeft.name} contentType: left="${partLeft.contentType}" right="${partRight.contentType}"`,
      );
    }
    // Decide compare strategy by content type / extension.
    if (partLeft.name.endsWith('.rels')) {
      const a = sortRelsItems(parseRels(decode(partLeft.data)));
      const b = sortRelsItems(parseRels(decode(partRight.data)));
      if (a.items.length !== b.items.length) {
        throw new Error(`${partLeft.name} rels count mismatch`);
      }
      for (let i = 0; i < a.items.length; i++) {
        const ai = a.items[i];
        const bi = b.items[i];
        if (
          ai?.id !== bi?.id ||
          ai?.type !== bi?.type ||
          ai?.target !== bi?.target ||
          ai?.targetMode !== bi?.targetMode
        ) {
          throw new Error(
            `${partLeft.name} rels[${i}] mismatch: ${JSON.stringify(ai)} vs ${JSON.stringify(bi)}`,
          );
        }
      }
    } else if (
      partLeft.contentType.endsWith('+xml') ||
      partLeft.contentType === 'application/xml'
    ) {
      if (!compareXmlBytes(partLeft.data, partRight.data)) {
        throw new Error(`${partLeft.name} XML semantic mismatch`);
      }
    } else {
      if (!compareUint8(partLeft.data, partRight.data)) {
        throw new Error(`${partLeft.name} binary bytes differ`);
      }
    }
  }

  // (The RELS_NS / CONTENT_TYPES_NS constants below are exported only so
  // tests that want to check namespace URIs against canonical strings can
  // import them.)
};

export { RELS_NS, CONTENT_TYPES_NS };
// And a re-export to keep type-only consumers happy.
export type { ContentTypes, Relationships };

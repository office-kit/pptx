// internal/opc — OPC package: ZIP, [Content_Types].xml, .rels.
// Allowed imports: internal/xml.

export type { CompressionLevel, ZipEntry, ZipReadResult } from './zip.ts';
export { readZip, shouldStore, writeZip } from './zip.ts';

export type { PartName } from './part-name.ts';
export {
  basename,
  dirname,
  fromZipPath,
  partName,
  partNamesEqual,
  relsPartNameFor,
  resolveTarget,
  toZipPath,
} from './part-name.ts';

export type { ContentTypeDefault, ContentTypeOverride, ContentTypes } from './content-types.ts';
export {
  emptyContentTypes,
  lookupContentType,
  parseContentTypes,
  serializeContentTypes,
} from './content-types.ts';

export type { Relationship, Relationships, TargetMode } from './rels.ts';
export { emptyRels, nextRelId, parseRels, serializeRels } from './rels.ts';

export type { ImageFormat } from './image-format.ts';
export { contentTypeForFormat, detectImageFormat, extensionForFormat } from './image-format.ts';

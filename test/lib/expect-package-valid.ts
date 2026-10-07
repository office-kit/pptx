// Saves a presentation, reloads it, and schema-validates every XML part whose
// schema the suite has: the PresentationML parts, themes and relationships.
// Master edits touch several parts at once (presentation.xml, its rels, the
// master, its layouts, a theme), so a test checks the whole package rather
// than naming parts.

import { expect } from 'vitest';
import {
  type PresentationData,
  listPackageParts,
  loadPresentation,
  readPackagePart,
  savePresentation,
  validatePresentation,
} from '../../src/api/index.ts';
import {
  expectSchemaValidAll,
  isSchemaValidationAvailable,
  type SchemaKind,
} from './expect-schema-valid.ts';

const decoder = new TextDecoder();

const schemaKind = (name: string, contentType: string): SchemaKind | null => {
  if (name.endsWith('.rels')) return 'rels';
  if (contentType.includes('presentationml') && contentType.endsWith('+xml')) return 'pml';
  if (contentType.endsWith('theme+xml')) return 'dml';
  return null;
};

export const expectPackageValid = async (pres: PresentationData): Promise<PresentationData> => {
  const reloaded = await loadPresentation(await savePresentation(pres));
  expect(validatePresentation(reloaded)).toEqual([]);
  if (!isSchemaValidationAvailable()) return reloaded;
  // One xmllint run per schema: a package has dozens of parts.
  const documents = new Map<SchemaKind, { name: string; xml: string }[]>();
  for (const part of listPackageParts(reloaded)) {
    const kind = schemaKind(part.name, part.contentType);
    if (kind === null) continue;
    const xml = decoder.decode(readPackagePart(reloaded, part.name)!);
    documents.set(kind, [...(documents.get(kind) ?? []), { name: part.name, xml }]);
  }
  for (const [kind, parts] of documents) expectSchemaValidAll(parts, kind);
  return reloaded;
};

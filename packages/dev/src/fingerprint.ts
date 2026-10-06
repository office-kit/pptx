import { createHash } from 'node:crypto';
import { strFromU8, strToU8, unzipSync } from 'fflate';

const CORE_PROPERTIES = 'docProps/core.xml';
const CORE_TIMESTAMPS = /<(dcterms:(?:created|modified))\b[^>]*>[\s\S]*?<\/\1>/g;

/** Every build stamps fresh creation and modification times into the core properties. */
export function comparablePart(name: string, bytes: Uint8Array): Uint8Array {
  if (name !== CORE_PROPERTIES) return bytes;
  return strToU8(strFromU8(bytes).replace(CORE_TIMESTAMPS, ''));
}

/** ZIP container dates and generated core timestamps are not source edits. */
export function sourceFingerprint(bytes: Uint8Array): string {
  const parts = unzipSync(bytes);
  const hash = createHash('sha256');
  for (const name of Object.keys(parts).sort()) {
    const content = comparablePart(name, parts[name]!);
    hash.update(JSON.stringify([name, createHash('sha256').update(content).digest('hex')]));
  }
  return hash.digest('hex');
}

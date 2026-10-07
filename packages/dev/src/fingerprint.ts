import { createHash } from 'node:crypto';
import { comparablePart } from '@office-kit/pptx-editor/merge';
import { unzipSync } from 'fflate';

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

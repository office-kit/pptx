// loadPresentation / loadPresentationFile with `limits` — opening untrusted
// decks. The crafted-archive cases live next to the reader in
// src/internal/opc/zip.test.ts; these check the option reaches it through
// both public loaders.

import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import {
  type LoadLimits,
  addContentSlide,
  createPresentation,
  getSlides,
  loadPresentation,
  loadPresentationFile,
  savePresentation,
} from '../src/node.ts';

const SERVER_LIMITS: LoadLimits = {
  maxEntries: 1000,
  maxTotalUncompressedBytes: 100 * 1024 * 1024,
  maxXmlPartBytes: 10 * 1024 * 1024,
  maxTotalXmlBytes: 50 * 1024 * 1024,
};

const deckBytes = async (): Promise<Uint8Array> => {
  const pres = createPresentation();
  addContentSlide(pres, { title: 'Q3', body: 'Revenue up' });
  addContentSlide(pres, { title: 'Q4', body: 'Plan' });
  return savePresentation(pres);
};

describe('loadPresentation with limits', () => {
  it('loads a normal deck under server-sized limits', async () => {
    const pres = await loadPresentation(await deckBytes(), { limits: SERVER_LIMITS });
    expect(getSlides(pres)).toHaveLength(2);
  });

  it('rejects a deck over a limit with an invalid ZIP error naming it', async () => {
    const bytes = await deckBytes();
    await expect(
      loadPresentation(bytes, { limits: { ...SERVER_LIMITS, maxEntries: 3 } }),
    ).rejects.toThrow(/^invalid ZIP: archive lists \d+ entries, over the maxEntries limit of 3$/);
    await expect(
      loadPresentation(bytes, { limits: { ...SERVER_LIMITS, maxTotalXmlBytes: 1024 } }),
    ).rejects.toThrow(/over the maxTotalXmlBytes limit of 1024$/);
  });

  it('applies limits passed to loadPresentationFile', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-kit-load-limits-'));
    try {
      const path = join(dir, 'deck.pptx');
      await writeFile(path, await deckBytes());
      expect(getSlides(await loadPresentationFile(path, { limits: SERVER_LIMITS }))).toHaveLength(
        2,
      );
      await expect(
        loadPresentationFile(path, { limits: { ...SERVER_LIMITS, maxXmlPartBytes: 16 } }),
      ).rejects.toThrow(/over the maxXmlPartBytes limit of 16$/);
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});

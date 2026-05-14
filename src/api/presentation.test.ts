import { describe, expect, it } from 'vitest';
import { buildSyntheticPackageBytes } from '../../test/lib/synthetic-package.ts';
import { Presentation } from './index.ts';

describe('Presentation.load', () => {
  it('loads a Uint8Array', async () => {
    const pres = await Presentation.load(buildSyntheticPackageBytes());
    expect(pres).toBeInstanceOf(Presentation);
  });

  it('loads an ArrayBuffer', async () => {
    const u8 = buildSyntheticPackageBytes();
    // Copy the bytes into a fresh ArrayBuffer to avoid the union of
    // ArrayBuffer | SharedArrayBuffer that recent TS lib types expose
    // on Uint8Array#buffer.
    const ab = new ArrayBuffer(u8.byteLength);
    new Uint8Array(ab).set(u8);
    const pres = await Presentation.load(ab);
    expect(pres).toBeInstanceOf(Presentation);
  });

  it('loads a Blob', async () => {
    const u8 = buildSyntheticPackageBytes();
    const ab = new ArrayBuffer(u8.byteLength);
    new Uint8Array(ab).set(u8);
    const blob = new Blob([ab]);
    const pres = await Presentation.load(blob);
    expect(pres).toBeInstanceOf(Presentation);
  });

  it('throws TypeError on unsupported input', async () => {
    // @ts-expect-error — deliberately wrong type
    await expect(Presentation.load(42)).rejects.toThrow(TypeError);
  });
});

describe('Presentation round-trip', () => {
  it('produces a valid OPC package that re-loads successfully', async () => {
    const original = buildSyntheticPackageBytes();
    const pres = await Presentation.load(original);
    const written = await pres.save();
    expect(written).toBeInstanceOf(Uint8Array);
    expect(written.length).toBeGreaterThan(0);

    // The output is itself loadable — proves we didn't break the OPC shape.
    const again = await Presentation.load(written);
    expect(again).toBeInstanceOf(Presentation);
  });
});

describe('Presentation.create', () => {
  it('returns an empty Presentation that round-trips', async () => {
    const pres = Presentation.create();
    const bytes = await pres.save();
    const reloaded = await Presentation.load(bytes);
    expect(reloaded).toBeInstanceOf(Presentation);
  });
});

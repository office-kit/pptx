// The opaque-handle symbol keys MUST live in the process-global symbol
// registry (`Symbol.for`), not be minted with plain `Symbol`. The library
// ships as two separate bundles — `pptx-kit` (dist/index.js) and
// `pptx-kit/node` (dist/node.js) — and companion packages (e.g.
// `@pptx-kit/preview`) bundle a third copy of the reader code. A handle built
// by one bundle is only readable by another if they agree on these keys; plain
// `Symbol` mints a fresh key per bundle, so e.g. `getSlides` (index bundle)
// would read `undefined` off a presentation from `loadPresentationFile` (node
// bundle) and crash. `Symbol.for` makes the keys bundle-independent.
//
// This is a white-box guard: it reaches into the internal symbols module so a
// regression (reverting to plain `Symbol`) fails fast, since the cross-bundle
// scenario itself is awkward to reproduce inside a single vitest realm.

import { describe, expect, it } from 'vitest';
import * as symbols from '../src/api/_internal-symbols.ts';

// Every exported value (the module's runtime exports are all symbols; its
// interfaces are type-only and absent here), paired with its export name.
const allEntries: ReadonlyArray<readonly [string, symbol]> = Object.entries(symbols);
const HANDLE_SYMBOLS = allEntries.filter(([, value]) => typeof value === 'symbol');

describe('internal handle symbols', () => {
  it('exports at least the known handle keys', () => {
    // Guards against the list silently emptying (e.g. a bad refactor) and the
    // per-symbol assertions below vacuously passing.
    expect(HANDLE_SYMBOLS.length).toBeGreaterThanOrEqual(16);
  });

  it('are all registered in the global symbol registry', () => {
    for (const [name, sym] of HANDLE_SYMBOLS) {
      const key = Symbol.keyFor(sym);
      // `undefined` means a plain `Symbol(...)` — the cross-bundle footgun.
      expect(key, `${name} must be a Symbol.for(...) registry symbol`).toBeDefined();
      expect(key).toMatch(/^pptx-kit\./);
    }
  });

  it('round-trip through Symbol.for resolves to the same symbol', () => {
    // The defining property a second bundle relies on: re-deriving the key
    // yields the identical symbol.
    for (const [, sym] of HANDLE_SYMBOLS) {
      const key = Symbol.keyFor(sym);
      expect(Symbol.for(key!)).toBe(sym);
    }
  });
});

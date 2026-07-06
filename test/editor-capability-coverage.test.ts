// COVERAGE GUARANTEE for the @office-kit/pptx editor UI.
//
// The editor's promise is that *every* authoring operation the library exposes
// is reachable from the UI. This test makes that a mechanically enforced fact
// rather than an aspiration:
//
//   1. It re-derives the set of mutating (state-changing) public exports
//      straight from the compiled library — the same verb-prefix rule the
//      manifest generator uses, but computed independently here.
//   2. It asserts the editor's capability manifest lists exactly that set:
//      neither missing a function (an unreachable capability) nor naming one
//      that no longer exists (a dead command).
//   3. It asserts every manifested id is a real callable on the library, so a
//      capability can't be "registered" against a typo or a removed function.
//
// Consequence: the moment someone adds a new `setX` / `addX` authoring function
// to the public API, `pnpm test` fails until it is added to the manifest (and
// thereby wired into the editor's command registry). Implementation effort can
// never silently drop a capability.

import { describe, expect, it } from 'vitest';
import * as pptx from '@office-kit/pptx';
// Import the generated data directly (pure JSON) rather than the resolved
// manifest module: the resolved manifest lives in the SvelteKit source tree
// whose tsconfig is only materialised by `svelte-kit sync`, and pulling it into
// the library's vitest run would couple the two toolchains. `overrides` never
// add or remove capabilities (guarded at runtime in `manifest/index.ts`), so
// the generated id set is exactly the resolved id set for coverage purposes.
import generated from '../site/src/lib/editor/manifest/capabilities.generated.json';

const capabilities = generated.capabilities as ReadonlyArray<{
  id: string;
  operand: string;
  category: string;
}>;
const capabilityById = new Map(capabilities.map((c) => [c.id, c]));

// Kept in lockstep with `manifest/generate.mjs::MUTATING_VERBS`.
const MUTATING_VERBS = [
  'add',
  'set',
  'clear',
  'replace',
  'remove',
  'insert',
  'duplicate',
  'bring',
  'send',
  'append',
  'group',
  'ungroup',
  'swap',
  'sort',
  'reverse',
  'rename',
  'move',
  'merge',
  'import',
  'copy',
  'create',
  'translate',
  'touch',
  'increment',
  'compact',
];

function isMutatingName(name: string): boolean {
  return MUTATING_VERBS.some(
    (v) =>
      name.startsWith(v) &&
      name.length > v.length &&
      name[v.length] === name[v.length]!.toUpperCase(),
  );
}

const libraryMutatingExports = Object.entries(pptx)
  .filter(([name, value]) => typeof value === 'function' && isMutatingName(name))
  .map(([name]) => name)
  .sort();

const manifestIds = capabilities.map((c) => c.id).sort();

describe('editor capability coverage', () => {
  it('manifests every mutating public export (no unreachable capability)', () => {
    const missing = libraryMutatingExports.filter((name) => !capabilityById.has(name));
    expect(
      missing,
      `these library authoring functions are missing from the editor manifest:\n${missing.join('\n')}`,
    ).toEqual([]);
  });

  it('has no manifest entry that does not exist in the library (no dead command)', () => {
    const dead = manifestIds.filter((id) => !(id in pptx));
    expect(
      dead,
      `these manifest capabilities reference functions not exported by the library:\n${dead.join('\n')}`,
    ).toEqual([]);
  });

  it('binds every capability to a real callable function', () => {
    const notCallable = manifestIds.filter(
      (id) => typeof (pptx as Record<string, unknown>)[id] !== 'function',
    );
    expect(notCallable, `not callable on the library:\n${notCallable.join('\n')}`).toEqual([]);
  });

  it('manifest and library mutating sets are exactly equal', () => {
    expect(manifestIds).toEqual(libraryMutatingExports);
  });

  it('assigns every capability an operand and a category', () => {
    const bad = capabilities.filter((c) => !c.operand || !c.category);
    expect(bad.map((c) => c.id)).toEqual([]);
  });
});

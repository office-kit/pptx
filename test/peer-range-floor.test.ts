import { readdirSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { expect, it } from 'vitest';

const CORE = '@office-kit/pptx';
const CARET_ON_CORE_MINOR = /^\^0\.(\d+)\.\d+$/;

const readManifest = (dir: string) =>
  JSON.parse(readFileSync(resolve(dir, 'package.json'), 'utf8')) as {
    name: string;
    version: string;
    peerDependencies?: Record<string, string>;
  };

const core = readManifest('.');
const packages = readdirSync('packages').map((entry) => readManifest(resolve('packages', entry)));

// Release config sets onlyUpdatePeerDependentsWhenOutOfRange, so an open `>=`
// peer range keeps every future core release in range and changesets never
// republishes the dependent. That is how the DSL fix for the 0.20.0
// setShapeBullets rename stayed unpublished while npm still installed the two
// side by side. A caret range goes out of range on the next breaking core
// minor, which is what makes the bump automatic.
it('ranges every core peer dependency on the current pre-1.0 minor', () => {
  const [major, minor] = core.version.split('.');
  expect(major).toBe('0');
  const declared = packages
    .filter((pkg) => pkg.peerDependencies?.[CORE] !== undefined)
    .map((pkg) => [pkg.name, pkg.peerDependencies![CORE]!] as const);
  expect(declared.length).toBeGreaterThan(0);
  for (const [name, range] of declared) {
    expect(range, `${name} peer range on ${CORE}`).toMatch(CARET_ON_CORE_MINOR);
    expect(CARET_ON_CORE_MINOR.exec(range)![1], `${name} peer range on ${CORE}`).toBe(minor);
  }
});

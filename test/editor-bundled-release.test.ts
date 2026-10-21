import { readdirSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { expect, it } from 'vitest';

const EDITOR = '@office-kit/pptx-editor';
const DEV = '@office-kit/pptx-dev';
const FRONTMATTER = /^---\n([\s\S]*?)\n---/;

// pptx-dev bundles the editor and lists it only as a devDependency, which
// changesets does not follow, so an editor change ships to dev users only when
// the same changeset also releases dev.
it('releases pptx-dev with every editor changeset', () => {
  const directory = resolve('.changeset');
  for (const file of readdirSync(directory).filter((name) => name.endsWith('.md'))) {
    const header = FRONTMATTER.exec(readFileSync(resolve(directory, file), 'utf8'))?.[1] ?? '';
    const released = new Set(
      header.split('\n').map((line) =>
        line
          .split(':')[0]!
          .trim()
          .replace(/^['"]|['"]$/g, ''),
      ),
    );
    if (released.has(EDITOR)) expect(released.has(DEV), `${file} releases ${EDITOR}`).toBe(true);
  }
});

// Part name utilities per OPC (ECMA-376 Part 2 §9.1.1).
//
// A part name is a URI path that begins with `/`, uses `/` as segment
// separator, and has no trailing slash (except the package root which is
// outside the scope of this module). Within a ZIP, the file name is the part
// name with the leading `/` removed.
//
// Part names are compared case-insensitively. We normalize them to lowercase
// when comparing but preserve the original casing for serialization, because
// the casing on disk is significant for compatibility with some tools.

export type PartName = string & { readonly __brand: 'PartName' };

/**
 * Asserts the given string is a syntactically valid OPC part name and returns
 * it as a branded `PartName`. Throws otherwise.
 *
 * Rules enforced (subset; see ECMA-376 Part 2 §9.1.1.1-9.1.1.6):
 *
 *   - Must start with `/`.
 *   - Must not end with `/`.
 *   - Must not contain `//` (empty segments).
 *   - Each segment must be non-empty and must not be `.` or `..`.
 */
export const partName = (raw: string): PartName => {
  if (raw === '') throw new Error('part name is empty');
  if (!raw.startsWith('/')) throw new Error(`part name must start with "/": ${raw}`);
  if (raw.length > 1 && raw.endsWith('/')) {
    throw new Error(`part name must not end with "/": ${raw}`);
  }
  const segments = raw.slice(1).split('/');
  for (const s of segments) {
    if (s === '') throw new Error(`part name contains empty segment: ${raw}`);
    if (s === '.' || s === '..') {
      throw new Error(`part name contains "." or ".." segment: ${raw}`);
    }
  }
  return raw as PartName;
};

/** Convert a part name to the corresponding ZIP entry path (no leading slash). */
export const toZipPath = (name: PartName): string => name.slice(1);

/** Convert a ZIP entry path to a part name (adds leading slash). */
export const fromZipPath = (zipPath: string): PartName => partName(`/${zipPath}`);

/** Case-insensitive part name equality per OPC §9.1.1.7. */
export const partNamesEqual = (a: PartName, b: PartName): boolean =>
  a.toLowerCase() === b.toLowerCase();

/**
 * Returns the directory portion of a part name (everything before the final
 * `/`). The result keeps the leading slash but has no trailing slash, except
 * for the root which returns `/`.
 */
export const dirname = (name: PartName): string => {
  const idx = name.lastIndexOf('/');
  if (idx <= 0) return '/';
  return name.slice(0, idx);
};

/** Returns the file portion of a part name (everything after the final `/`). */
export const basename = (name: PartName): string => {
  const idx = name.lastIndexOf('/');
  return idx < 0 ? name : name.slice(idx + 1);
};

/**
 * Computes the part name of the `.rels` file that holds relationships
 * originating from `name`. Per OPC §9.3.2: for a part `/x/y/z.xml` the rels
 * file is `/x/y/_rels/z.xml.rels`. The package-root rels file is `/_rels/.rels`.
 */
export const relsPartNameFor = (name: PartName): PartName => {
  const dir = dirname(name);
  const file = basename(name);
  const prefix = dir === '/' ? '/_rels/' : `${dir}/_rels/`;
  return partName(`${prefix}${file}.rels`);
};

/**
 * Resolves a relationship Target (which may be relative) against a base part
 * name. Returns the absolute part name (always starts with `/`).
 *
 * `target` is from a Relationship element's `Target` attribute when
 * `TargetMode="Internal"`. External targets must not be passed through here.
 */
export const resolveTarget = (basePart: PartName, target: string): PartName => {
  if (target === '') throw new Error('relationship target is empty');
  if (target.startsWith('/')) return partName(target);

  const baseDir = dirname(basePart);
  const stack = baseDir === '/' ? [] : baseDir.slice(1).split('/');
  for (const segment of target.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') {
      if (stack.length === 0) {
        throw new Error(`relationship target "${target}" escapes the package root`);
      }
      stack.pop();
      continue;
    }
    stack.push(segment);
  }
  return partName(`/${stack.join('/')}`);
};

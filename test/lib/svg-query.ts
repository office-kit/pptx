// Minimal SVG query helpers for unit tests.
//
// These work on SVG strings without a DOM. Regex-based parsing is acceptable
// here because we only ever parse SVG that we ourselves emit: attributes are
// always double-quoted, attribute values never contain `>`, and tag names are
// predictable ASCII. A general-purpose HTML/XML parser is intentionally absent
// to keep the test helper dependency-free.

/**
 * Counts how many times `<tagName` appears in `svg` (opening tags only).
 * Matches regardless of namespace prefix — e.g. `countTags(svg, 'image')`
 * matches `<image`, `<svg:image`, etc.
 */
export const countTags = (svg: string, tagName: string): number => {
  const re = new RegExp(`<(?:[a-zA-Z0-9_-]+:)?${tagName}[\\s/>]`, 'g');
  return (svg.match(re) ?? []).length;
};

/** A map of attribute name → value for a single tag occurrence. */
export type AttrMap = Record<string, string>;

/**
 * Returns an array of attribute maps for every opening `<tagName ...>` found
 * in `svg`. Only double-quoted attribute values are captured; unquoted or
 * single-quoted attributes in our own SVG output would be a bug in the
 * renderer, so we do not handle them here.
 */
export const attrsOf = (svg: string, tagName: string): AttrMap[] => {
  const tagRe = new RegExp(`<(?:[a-zA-Z0-9_-]+:)?${tagName}(\\s[^>]*)?>`, 'g');
  const attrRe = /([a-zA-Z_:][a-zA-Z0-9_:.-]*)="([^"]*)"/g;
  const results: AttrMap[] = [];
  let tagMatch: RegExpExecArray | null;
  while ((tagMatch = tagRe.exec(svg)) !== null) {
    const attrs: AttrMap = {};
    const inner = tagMatch[1] ?? '';
    let attrMatch: RegExpExecArray | null;
    while ((attrMatch = attrRe.exec(inner)) !== null) {
      attrs[attrMatch[1]!] = attrMatch[2]!;
    }
    results.push(attrs);
  }
  return results;
};

/**
 * Strips all XML/SVG tags from `svg` and returns the concatenated text content.
 * Useful for asserting that specific text appears somewhere in the output
 * regardless of how the renderer structures its `<text>` / `<tspan>` elements.
 */
export const textContentOf = (svg: string): string => svg.replace(/<[^>]*>/g, '');

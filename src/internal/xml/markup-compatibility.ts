// Markup Compatibility (ISO/IEC 29500-3 §10.2). An `mc:AlternateContent`
// offers one or more `mc:Choice` branches, each naming the namespaces it needs
// in `Requires`, and an optional last `mc:Fallback`. A reader takes the first
// choice whose namespaces it all understands, else the fallback.

import { qname } from './ast.ts';
import type { XmlElement } from './ast.ts';
import { NS } from './namespaces.ts';
import { getAttrValue } from './query.ts';

const ATTR_REQUIRES = qname('', 'Requires', '');

export const isAlternateContent = (el: XmlElement): boolean =>
  el.name.namespaceURI === NS.mc && el.name.localName === 'AlternateContent';

/** The prefix bindings in effect inside `el`, given those in effect around it. */
export const withPrefixDecls = (
  scope: ReadonlyMap<string, string>,
  el: XmlElement,
): ReadonlyMap<string, string> =>
  el.prefixDecls.size === 0 ? scope : new Map([...scope, ...el.prefixDecls]);

/**
 * The branch of `alternateContent` a reader that understands `understood`
 * (namespace URIs) takes, or `null` when no choice qualifies and there is no
 * fallback. `scope` holds the prefix bindings in effect around the element;
 * `Requires` names prefixes, so they resolve against it.
 */
export const selectAlternateContentBranch = (
  alternateContent: XmlElement,
  understood: ReadonlySet<string>,
  scope: ReadonlyMap<string, string>,
): XmlElement | null => {
  const outer = withPrefixDecls(scope, alternateContent);
  for (const branch of alternateContent.children) {
    if (branch.kind !== 'element' || branch.name.namespaceURI !== NS.mc) continue;
    if (branch.name.localName === 'Fallback') return branch;
    if (branch.name.localName !== 'Choice') continue;
    const inner = withPrefixDecls(outer, branch);
    const requires = (getAttrValue(branch, ATTR_REQUIRES) ?? '').split(/\s+/).filter(Boolean);
    if (requires.every((prefix) => understood.has(inner.get(prefix) ?? ''))) return branch;
  }
  return null;
};

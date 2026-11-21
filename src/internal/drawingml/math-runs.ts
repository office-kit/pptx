// Locates the run properties inside a paragraph's equations, for formatting
// that applies to all of a shape's text. Kept apart from math.ts, which
// depends on text-format.ts, so text-format.ts can use it without a cycle.

import { NS, elem, firstChildElement, qname, type XmlElement } from '../xml/index.ts';

const NAME_A_RPR = qname('a', 'rPr', NS.dml);

const isMath = (el: XmlElement, localName: string): boolean =>
  el.name.namespaceURI === NS.math && el.name.localName === localName;

// A run's `<a:rPr>`, created where the schema puts it when absent: first in
// `<a:r>`, after the optional `<m:rPr>` in `<m:r>`.
const runProperties = (run: XmlElement): XmlElement => {
  const existing = firstChildElement(run, NAME_A_RPR);
  if (existing) return existing;
  const rPr = elem(NAME_A_RPR);
  const first = run.children[0];
  const afterMathProps = first?.kind === 'element' && isMath(first, 'rPr') ? 1 : 0;
  run.children.splice(afterMathProps, 0, rPr);
  return rPr;
};

const collectMathRuns = (el: XmlElement, out: XmlElement[]): void => {
  for (const child of el.children) {
    if (child.kind !== 'element') continue;
    if (isMath(child, 'r')) out.push(runProperties(child));
    else collectMathRuns(child, out);
  }
};

const isZone = (el: XmlElement): boolean =>
  el.name.namespaceURI === NS.a14 && el.name.localName === 'm';

/**
 * The `<a:rPr>` of every run in `paragraph`'s equations: the math runs of
 * each `<a14:m>`, and the text runs of the `mc:Fallback` the library writes
 * beside one. Missing properties are created.
 */
export const equationRunProperties = (
  paragraph: XmlElement,
): { readonly math: XmlElement[]; readonly fallback: XmlElement[] } => {
  const math: XmlElement[] = [];
  const fallback: XmlElement[] = [];
  for (const child of paragraph.children) {
    if (child.kind !== 'element') continue;
    if (isZone(child)) {
      collectMathRuns(child, math);
      continue;
    }
    if (child.name.namespaceURI !== NS.mc || child.name.localName !== 'AlternateContent') continue;
    const branches = child.children.filter(
      (branch): branch is XmlElement =>
        branch.kind === 'element' && branch.name.namespaceURI === NS.mc,
    );
    const choice = branches.find(
      (branch) =>
        branch.name.localName === 'Choice' &&
        branch.children.some((zone) => zone.kind === 'element' && isZone(zone)),
    );
    if (!choice) continue;
    collectMathRuns(choice, math);
    for (const branch of branches) {
      if (branch.name.localName !== 'Fallback') continue;
      for (const run of branch.children) {
        if (
          run.kind === 'element' &&
          run.name.namespaceURI === NS.dml &&
          run.name.localName === 'r'
        ) {
          fallback.push(runProperties(run));
        }
      }
    }
  }
  return { math, fallback };
};

// Minimal read-only access to `a:txBody` text content.
//
// ECMA-376 Part 1 §21.1.2.2 defines text bodies as a sequence of paragraphs:
//
//   <a:txBody>
//     <a:bodyPr .../>          (formatting properties — not modeled here)
//     <a:lstStyle .../>        (list styles — not modeled here)
//     <a:p>                    (paragraph)
//       <a:pPr/>               (paragraph properties — not modeled here)
//       <a:r>                  (run)
//         <a:rPr/>             (run properties — not modeled here)
//         <a:t>...</a:t>       (literal text)
//       </a:r>
//       <a:fld/>               (field — text generated at render time)
//       <a:br/>                (line break)
//     </a:p>
//   </a:txBody>
//
// At this phase we only need text extraction. Authoring and formatting
// cascade resolution land later, when DrawingML grows real width.

import { NS, allChildElements, qname, textContent } from '../xml/index.ts';
import type { XmlElement } from '../xml/index.ts';

const NAME_P = qname('a', 'p', NS.dml);
const NAME_R = qname('a', 'r', NS.dml);
const NAME_T = qname('a', 't', NS.dml);
const NAME_FLD = qname('a', 'fld', NS.dml);
const NAME_BR = qname('a', 'br', NS.dml);

/**
 * Returns the visible text of a single paragraph (`a:p`). Concatenates all
 * `a:t` runs and field results in document order; treats `a:br` as a
 * newline within the paragraph.
 */
export const paragraphText = (paragraph: XmlElement): string => {
  let out = '';
  for (const child of paragraph.children) {
    if (child.kind !== 'element') continue;
    if (child.name.namespaceURI !== NS.dml) continue;
    switch (child.name.localName) {
      case 'r':
      case 'fld': {
        const t = allChildElements(child, NAME_T)[0];
        if (t) out += textContent(t);
        break;
      }
      case 'br':
        out += '\n';
        break;
      default:
        // a:rPr, a:pPr, other formatting elements — ignored for text.
        break;
    }
  }
  return out;
};

/**
 * Returns the visible text of an `a:txBody` element, joining each
 * paragraph with `\n`. Empty paragraphs become empty lines, so a body of
 * three paragraphs always produces two newlines.
 */
export const textBodyText = (txBody: XmlElement): string => {
  const parts: string[] = [];
  for (const p of allChildElements(txBody, NAME_P)) {
    parts.push(paragraphText(p));
  }
  return parts.join('\n');
};

/** Iterates over each paragraph (`a:p`) element inside `a:txBody`. */
export const paragraphsOf = (txBody: XmlElement): ReadonlyArray<XmlElement> =>
  allChildElements(txBody, NAME_P);

/** Iterates over each run (`a:r`) element inside an `a:p` element. */
export const runsOf = (paragraph: XmlElement): ReadonlyArray<XmlElement> =>
  allChildElements(paragraph, NAME_R);

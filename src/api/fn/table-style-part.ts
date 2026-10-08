// Writes the reference desktop app's definition of a built-in table style into the
// presentation's `tableStyles.xml` when a table starts using it.
//
// The reference desktop app itself draws a built-in GUID from its own definitions, but it
// also serializes the definition of every style a deck uses, and other
// consumers (Keynote, Google Slides, LibreOffice) render from that part.

import { emptyRels, nextRelId, partName, resolveTarget } from '../../internal/opc/index.ts';
import type { OpcPackage } from '../../internal/parts/index.ts';
import {
  DEFAULT_TABLE_STYLE_ID,
  REL_TYPES,
  builtinTableStyleXml,
} from '../../internal/presentationml/index.ts';
import {
  NS,
  allChildElements,
  getAttrValue,
  parseXml,
  qname,
  serializeXml,
} from '../../internal/xml/index.ts';
import { PRES_PART_NAME, decode, encode } from './_helpers.ts';

const TABLE_STYLES_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml';
const NAME_A_TBL_STYLE = qname('a', 'tblStyle', NS.dml);

/** No-op unless `styleId` is a built-in GUID the part does not define yet. */
export const ensureBuiltinTableStyleDefinition = (pkg: OpcPackage, styleId: string): void => {
  const definition = builtinTableStyleXml(styleId);
  if (definition === null) return;
  const id = styleId.trim().toUpperCase();
  const presRels = pkg.getRels(PRES_PART_NAME) ?? emptyRels();
  const rel = presRels.items.find(
    (item) => item.type === REL_TYPES.tableStyles && item.targetMode === 'Internal',
  );
  const name = rel ? resolveTarget(PRES_PART_NAME, rel.target) : partName('/ppt/tableStyles.xml');
  const part = pkg.getPart(name);
  if (part === null) {
    const xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      `<a:tblStyleLst xmlns:a="${NS.dml}" def="${DEFAULT_TABLE_STYLE_ID}">${definition}</a:tblStyleLst>`;
    pkg.addPart(name, TABLE_STYLES_CONTENT_TYPE, encode(xml));
    if (!rel) {
      presRels.items.push({
        id: nextRelId(presRels.items.map((item) => item.id)),
        type: REL_TYPES.tableStyles,
        target: name,
        targetMode: 'Internal',
      });
      pkg.setRels(PRES_PART_NAME, presRels);
    }
    return;
  }
  const doc = parseXml(decode(part.data));
  const defined = allChildElements(doc.root, NAME_A_TBL_STYLE).some(
    (style) =>
      getAttrValue(style, qname('', 'styleId', ''))
        ?.trim()
        .toUpperCase() === id,
  );
  if (defined) return;
  const style = parseXml(`<a:tblStyleLst xmlns:a="${NS.dml}">${definition}</a:tblStyleLst>`).root
    .children[0];
  if (style?.kind !== 'element') throw new Error('builtin table style markup is not an element');
  // The appended style uses the `a` prefix; declare it unless the list
  // already binds `a` to DrawingML.
  if (doc.root.prefixDecls.get('a') !== NS.dml) style.prefixDecls.set('a', NS.dml);
  doc.root.children.push(style);
  part.data = encode(serializeXml(doc));
};

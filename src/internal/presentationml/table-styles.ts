// The reference desktop app's built-in DrawingML table styles. The reference desktop app identifies
// them by GUID alone and draws them from its own definitions, so a deck may reference
// one without carrying it in `tableStyles.xml`.

import { NS, parseXml, type XmlElement } from '../xml/index.ts';
import {
  BUILTIN_TABLE_STYLE_BODIES,
  BUILTIN_TABLE_STYLE_DEFINITIONS,
  BUILTIN_TABLE_STYLES,
} from './builtin-table-styles.generated.ts';

export { BUILTIN_TABLE_STYLES };

/** English name the reference desktop app writes as `styleName` for a built-in table style. */
export type BuiltinTableStyleName = (typeof BUILTIN_TABLE_STYLES)[number]['name'];

/** One entry of {@link BUILTIN_TABLE_STYLES}. */
export type BuiltinTableStyle = (typeof BUILTIN_TABLE_STYLES)[number];

const builtinNames = new Map<string, string>(BUILTIN_TABLE_STYLES.map((s) => [s.name, s.id]));
const builtinIds = new Map<string, string>(BUILTIN_TABLE_STYLES.map((s) => [s.id, s.name]));

/** GUID of the built-in style with this English name, or `null`. */
export const builtinTableStyleIdByName = (name: string): string | null =>
  builtinNames.get(name) ?? null;

/**
 * The `a:tblStyle` markup the reference desktop app writes into `tableStyles.xml` for a
 * built-in style (without a namespace declaration), or `null` when the GUID is
 * not built in.
 */
export const builtinTableStyleXml = (styleId: string): string | null => {
  const id = styleId.trim().toUpperCase();
  const definition = BUILTIN_TABLE_STYLE_DEFINITIONS[id];
  const name = builtinIds.get(id);
  if (!definition || name === undefined) return null;
  const [body, accentA, accentB] = definition;
  const inner = BUILTIN_TABLE_STYLE_BODIES[body]!.replaceAll(
    'val="$A"',
    `val="accent${accentA}"`,
  ).replaceAll('val="$B"', `val="accent${accentB}"`);
  return `<a:tblStyle styleId="${id}" styleName="${name}">${inner}</a:tblStyle>`;
};

// Parsing is lazy and cached: a deck typically uses one or two styles.
const builtinStyles = new Map<string, XmlElement>();

export const getBuiltinTableStyle = (styleId: string): XmlElement | null => {
  const id = styleId.trim().toUpperCase();
  const cached = builtinStyles.get(id);
  if (cached) return cached;
  const xml = builtinTableStyleXml(id);
  if (xml === null) return null;
  const style = parseXml(xml.replace('<a:tblStyle ', `<a:tblStyle xmlns:a="${NS.dml}" `)).root;
  builtinStyles.set(id, style);
  return style;
};

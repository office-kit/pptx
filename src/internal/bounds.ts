// ECMA-376 simple-type numeric/string bounds, enforced at the authoring boundary.
//
// A recurring defect class in a typed OOXML writer: a caller-supplied number is
// rounded and serialized straight into a constrained attribute (a coordinate, a
// line width, a font size, a duration, a percentage, an angle), and a value
// outside the schema's range produces a `.pptx` that PowerPoint marks corrupt
// and "repairs". Authoring input is an external boundary (per the project's
// defensive-programming rule: validate at boundaries, trust internally), so the
// value is checked HERE, once, before it reaches the wire. An out-of-range value
// is a caller error and throws `RangeError` naming the field and the allowed
// range — it is never silently clamped (that would hide the mistake) nor emitted.
//
// Leaf module: no imports, so any layer (drawingml / presentationml / chartml)
// may use it. The ranges below are transcribed from the ECMA-376 XSDs; the
// comment on each names the simple type so the next reader can verify it.

const RANGES = {
  coordinate: [-27273042329600, 27273042316900], // ST_Coordinate (EMU, xsd:long)
  positiveCoordinate: [0, 27273042316900], // ST_PositiveCoordinate
  coordinate32: [-2147483648, 2147483647], // ST_Coordinate32 (xsd:int)
  positiveCoordinate32: [0, 2147483647], // ST_PositiveCoordinate32
  lineWidth: [0, 20116800], // ST_LineWidth
  angle: [-2147483648, 2147483647], // ST_Angle (1/60000 degree, xsd:int)
  fontSize: [100, 400000], // ST_TextFontSize (1/100 pt → 1..4000 pt)
  textPoint: [-400000, 400000], // ST_TextPoint (1/100 pt)
  textSpacingPoint: [0, 158400], // ST_TextSpacingPoint (1/100 pt)
  unsignedInt: [0, 4294967295], // xsd:unsignedInt (advTm; ST_TLTime numeric form)
  columnCount: [1, 16], // ST_TextColumnCount
  overlap: [-100, 100], // ST_OverlapByte (bar/column overlap %)
  // ST_GapAmount unions ST_GapAmountPercent and ST_GapAmountUShort; both cap at
  // 500 (we emit the numeric/UShort form). The earlier 65535 was the raw
  // xsd:unsignedShort domain, not the schema's maxInclusive — values 501..65535
  // serialized to a `<c:gapWidth>` PowerPoint rejects.
  gapAmount: [0, 500], // ST_GapAmountUShort (bar/column gap width %)
  holeSize: [1, 90], // ST_HoleSizeUByte (doughnut hole %)
  firstSliceAng: [0, 360], // ST_FirstSliceAng (pie/doughnut start angle, degrees)
} as const;

type RangeKey = keyof typeof RANGES;

// Validates `value` as an integer in the named ECMA-376 range, returning the
// rounded integer. Non-finite or out-of-range input throws.
export const boundedInt = (value: number, key: RangeKey, field: string): number => {
  if (!Number.isFinite(value)) {
    throw new RangeError(`${field}: expected a finite number, got ${value}`);
  }
  const n = Math.round(value);
  const [min, max] = RANGES[key];
  if (n < min || n > max) {
    throw new RangeError(`${field}: ${value} is out of range for ${key} (${min}..${max})`);
  }
  return n;
};

/** EMU position coordinate (`<a:off>` x/y) — ST_Coordinate. */
export const emuCoordinate = (v: number, field: string): number =>
  boundedInt(v, 'coordinate', field);
/** EMU extent (`<a:ext>` cx/cy, column/table widths, row heights) — ST_PositiveCoordinate. */
export const emuExtent = (v: number, field: string): number =>
  boundedInt(v, 'positiveCoordinate', field);
/** 32-bit EMU (text-body insets, cell margins) — ST_Coordinate32. */
export const emuCoordinate32 = (v: number, field: string): number =>
  boundedInt(v, 'coordinate32', field);
/** Non-negative 32-bit EMU (column gap) — ST_PositiveCoordinate32. */
export const emuPositiveCoordinate32 = (v: number, field: string): number =>
  boundedInt(v, 'positiveCoordinate32', field);
/** Line width in EMU — ST_LineWidth. */
export const lineWidthEmu = (v: number, field: string): number => boundedInt(v, 'lineWidth', field);
/** Angle in 1/60000 degree — ST_Angle. */
export const angle60000 = (v: number, field: string): number => boundedInt(v, 'angle', field);
/** Font size in 1/100 pt — ST_TextFontSize. */
export const fontSizeHundredthPt = (v: number, field: string): number =>
  boundedInt(v, 'fontSize', field);
/** Character spacing in 1/100 pt — ST_TextPoint. */
export const textPointSpacing = (v: number, field: string): number =>
  boundedInt(v, 'textPoint', field);
/** Paragraph spacing in 1/100 pt — ST_TextSpacingPoint. */
export const textSpacingPoint = (v: number, field: string): number =>
  boundedInt(v, 'textSpacingPoint', field);
/** Auto-advance / timing duration in ms — xsd:unsignedInt. */
export const unsignedIntMs = (v: number, field: string): number =>
  boundedInt(v, 'unsignedInt', field);
/** Text column count — ST_TextColumnCount (1..16). */
export const textColumnCount = (v: number, field: string): number =>
  boundedInt(v, 'columnCount', field);
/** Bar/column series overlap percent — ST_OverlapByte (-100..100). */
export const overlapPercent = (v: number, field: string): number => boundedInt(v, 'overlap', field);
/** Bar/column gap width percent — ST_GapAmount (0..500). */
export const gapAmountPercent = (v: number, field: string): number =>
  boundedInt(v, 'gapAmount', field);
/** Doughnut hole size percent — ST_HoleSizeUByte (1..90). */
export const holeSizePercent = (v: number, field: string): number =>
  boundedInt(v, 'holeSize', field);
/** Pie/doughnut first-slice angle in degrees — ST_FirstSliceAng (0..360). */
export const firstSliceAngle = (v: number, field: string): number =>
  boundedInt(v, 'firstSliceAng', field);

// ST_Guid requires UPPERCASE hex inside braces. PowerPoint emits uppercase, and
// `crypto.randomUUID()` yields lowercase, so we accept either case and normalize
// to upper rather than rejecting a perfectly good GUID over case alone. A value
// that is not GUID-shaped at all is a caller error and throws.
const GUID_RE = /^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$/;
export const normalizeGuid = (value: string, field: string): string => {
  if (!GUID_RE.test(value)) {
    throw new RangeError(
      `${field}: ${JSON.stringify(value)} is not a GUID of the form {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`,
    );
  }
  return value.toUpperCase();
};

// Generic enum-membership guard for attributes whose XSD type is an enumeration
// (e.g. a pattern preset, a transition effect element name). Returns the value
// narrowed to the allowed set, or throws naming the field and the legal tokens.
export const oneOf = <T extends string>(value: string, allowed: readonly T[], field: string): T => {
  if (!(allowed as readonly string[]).includes(value)) {
    throw new RangeError(`${field}: ${JSON.stringify(value)} is not one of: ${allowed.join(', ')}`);
  }
  return value as T;
};

// EMU is the canonical internal length unit in OOXML.
// 1 inch = 914400 EMU, 1 cm = 360000 EMU, 1 pt = 12700 EMU.
// Branded so internal code can never confuse an Emu with a raw number.

declare const emuBrand: unique symbol;
export type Emu = number & { readonly [emuBrand]: 'Emu' };

const EMU_PER_INCH = 914400;
const EMU_PER_CM = 360000;
const EMU_PER_PT = 12700;

// EMU is an integer unit: `<a:off>` / `<a:ext>` coordinates are ST_Coordinate
// (xsd:long). A fractional value like `3090672.0000000005` (floating-point
// drift from unit conversion) is schema-invalid and makes PowerPoint mark the
// file corrupt and "repair" it — zeroing the offending offsets, which collapses
// shapes to the origin. Round at the unit boundary so conversions always yield
// whole EMU.
export const inches = (n: number): Emu => Math.round(n * EMU_PER_INCH) as Emu;
export const cm = (n: number): Emu => Math.round(n * EMU_PER_CM) as Emu;
export const mm = (n: number): Emu => Math.round((n * EMU_PER_CM) / 10) as Emu;
export const pt = (n: number): Emu => Math.round(n * EMU_PER_PT) as Emu;

// Escape hatch for callers that already hold an EMU value (e.g. read from an
// existing pptx). Use sparingly. Rounded too: EMU is integer-valued, and a
// fractional value here would reach the XML and trip PowerPoint's repair.
export const emu = (n: number): Emu => Math.round(n) as Emu;

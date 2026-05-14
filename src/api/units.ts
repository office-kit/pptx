// EMU is the canonical internal length unit in OOXML.
// 1 inch = 914400 EMU, 1 cm = 360000 EMU, 1 pt = 12700 EMU.
// Branded so internal code can never confuse an Emu with a raw number.

declare const emuBrand: unique symbol;
export type Emu = number & { readonly [emuBrand]: 'Emu' };

const EMU_PER_INCH = 914400;
const EMU_PER_CM = 360000;
const EMU_PER_PT = 12700;

export const inches = (n: number): Emu => (n * EMU_PER_INCH) as Emu;
export const cm = (n: number): Emu => (n * EMU_PER_CM) as Emu;
export const mm = (n: number): Emu => ((n * EMU_PER_CM) / 10) as Emu;
export const pt = (n: number): Emu => (n * EMU_PER_PT) as Emu;

// Escape hatch for callers that already hold an EMU value (e.g. read from an
// existing pptx). Use sparingly.
export const emu = (n: number): Emu => n as Emu;

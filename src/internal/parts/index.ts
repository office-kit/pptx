// internal/parts — typed parts graph, relationship resolver, ID allocators.
// Allowed imports: internal/opc, internal/xml.

export type { Part } from './package.ts';
export { OpcPackage } from './package.ts';
export type { BlankDeckAspect } from './blank-deck.ts';
export { OFFICE_THEME_XML, buildBlankDeck } from './blank-deck.ts';

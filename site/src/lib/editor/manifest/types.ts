// Types describing the capability manifest — the machine-readable catalogue of
// every authoring operation the editor must expose. See `generate.mjs` for how
// the base data is derived, `overrides.ts` for human refinements, and
// `coverage.test.ts` for the guarantee that this catalogue stays exhaustive.

/** What a capability operates on. Drives which selection makes it applicable. */
export type Operand = 'presentation' | 'slide' | 'shape' | 'cell';

/** UI-facing classification of a parameter, chosen so a generic form renderer
 *  can pick the right control. `object` falls back to a structured JSON editor. */
export type ParamKind =
  | 'string'
  | 'number'
  | 'emu'
  | 'color'
  | 'boolean'
  | 'enum'
  | 'index'
  | 'object'
  | 'array';

export interface ParamSpec {
  readonly name: string;
  /** The raw TypeScript type text, kept for the JSON-editor fallback + docs. */
  readonly type: string;
  readonly kind: ParamKind;
  readonly optional: boolean;
  readonly default?: string;
  readonly enumValues?: readonly string[];
  /** Human label (defaults to `name`). */
  readonly label?: string;
  /** For `kind: 'object'` — sub-fields to render and assemble into an object.
   *  When present the dialog builds a field-based form instead of a JSON blob. */
  readonly fields?: readonly ParamSpec[];
  /** For `kind: 'array'` — the spec of each element (repeatable in the UI). */
  readonly item?: ParamSpec;
}

/** Coarse grouping used to organise the ribbon and command palette. */
export type CategoryId =
  | 'slide'
  | 'shape'
  | 'text'
  | 'paragraph'
  | 'fill'
  | 'stroke'
  | 'effect'
  | 'image'
  | 'table'
  | 'chart'
  | 'animation'
  | 'transition'
  | 'comment'
  | 'notes'
  | 'hyperlink'
  | 'slide-background'
  | 'section'
  | 'theme'
  | 'presentation'
  | 'misc';

export interface Capability {
  /** Matches the exported function name in `@office-kit/pptx` exactly. */
  readonly id: string;
  readonly operand: Operand;
  /** True when the library function's first argument is the operand object
   *  (the common case). False for factories like `createPresentation(options)`
   *  whose first argument is a real user parameter, not the operand. */
  readonly takesOperand: boolean;
  readonly category: CategoryId;
  /** Source file the function is declared in (for docs / traceability). */
  readonly file: string;
  readonly returns: string;
  /** True for operations the on-canvas ribbon surfaces; false for
   *  package-level plumbing that only appears in the command palette. */
  readonly canvas: boolean;
  readonly params: readonly ParamSpec[];
}

/** Human-authored refinement merged over a generated capability. Every field
 *  is optional; only what is specified overrides. `params` replaces wholesale
 *  when present (so a hand-tuned schema wins over the parsed one). */
export interface CapabilityOverride {
  readonly labelEn?: string;
  readonly labelJa?: string;
  readonly category?: CategoryId;
  readonly ribbonGroup?: string;
  readonly params?: readonly ParamSpec[];
  /** Marks the primary/most-common ops for prominent ribbon placement. */
  readonly primary?: boolean;
}

/** A capability enriched with its human-authored metadata. */
export interface ResolvedCapability extends Capability {
  readonly labelEn: string;
  readonly labelJa: string;
  readonly ribbonGroup?: string;
  readonly primary: boolean;
}

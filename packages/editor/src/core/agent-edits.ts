// A small set of edits a language model applies to a presentation without an
// editor: the `@office-kit/pptx-editor/headless` entry.
//
// The operations are nine of the editor's tools, run by the same `runTool`.
// Their schemas are the generated ones narrowed to what a small revision
// needs, so the whole set stays a few kilobytes in a model's context. A
// narrowed schema is a subset of the generated one, so whatever passes it is
// also valid input for the library function.

import * as pptx from '@office-kit/pptx';
import type { PresentationData } from '@office-kit/pptx';
import { generatedTools } from '../manifest/tools.generated.ts';
import {
  type JsonSchema,
  type JsonValue,
  type ToolHost,
  runTool,
  schemaErrors,
  shapeSummary,
} from './agent-tools.ts';

const MAX_EDITS = 50;

type Properties = Readonly<Record<string, JsonSchema>>;

interface Operation {
  /** The input property holding the reference to what the edit changes. */
  readonly target: 'shape' | 'cell' | 'chart' | 'slide';
  readonly input: JsonSchema;
  readonly description: string;
  /** Arguments set after decoding, which the model cannot choose. */
  readonly fixed?: Readonly<Record<string, JsonValue>>;
}

const generated = new Map(generatedTools.tools.map((tool) => [tool.name, tool]));

function generatedTool(name: string) {
  const tool = generated.get(name);
  // The tests keep these names in the generated manifest.
  if (!tool) throw new Error(`The editor has no tool ${name}.`);
  return tool;
}

const propertiesOf = (schema: JsonSchema): Properties => (schema.properties ?? {}) as Properties;

/** `schema` (an object schema) with only the properties `keep` names. */
function pick(schema: JsonSchema, keep: readonly string[]): JsonSchema {
  const properties = propertiesOf(schema);
  const narrowed: Record<string, JsonValue> = {
    ...schema,
    properties: Object.fromEntries(keep.map((key) => [key, properties[key]!])),
  };
  // The definitions are collected once for the whole set.
  delete narrowed.$defs;
  if (Array.isArray(schema.required))
    narrowed.required = schema.required.filter((key) => keep.includes(key as string));
  return narrowed;
}

const inputOf = (name: string, keep: readonly string[]): JsonSchema =>
  pick(generatedTool(name).input_schema, keep);

// A whole-text edit that keeps the formatting of what it does not change
// (`setShapeText` without it rebuilds the text as one run per paragraph).
const PRESERVE_FORMATTING = { options: { preserveFormatting: true } } as const;

const OPERATIONS: Readonly<Record<string, Operation>> = {
  setShapeText: {
    target: 'shape',
    input: inputOf('setShapeText', ['shape', 'value']),
    description:
      "Replaces a shape's whole text (`\\n` between paragraphs) and keeps the formatting of the text before and after what changed. To change two separate places in one shape, apply it twice, each time with the whole text changed in one place: a change in between loses the formatting of the text between them. `{ case }` changes letter case instead.",
    fixed: PRESERVE_FORMATTING,
  },
  setTableCellText: {
    target: 'cell',
    input: inputOf('setTableCellText', ['cell', 'text']),
    description:
      "Replaces a table cell's whole text (`\\n` between paragraphs) and keeps the formatting of the text before and after what changed. To change two separate places in one cell, apply it twice, each time with the whole text changed in one place. `{ case }` changes letter case instead.",
    fixed: PRESERVE_FORMATTING,
  },
  setChartData: {
    target: 'chart',
    input: inputOf('setChartData', ['chart', 'data']),
    description: generatedTool('setChartData').description,
  },
  setShapeTextFormat: {
    target: 'shape',
    input: (() => {
      const input = inputOf('setShapeTextFormat', ['shape', 'format']);
      const format = propertiesOf(input).format!;
      return {
        ...input,
        properties: {
          ...propertiesOf(input),
          format: pick(format, ['size', 'bold', 'italic', 'color']),
        },
      };
    })(),
    description:
      'Sets the font size (points), bold, italic or color of all the text in a shape. Properties left out stay as they are.',
  },
  setShapeFill: {
    target: 'shape',
    input: inputOf('setShapeFill', ['shape', 'color']),
    description: generatedTool('setShapeFill').description,
  },
  setShapeBounds: {
    target: 'shape',
    input: inputOf('setShapeBounds', ['shape', 'bounds']),
    description: 'Moves and resizes a shape: `bounds` is its position and size in EMU.',
  },
  removeShape: {
    target: 'shape',
    input: inputOf('removeShape', ['shape']),
    description: generatedTool('removeShape').description,
  },
  moveSlide: {
    target: 'slide',
    input: (() => {
      const input = inputOf('moveSlide', ['slide', 'toIndex']);
      return {
        ...input,
        properties: { ...propertiesOf(input), toIndex: { type: 'integer', minimum: 0 } },
      };
    })(),
    description: 'Moves a slide to the 0-based position `toIndex`.',
  },
  removeSlide: {
    target: 'slide',
    input: inputOf('removeSlide', ['slide']),
    description: generatedTool('removeSlide').description,
  },
};

// One `$defs` for the whole set, with only the definitions the narrowed
// schemas still reach: the generated ones repeat them per tool.
function collectDefs(): Record<string, JsonSchema> {
  const available: Record<string, JsonSchema> = {};
  for (const name of Object.keys(OPERATIONS))
    Object.assign(available, generatedTool(name).input_schema.$defs);
  const defs: Record<string, JsonSchema> = {};
  const visit = (schema: JsonSchema) => {
    for (const [, ref] of JSON.stringify(schema).matchAll(/"#\/\$defs\/([A-Za-z0-9]+)"/g)) {
      if (ref! in defs) continue;
      defs[ref!] = available[ref!]!;
      visit(defs[ref!]!);
    }
  };
  for (const operation of Object.values(OPERATIONS)) visit(operation.input);
  return Object.fromEntries(Object.entries(defs).sort(([a], [b]) => a.localeCompare(b)));
}

const editsSchema = {
  type: 'array',
  minItems: 1,
  maxItems: MAX_EDITS,
  items: {
    anyOf: Object.entries(OPERATIONS).map(([name, { input, description }]) => ({
      type: 'object',
      description,
      properties: { tool: { type: 'string', const: name }, input },
      required: ['tool', 'input'],
      additionalProperties: false,
    })),
  },
} satisfies JsonSchema;

/**
 * The JSON Schema (draft 2020-12) of a batch of edits: `{ edits: [{ tool,
 * input }, …] }`, applied in order. Give it to a model as one tool's input.
 */
export const agentEditSchema: JsonSchema & { readonly type: 'object' } = {
  type: 'object',
  properties: { edits: editsSchema },
  required: ['edits'],
  additionalProperties: false,
  $defs: collectDefs(),
};

/** Why a batch of edits could not be applied. */
export class AgentEditError extends Error {
  /** The failing edit's position in the batch; `null` when the batch did not match the schema. */
  readonly index: number | null;

  constructor(index: number | null, message: string) {
    super(message);
    this.name = 'AgentEditError';
    this.index = index;
  }
}

interface DecodedAgentEdit {
  readonly tool: string;
  readonly input: Readonly<Record<string, unknown>>;
}

// Only `decodeAgentEdits` makes the type, so the fixed arguments cannot be skipped.
const decoded: unique symbol = Symbol('decoded');

/** Edits `decodeAgentEdits` has checked, ready for `applyAgentEdits`. */
export type DecodedAgentEdits = readonly DecodedAgentEdit[] & { readonly [decoded]: true };

/**
 * Checks a model's `edits` (the `edits` of `agentEditSchema`) without
 * touching a presentation. Throws an `AgentEditError` with `index: null`
 * that lists every problem by its JSON Pointer. `touchedSlides` are the
 * part names of the slides the edits change, in order, each once.
 */
export function decodeAgentEdits(edits: unknown): {
  edits: DecodedAgentEdits;
  touchedSlides: readonly string[];
} {
  const errors = schemaErrors({ ...editsSchema, $defs: agentEditSchema.$defs! }, edits);
  if (errors.length > 0)
    throw new AgentEditError(
      null,
      `Invalid edits:\n${errors.map((error) => `- ${error}`).join('\n')}`,
    );
  const touched = new Set<string>();
  const out = (edits as readonly DecodedAgentEdit[]).map(({ tool, input }) => {
    const operation = OPERATIONS[tool]!;
    const target = input[operation.target] as string | { readonly slide: string };
    touched.add(typeof target === 'string' ? target : target.slide);
    return { tool, input: { ...input, ...operation.fixed } };
  });
  return { edits: Object.assign(out, { [decoded]: true as const }), touchedSlides: [...touched] };
}

/**
 * Applies `edits` to `presentation` in order, each reading references
 * against the presentation as the edits before it left it, and resolves
 * with each edit's result. When edit `i` fails it rejects with an
 * `AgentEditError` of `index: i`; the edits before it have been applied,
 * so discard the presentation.
 */
export async function applyAgentEdits(
  presentation: PresentationData,
  edits: DecodedAgentEdits,
): Promise<readonly JsonValue[]> {
  const host: ToolHost = {
    presentation: () => presentation,
    apply: async (_, edit) => edit(presentation),
  };
  const results: JsonValue[] = [];
  for (const [index, { tool, input }] of edits.entries()) {
    try {
      // Each edit reads what the previous one left, so they run one by one.
      results.push(await runTool(host, tool, input));
    } catch (error) {
      throw new AgentEditError(index, error instanceof Error ? error.message : String(error));
    }
  }
  return results;
}

/**
 * What a model reads before editing: for each slide (all of them, or the
 * ones `slides` names by part name, in presentation order) its part name,
 * 0-based index, title, whether it is hidden, its speaker notes and its
 * shapes as the editor's `listShapes` reports them.
 */
export function readForAgent(
  presentation: PresentationData,
  slides: readonly string[] | null,
): JsonValue {
  const all = pptx.getSlides(presentation);
  const wanted = slides && new Set(slides);
  if (wanted) {
    const known = new Set(all.map((slide) => pptx.getSlidePartName(slide)));
    const missing = [...wanted].filter((name) => !known.has(name));
    if (missing.length > 0) throw new Error(`There is no slide ${missing.join(', ')}.`);
  }
  return {
    slides: all.flatMap((slide, index) => {
      const partName = pptx.getSlidePartName(slide);
      if (wanted && !wanted.has(partName)) return [];
      return [
        {
          slide: partName,
          index,
          title: pptx.getSlideTitle(slide),
          hidden: pptx.isSlideHidden(slide),
          notes: pptx.getSlideNotes(slide),
          shapes: pptx.getSlideShapes(slide).map((shape) => shapeSummary(presentation, shape)),
        },
      ];
    }),
  };
}

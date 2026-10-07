// The editor's capabilities as tools a language model can call.
//
// Every mutating `@office-kit/pptx` export is a tool named after it, described
// by its TSDoc and typed by a JSON Schema that `manifest/generate.mjs` derives
// from its signature. Four read tools, written here, are how a model finds the
// slides, shapes, layouts and comments those schemas refer to.
//
// A model's input is untrusted: it is checked against the schema before
// anything runs, and the check fails with a message that names every problem
// by its path, so the model can correct the call. Nothing is coerced.
//
// This module is loaded on the first `tools()` or `run()`: the schemas are a
// large share of the editor's size and most embeddings never use them.

import * as pptx from '@office-kit/pptx';
import type {
  PresentationData,
  SlideCommentData,
  SlideData,
  SlideLayoutData,
  SlideShapeData,
} from '@office-kit/pptx';
import { generatedTools } from '../manifest/tools.generated.ts';
import { shapeRef } from './shape-ref.ts';

/** A JSON value: what a tool takes as input and gives back. */
export type JsonValue =
  | null
  | boolean
  | number
  | string
  | readonly JsonValue[]
  | { readonly [key: string]: JsonValue };

/** A JSON Schema (draft 2020-12) object. */
export type JsonSchema = { readonly [key: string]: JsonValue };

/**
 * A tool definition in the shape model APIs take: Anthropic's `tools` entries
 * as they are; for OpenAI, `{ type: 'function', function: { name, description,
 * parameters: input_schema } }`.
 */
export interface EditorTool {
  /** The `@office-kit/pptx` export it calls, or a read tool's name. */
  readonly name: string;
  readonly description: string;
  /** The input, as a self-contained JSON Schema (draft 2020-12) object. */
  readonly input_schema: JsonSchema & { readonly type: 'object' };
}

interface MutatingTool extends EditorTool {
  /** Positional arguments: an input property, or `null` for the presentation. */
  readonly args: readonly (string | null)[];
  /** The result's schema; references in it are turned back into refs. */
  readonly output: JsonSchema;
}

/** What `manifest/generate.mjs` writes to `tools.generated.ts`. */
export interface GeneratedTools {
  /** Capabilities that are not tools, with the reason. */
  readonly notTools: Readonly<Record<string, string>>;
  /** The schemas of the references tools take, by `$defs` name. */
  readonly refs: Readonly<Record<string, JsonSchema>>;
  readonly tools: readonly MutatingTool[];
}

interface ReadTool extends EditorTool {
  read(presentation: PresentationData, input: Record<string, unknown>): JsonValue;
}

// A dynamic view of the library, to call each tool's export by its name.
const library: Readonly<Record<string, unknown>> = pptx;

// --- References -----------------------------------------------------------------

const DEFS = '#/$defs/';

interface ShapeLocation {
  readonly slide: string;
  readonly shapeId: number;
}

function slideByPartName(presentation: PresentationData, partName: string): SlideData {
  const slide = pptx.findSlideByPartName(presentation, partName);
  if (!slide) throw new Error(`There is no slide ${partName}; listSlides gives the current ones.`);
  return slide;
}

function shapeAt(presentation: PresentationData, { slide, shapeId }: ShapeLocation) {
  const shape = pptx.findShapeById(slideByPartName(presentation, slide), shapeId);
  if (!shape)
    throw new Error(`There is no shape ${shapeId} on ${slide}; listShapes gives the current ones.`);
  return shape;
}

// How each `$defs` reference becomes the library object it names. The value
// has been checked against the reference's schema, which is what lets each
// function take it in the shape that schema describes.
const RESOLVE: Readonly<Record<string, (presentation: PresentationData, value: never) => unknown>> =
  {
    SlideRef: slideByPartName,
    ShapeRef: shapeAt,
    ChartRef(presentation, ref: ShapeLocation) {
      const shape = shapeAt(presentation, ref);
      const chart = pptx.getSlideCharts(pptx.getShapeSlide(shape)).find((c) => c.shape === shape);
      if (!chart) throw new Error(`Shape ${ref.shapeId} on ${ref.slide} is not a chart.`);
      return chart;
    },
    CellRef: (presentation, ref: ShapeLocation & { row: number; col: number }) =>
      pptx.getTableCell(shapeAt(presentation, ref), ref.row, ref.col),
    LayoutRef(presentation, partName: string) {
      const layout = pptx.findSlideLayoutByPartName(presentation, partName);
      if (!layout)
        throw new Error(`There is no layout ${partName}; listLayouts gives the current ones.`);
      return layout;
    },
    CommentRef(presentation, { slide, index }: { slide: string; index: number }) {
      const comment = pptx.getSlideComments(slideByPartName(presentation, slide))[index];
      if (!comment)
        throw new Error(`There is no comment ${index} on ${slide}; listComments gives them.`);
      return comment;
    },
    Bytes: (_, base64: string) => Uint8Array.from(atob(base64), (c) => c.charCodeAt(0)),
    DateTime: (_, text: string) => new Date(text),
  };

// How library objects in a result become references again.
const ENCODE: Readonly<
  Record<string, (presentation: PresentationData, value: never) => JsonValue>
> = {
  SlideRef: (_, slide: SlideData) => pptx.getSlidePartName(slide),
  ShapeRef: (presentation, shape: SlideShapeData) => ({ ...shapeRef(presentation, shape) }),
  LayoutRef: (_, layout: SlideLayoutData) => pptx.getSlideLayoutPartName(layout),
  CommentRef(_, comment: SlideCommentData) {
    const slide = pptx.getCommentSlide(comment);
    return {
      slide: pptx.getSlidePartName(slide),
      index: pptx.getSlideComments(slide).findIndex((c) => sameComment(c, comment)),
    };
  },
};

// `getSlideComments` builds new objects on every call, and a comment has no
// public id: it is the same one when it says the same thing from the same
// author at the same time.
function sameComment(a: SlideCommentData, b: SlideCommentData): boolean {
  return (
    pptx.getCommentText(a) === pptx.getCommentText(b) &&
    pptx.getCommentAuthor(a).name === pptx.getCommentAuthor(b).name &&
    pptx.getCommentDate(a) === pptx.getCommentDate(b)
  );
}

// --- Checking input ----------------------------------------------------------------

const MAX_LISTED_VALUES = 30;
const BASE64 = /^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/;
const DATE_TIME = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})$/i;

interface Decoding {
  readonly defs: Readonly<Record<string, JsonSchema>>;
  readonly errors: string[];
  /** Present when references are to be resolved, absent while only checking. */
  readonly presentation?: PresentationData;
}

const jsonType = (value: unknown): string =>
  value === null
    ? 'null'
    : Array.isArray(value)
      ? 'array'
      : typeof value === 'number' && Number.isInteger(value)
        ? 'integer'
        : typeof value;

const show = (value: unknown): string => {
  const text = JSON.stringify(value);
  return text.length > 60 ? `${text.slice(0, 57)}...` : text;
};

function listValues(values: readonly JsonValue[]): string {
  const shown = values.slice(0, MAX_LISTED_VALUES).map((v) => JSON.stringify(v));
  const more = values.length - shown.length;
  return `${shown.join(', ')}${more > 0 ? `, … (${more} more in the schema)` : ''}`;
}

function matchesType(value: unknown, type: string): boolean {
  const actual = jsonType(value);
  return actual === type || (type === 'number' && actual === 'integer');
}

/** `schema` with `$ref`s followed, except to the references tools resolve. */
function deref(schema: JsonSchema, defs: Readonly<Record<string, JsonSchema>>): JsonSchema {
  let current = schema;
  while (typeof current.$ref === 'string' && !(current.$ref.slice(DEFS.length) in RESOLVE))
    current = defs[current.$ref.slice(DEFS.length)]!;
  return current;
}

/** The alternatives of an `anyOf`, with nested ones (`Color | null`) spread out. */
function alternatives(
  schema: JsonSchema,
  defs: Readonly<Record<string, JsonSchema>>,
): JsonSchema[] {
  return (schema.anyOf as readonly JsonSchema[]).flatMap((branch) => {
    const target = deref(branch, defs);
    return Array.isArray(target.anyOf) ? alternatives(target, defs) : [branch];
  });
}

/** The literal values an object schema fixes its property `key` to, if it does. */
function literals(
  schema: JsonSchema,
  key: string,
  defs: Readonly<Record<string, JsonSchema>>,
): readonly JsonValue[] | undefined {
  const property = (schema.properties as Readonly<Record<string, JsonSchema>> | undefined)?.[key];
  if (!property) return undefined;
  const target = deref(property, defs);
  if ('const' in target) return [target.const!];
  return Array.isArray(target.enum) ? target.enum : undefined;
}

/**
 * The property of `value` that tells the object alternatives apart (`kind`):
 * one every alternative fixes to literals, not all to the same ones.
 */
function discriminator(
  objects: readonly JsonSchema[],
  value: Readonly<Record<string, unknown>>,
  defs: Readonly<Record<string, JsonSchema>>,
): string | undefined {
  return Object.keys(value).find((key) => {
    const sets = objects.map((schema) => literals(schema, key, defs));
    return (
      sets.every((set) => set !== undefined) &&
      new Set(sets.map((set) => JSON.stringify(set))).size > 1
    );
  });
}

// Says why `value` matched no alternative, as narrowly as it can: only the
// alternatives of the value's JSON type, and of those only the ones whose
// discriminator (`kind: 'bar'`) the value names. A model then reads the
// problems with the form it meant, not with every form there is.
function explainMismatch(
  branches: readonly JsonSchema[],
  tried: readonly (readonly string[])[],
  value: unknown,
  path: string,
  decoding: Decoding,
): undefined {
  const { defs, errors } = decoding;
  const at = path || '(input)';
  const targets = branches.map((branch) => deref(branch, defs));
  const ofType = (target: JsonSchema) => {
    const types = typesOf(target, defs);
    return types.length === 0 || types.some((type) => matchesType(value, type));
  };
  let candidates = targets.map((_, i) => i).filter((i) => ofType(targets[i]!));
  if (candidates.length === 0) {
    const types = [...new Set(targets.flatMap((target) => typesOf(target, defs)))];
    errors.push(`${at}: expected ${types.join(' or ')}, got ${jsonType(value)} ${show(value)}`);
    return undefined;
  }
  const object = jsonType(value) === 'object' ? (value as Record<string, unknown>) : undefined;
  const key =
    object &&
    discriminator(
      candidates.map((i) => targets[i]!),
      object,
      defs,
    );
  if (object && key !== undefined) {
    const given = object[key];
    const named = candidates.filter((i) =>
      literals(targets[i]!, key, defs)!.includes(given as JsonValue),
    );
    if (named.length === 0) {
      const allowed = new Set(candidates.flatMap((i) => literals(targets[i]!, key, defs)!));
      errors.push(`${path}/${key}: must be one of ${listValues([...allowed])}; got ${show(given)}`);
      return undefined;
    }
    candidates = named;
  }
  if (candidates.length === 1) {
    errors.push(...tried[candidates[0]!]!);
    return undefined;
  }
  const local = (error: string) =>
    error.startsWith(`${at}: `) ? error.slice(at.length + 2) : error;
  errors.push(
    `${at}: matches none of the ${candidates.length} forms it could take:\n${candidates
      .map((i, n) => `    ${n + 1}. ${tried[i]!.map(local).join('; ')}`)
      .join('\n')}`,
  );
  return undefined;
}

function typesOf(schema: JsonSchema, defs: Readonly<Record<string, JsonSchema>>): string[] {
  if (typeof schema.$ref === 'string') return typesOf(defs[schema.$ref.slice(DEFS.length)]!, defs);
  if (schema.type !== undefined)
    return (Array.isArray(schema.type) ? schema.type : [schema.type]) as string[];
  return [];
}

/**
 * `value` as the library takes it, checked against `schema`. Problems are
 * collected in `decoding.errors` with the JSON Pointer of the offending value;
 * the return value is meaningless when there are any.
 */
function decode(schema: JsonSchema, value: unknown, path: string, decoding: Decoding): unknown {
  const at = path || '(input)';
  const fail = (message: string) => {
    decoding.errors.push(`${at}: ${message}`);
    return undefined;
  };
  if (typeof schema.$ref === 'string') {
    const name = schema.$ref.slice(DEFS.length);
    const before = decoding.errors.length;
    const checked = decode(decoding.defs[name]!, value, path, decoding);
    const resolve = RESOLVE[name];
    if (decoding.errors.length > before || !resolve) return checked;
    if (name === 'Bytes' && !BASE64.test(value as string)) return fail('is not valid base64');
    if (
      name === 'DateTime' &&
      !(DATE_TIME.test(value as string) && !Number.isNaN(Date.parse(value as string)))
    )
      return fail('is not an RFC 3339 date-time such as "2026-10-08T09:30:00Z"');
    return decoding.presentation ? resolve(decoding.presentation, value as never) : value;
  }
  if (Array.isArray(schema.anyOf)) {
    const branches = alternatives(schema, decoding.defs);
    const tried = branches.map((branch) => {
      const trial: Decoding = { defs: decoding.defs, errors: [] };
      decode(branch, value, path, trial);
      return trial.errors;
    });
    const match = tried.findIndex((errors) => errors.length === 0);
    if (match >= 0) return decode(branches[match]!, value, path, decoding);
    return explainMismatch(branches, tried, value, path, decoding);
  }
  if (schema.type !== undefined) {
    const types = (Array.isArray(schema.type) ? schema.type : [schema.type]) as string[];
    if (!types.some((type) => matchesType(value, type)))
      return fail(`expected ${types.join(' or ')}, got ${jsonType(value)} ${show(value)}`);
  }
  if ('const' in schema && value !== schema.const)
    return fail(`must be ${JSON.stringify(schema.const)}, got ${show(value)}`);
  if (Array.isArray(schema.enum) && !schema.enum.includes(value as JsonValue))
    return fail(`must be one of ${listValues(schema.enum)}; got ${show(value)}`);
  if (typeof value === 'string' && typeof schema.pattern === 'string') {
    if (!new RegExp(schema.pattern, 'u').test(value))
      return fail(`${show(value)} does not match the pattern ${schema.pattern}`);
  }
  if (typeof value === 'number' && typeof schema.minimum === 'number' && value < schema.minimum)
    return fail(`must be at least ${schema.minimum}, got ${value}`);
  if (Array.isArray(value)) return decodeArray(schema, value, path, decoding);
  if (value !== null && typeof value === 'object')
    return decodeObject(schema, value as Record<string, unknown>, path, decoding);
  return value;
}

function decodeArray(
  schema: JsonSchema,
  value: readonly unknown[],
  path: string,
  decoding: Decoding,
): unknown[] {
  const at = path || '(input)';
  if (typeof schema.minItems === 'number' && value.length < schema.minItems)
    decoding.errors.push(`${at}: needs at least ${schema.minItems} items, got ${value.length}`);
  if (typeof schema.maxItems === 'number' && value.length > schema.maxItems)
    decoding.errors.push(`${at}: takes at most ${schema.maxItems} items, got ${value.length}`);
  const prefix = (schema.prefixItems ?? []) as readonly JsonSchema[];
  return value.map((item, i) => {
    const itemSchema = prefix[i] ?? (schema.items as JsonSchema | undefined);
    return itemSchema ? decode(itemSchema, item, `${path}/${i}`, decoding) : item;
  });
}

function decodeObject(
  schema: JsonSchema,
  value: Record<string, unknown>,
  path: string,
  decoding: Decoding,
): Record<string, unknown> {
  const at = path || '(input)';
  const properties = (schema.properties ?? {}) as Readonly<Record<string, JsonSchema>>;
  const missing = ((schema.required ?? []) as string[]).filter((name) => !(name in value));
  if (missing.length > 0)
    decoding.errors.push(
      `${at}: missing required ${missing.map((name) => JSON.stringify(name)).join(', ')}`,
    );
  const out: Record<string, unknown> = {};
  for (const [key, item] of Object.entries(value)) {
    const childPath = `${path}/${key.replace(/~/g, '~0').replace(/\//g, '~1')}`;
    const propertySchema = properties[key];
    if (propertySchema) out[key] = decode(propertySchema, item, childPath, decoding);
    else if (schema.additionalProperties === false) {
      const allowed = Object.keys(properties);
      decoding.errors.push(
        `${at}: unknown property ${JSON.stringify(key)}${
          allowed.length > 0 ? ` (allowed: ${allowed.join(', ')})` : ''
        }`,
      );
    } else if (schema.additionalProperties && typeof schema.additionalProperties === 'object')
      out[key] = decode(schema.additionalProperties as JsonSchema, item, childPath, decoding);
    else out[key] = item;
  }
  return out;
}

// --- Encoding results ------------------------------------------------------------

function encode(
  schema: JsonSchema,
  value: unknown,
  defs: Readonly<Record<string, JsonSchema>>,
  presentation: PresentationData,
): JsonValue {
  if (value === undefined || value === null) return null;
  if (typeof schema.$ref === 'string') {
    const name = schema.$ref.slice(DEFS.length);
    const toRef = ENCODE[name];
    return toRef
      ? toRef(presentation, value as never)
      : encode(defs[name]!, value, defs, presentation);
  }
  if (Array.isArray(schema.anyOf)) {
    const branch = (schema.anyOf as readonly JsonSchema[]).find((b) => accepts(b, value, defs));
    return branch ? encode(branch, value, defs, presentation) : (value as JsonValue);
  }
  if (Array.isArray(value))
    return value.map((item) =>
      encode((schema.items ?? {}) as JsonSchema, item, defs, presentation),
    );
  if (typeof value === 'object') {
    const properties = (schema.properties ?? {}) as Readonly<Record<string, JsonSchema>>;
    return Object.fromEntries(
      Object.entries(value).map(([key, item]) => [
        key,
        encode(properties[key] ?? {}, item, defs, presentation),
      ]),
    );
  }
  return value as JsonValue;
}

// Which `anyOf` branch a library value belongs to: by its JSON kind, with
// library objects going to the reference branch.
function accepts(
  schema: JsonSchema,
  value: unknown,
  defs: Readonly<Record<string, JsonSchema>>,
): boolean {
  if (typeof schema.$ref === 'string') {
    const name = schema.$ref.slice(DEFS.length);
    return name in ENCODE ? typeof value === 'object' : accepts(defs[name]!, value, defs);
  }
  const types = (Array.isArray(schema.type) ? schema.type : [schema.type]) as string[];
  return types.some((type) => matchesType(value, type));
}

// --- Read tools ------------------------------------------------------------------

const SLIDE_INPUT = {
  type: 'object',
  properties: { slide: { $ref: `${DEFS}SlideRef` } },
  required: ['slide'],
  additionalProperties: false,
  $defs: { SlideRef: generatedTools.refs.SlideRef! },
} as const satisfies EditorTool['input_schema'];

const NO_INPUT = {
  type: 'object',
  properties: {},
  additionalProperties: false,
} as const satisfies EditorTool['input_schema'];

function shapeSummary(presentation: PresentationData, shape: SlideShapeData): JsonValue {
  const bounds = pptx.getShapeBoundsResolved(presentation, shape);
  const summary: Record<string, JsonValue> = {
    shape: { ...shapeRef(presentation, shape) },
    kind: pptx.isTableShape(shape)
      ? 'table'
      : pptx.isChartShape(shape)
        ? 'chart'
        : pptx.getShapeKind(shape),
    placeholder: pptx.getShapePlaceholderType(shape),
    bounds: bounds && { x: bounds.x, y: bounds.y, w: bounds.w, h: bounds.h },
  };
  if (pptx.isTableShape(shape)) {
    const { rows, cols } = pptx.getTableDimensions(shape);
    summary.cells = Array.from({ length: rows }, (_, row) =>
      Array.from({ length: cols }, (_, col) =>
        pptx.getTableCellText(pptx.getTableCell(shape, row, col)),
      ),
    );
  } else if (pptx.hasShapeText(shape)) summary.text = pptx.getShapeText(shape);
  return summary;
}

const readTools: readonly ReadTool[] = [
  {
    name: 'listSlides',
    description:
      'Lists the slides in order: each slide\'s ref (its part name, which other tools take as "slide"), 0-based index, title, layout ref and whether it is hidden.',
    input_schema: NO_INPUT,
    read: (presentation) =>
      pptx.getSlides(presentation).map((slide, index) => {
        const layout = pptx.getSlideLayout(slide);
        return {
          slide: pptx.getSlidePartName(slide),
          index,
          title: pptx.getSlideTitle(slide),
          layout: layout && pptx.getSlideLayoutPartName(layout),
          hidden: pptx.isSlideHidden(slide),
        };
      }),
  },
  {
    name: 'listShapes',
    description: `Lists the shapes on a slide, back to front: each shape's ref (which other tools take as "shape", or with "row" and "col" added as a table "cell"), its kind, placeholder type, bounds and text (for a table, the text of every cell). Bounds are in ${'EMU: 914400 per inch, 360000 per cm, 12700 per point.'}`,
    input_schema: SLIDE_INPUT,
    read: (presentation, input) =>
      pptx
        .getSlideShapes(input.slide as SlideData)
        .map((shape) => shapeSummary(presentation, shape)),
  },
  {
    name: 'listLayouts',
    description:
      "Lists the slide layouts: each layout's ref (its part name, which other tools take as a layout), name, type and the part name of its slide master.",
    input_schema: NO_INPUT,
    read: (presentation) =>
      pptx.getSlideLayouts(presentation).map((layout) => ({
        layout: pptx.getSlideLayoutPartName(layout),
        name: pptx.getSlideLayoutName(layout),
        type: pptx.getSlideLayoutType(layout),
        master: pptx.getSlideMasterPartName(layout),
      })),
  },
  {
    name: 'listComments',
    description:
      'Lists the comments on a slide: each comment\'s ref (which other tools take as "comment"), author, text and status. Positions change when a comment is removed; list again after removing one.',
    input_schema: SLIDE_INPUT,
    read: (_, input) =>
      pptx.getSlideComments(input.slide as SlideData).map((comment, index) => ({
        comment: { slide: pptx.getSlidePartName(input.slide as SlideData), index },
        author: pptx.getCommentAuthor(comment).name,
        text: pptx.getCommentText(comment),
        status: pptx.getCommentStatus(comment),
      })),
  },
];

const byName: ReadonlyMap<string, MutatingTool | ReadTool> = new Map(
  [...generatedTools.tools, ...readTools].map((tool) => [tool.name, tool]),
);

/** Every tool, sorted by name. */
export const editorTools: readonly EditorTool[] = [...byName.values()]
  .map(({ name, description, input_schema }) => ({ name, description, input_schema }))
  .sort((a, b) => a.name.localeCompare(b.name));

// --- Running ------------------------------------------------------------------------

/** What `run` needs from the editor: the open presentation and `apply`. */
export interface ToolHost {
  presentation(): PresentationData;
  apply(label: string, edit: (presentation: PresentationData) => void): Promise<void>;
}

function defsOf(schema: JsonSchema): Readonly<Record<string, JsonSchema>> {
  return (schema.$defs ?? {}) as Readonly<Record<string, JsonSchema>>;
}

/**
 * Runs tool `name` with the model's `input`. Rejects, editing nothing, when
 * there is no such tool or the input does not match its schema; the message
 * lists each problem by its JSON Pointer, for the model to correct.
 */
export async function runTool(host: ToolHost, name: string, input: unknown): Promise<JsonValue> {
  const tool = byName.get(name);
  if (!tool)
    throw new Error(`There is no tool ${JSON.stringify(name)}. Call one of the tools listed.`);
  const defs = defsOf(tool.input_schema);
  const checking: Decoding = { defs, errors: [] };
  decode(tool.input_schema, input, '', checking);
  if (checking.errors.length > 0)
    throw new Error(
      `Invalid input for ${name}:\n${checking.errors.map((error) => `- ${error}`).join('\n')}`,
    );
  const resolve = (presentation: PresentationData) =>
    decode(tool.input_schema, input, '', { defs, errors: [], presentation }) as Record<
      string,
      unknown
    >;
  if ('read' in tool) {
    const presentation = host.presentation();
    return tool.read(presentation, resolve(presentation));
  }
  const call = library[name];
  // The coverage test keeps every generated tool a library function; a name
  // that is not one is a stale build, not bad input.
  if (typeof call !== 'function') throw new Error(`@office-kit/pptx has no function ${name}.`);
  let result: JsonValue = null;
  await host.apply(name, (presentation) => {
    const args = resolve(presentation);
    const returned: unknown = call(...tool.args.map((arg) => (arg ? args[arg] : presentation)));
    result = encode(tool.output, returned, defsOf(tool.output), presentation);
  });
  return result;
}

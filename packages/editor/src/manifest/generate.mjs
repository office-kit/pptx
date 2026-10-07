// Capability-manifest generator — the backbone of the coverage guarantee.
//
// Type-checks the @office-kit/pptx public API with the TypeScript compiler,
// enumerates every *mutating* export (the verbs a PowerPoint-style UI must
// expose as an operation), and emits two files:
//
// - `capabilities.generated.json`: one entry per capability with its operand,
//   parameter list and a heuristic category, for the editor's command UI.
// - `tools.generated.json`: the same capabilities as model-callable tools — the
//   export's TSDoc as the description and a JSON Schema (draft 2020-12) of its
//   parameters, derived from the checked types rather than from source text.
//
// The checker, not a regex over the source, is what finds the exports: an
// overloaded `export function` is as visible to it as an `export const` arrow,
// and a type spelled through aliases, `Partial<…>` or template literals
// resolves to what callers actually pass.
//
// The coverage test (`test/editor-capability-coverage.test.ts`) independently
// re-derives the mutating-export set from the compiled library and fails if it
// drifts from the manifest — so a new authoring function cannot be silently
// left out of the editor. Human refinements (labels, ribbon groups, richer
// dialog schemas) live in `overrides.ts`; they never remove entries.
//
// Run: `node packages/editor/src/manifest/generate.mjs` (`--check` only compares)
import { spawnSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import ts from 'typescript';

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = join(here, '..', '..', '..', '..');
const srcRoot = join(repoRoot, 'src');
const apiIndex = join(srcRoot, 'api', 'index.ts');

// Verb prefixes that denote a state-changing (authoring) operation. Kept in
// sync with `test/editor-capability-coverage.test.ts` — the two MUST agree.
const MUTATING_VERBS = [
  'add',
  'set',
  'clear',
  'replace',
  'remove',
  'insert',
  'duplicate',
  'bring',
  'send',
  'append',
  'group',
  'ungroup',
  'swap',
  'sort',
  'reverse',
  'rename',
  'move',
  'merge',
  'split',
  'import',
  'copy',
  'create',
  'translate',
  'touch',
  'increment',
  'compact',
  'reset',
  'update',
  'apply',
  'transform',
];

// A handful of mutating-verb exports are not slide-authoring operations the
// canvas UI drives; they still get a command (reachable via palette) but we
// tag them so the ribbon layer can skip them.
const NON_CANVAS = new Set([
  'compactPackage',
  'incrementRevision',
  'touchModified',
  'mergePresentations',
  'importSlide',
  'setCoreProperties',
  'setExtendedProperties',
  'setMediaPartBytes',
  'removeThumbnail',
  'setThumbnail',
]);

// Capabilities a model cannot call with JSON, and why. Every other capability
// becomes a tool; one whose parameters cannot be expressed fails generation
// until it is either expressible or listed here.
const NOT_TOOLS = {
  createPresentation: 'returns a new presentation instead of editing the open one',
  importSlide: 'copies a slide from another presentation, which a model cannot pass',
  mergePresentations: 'merges another presentation, which a model cannot pass',
  sortSlides: 'takes a comparison function',
};

function isMutating(name) {
  return MUTATING_VERBS.some(
    (v) =>
      name.startsWith(v) &&
      name.length > v.length &&
      name[v.length] === name[v.length].toUpperCase(),
  );
}

const config = ts.getParsedCommandLineOfConfigFile(
  join(repoRoot, 'tsconfig.json'),
  {},
  {
    ...ts.sys,
    onUnRecoverableConfigFileDiagnostic(diagnostic) {
      throw new Error(ts.flattenDiagnosticMessageText(diagnostic.messageText, '\n'));
    },
  },
);
const program = ts.createProgram([apiIndex], config.options);
const checker = program.getTypeChecker();
const apiModule = checker.getSymbolAtLocation(program.getSourceFile(apiIndex));

const OPERAND_BY_TYPE = {
  PresentationData: 'presentation',
  SlideData: 'slide',
  SlideShapeData: 'shape',
  TableCellData: 'cell',
};

// Heuristic category from the source file name + export name.
function categoryOf(name, file) {
  const f = file;
  if (/charts\.ts$/.test(f)) return 'chart';
  if (/tables\.ts$/.test(f)) return 'table';
  if (/transition/.test(f)) return 'transition';
  if (/animation/.test(f)) return 'animation';
  if (/comments\.ts$/.test(f)) return 'comment';
  if (/notes/.test(f)) return 'notes';
  if (/theme|color-map|features/.test(f)) return 'theme';
  if (/background/.test(f)) return 'slide-background';
  if (/sections/.test(f)) return 'section';
  if (
    /\bslide-(deck|query|size|title)\b/.test(f) ||
    (/Slide/.test(name) && /slides?\b/i.test(name))
  )
    return 'slide';
  if (/hyperlink/i.test(name)) return 'hyperlink';
  if (/image/i.test(name)) return 'image';
  if (/gradient|patternfill|nofill|fill/i.test(name)) return 'fill';
  if (/stroke|arrow|dash|cap|join|compound/i.test(name)) return 'stroke';
  if (/glow|shadow|effect|reflection/i.test(name)) return 'effect';
  if (/paragraph|bullet/i.test(name)) return 'paragraph';
  if (/run|text|font|anchor|autofit|wrap|margin|column|direction/i.test(name)) return 'text';
  if (/^set?Shape|Shape/.test(name)) return 'shape';
  if (/Slide/.test(name)) return 'slide';
  if (/Presentation|Core|Extended|Revision|Modified|Media|Thumbnail|Package/.test(name))
    return 'presentation';
  return 'misc';
}

// --- Command-UI parameter specs ---------------------------------------------

// Classify a TS type string into a UI param kind.
function kindOf(typeText, paramName) {
  const t = typeText.replace(/\s+/g, ' ').trim();
  const lname = (paramName || '').toLowerCase();
  if (/^Emu\b/.test(t) || /\bEmu$/.test(t)) return 'emu';
  if (lname === 'color' || /\bcolor\b/i.test(lname)) return 'color';
  if (/^number\b/.test(t)) return 'number';
  if (/^boolean\b/.test(t)) return 'boolean';
  if (/^string\b/.test(t)) return 'string';
  // string-literal union → enum
  const lits = t.match(/'[^']+'/g);
  if (
    lits &&
    new RegExp(`^(?:'[^']+'\\s*\\|?\\s*)+$`).test(t.replace(/\bnull\b|\bundefined\b/g, '').trim())
  ) {
    return 'enum';
  }
  return 'object';
}

// The string choices of a parameter whose type spells them through an alias
// (`TextCase`, `BuiltinTableStyleName | (string & {})`): `enum` with the
// literals when that is all it takes, `string` when any string goes.
function stringKind(type) {
  const members = (type.isUnion() ? type.types : [type]).filter(
    (t) => !(t.flags & (ts.TypeFlags.Null | ts.TypeFlags.Undefined)),
  );
  const isString = (t) =>
    (t.flags & (ts.TypeFlags.StringLike | ts.TypeFlags.TemplateLiteral)) !== 0 ||
    (t.isIntersection() && t.types.some((part) => part.flags & ts.TypeFlags.String));
  if (members.length === 0 || !members.every(isString)) return null;
  return members.every((t) => t.isStringLiteral())
    ? { kind: 'enum', enumValues: members.map((t) => t.value) }
    : { kind: 'string' };
}

function paramSpec(parameter) {
  const declaration = parameter.valueDeclaration;
  const name = ts.isIdentifier(declaration.name) ? declaration.name.text : 'options';
  const type = (declaration.type?.getText() ?? 'unknown').replace(/\s+/g, ' ');
  let kind = kindOf(type, name);
  let enumValues =
    kind === 'enum' ? (type.match(/'([^']+)'/g) || []).map((x) => x.slice(1, -1)) : undefined;
  if (kind === 'object') {
    const byType = stringKind(checker.getTypeOfSymbol(parameter));
    if (byType) ({ kind, enumValues } = byType);
  }
  const spec = { name, type, kind, optional: checker.isOptionalParameter(declaration) };
  if (declaration.initializer) spec.default = declaration.initializer.getText();
  if (enumValues) spec.enumValues = enumValues;
  return spec;
}

// --- Tool schemas ---------------------------------------------------------------

// Library objects a model names by reference. The schemas live in `$defs`; the
// editor resolves each against the open presentation when the tool runs, and
// turns returned objects back into the same references.
const EMU_NOTE = 'EMU: 914400 per inch, 360000 per cm, 12700 per point.';
const REF_DEFS = {
  SlideRef: {
    type: 'string',
    pattern: '^/ppt/slides/[^/]+\\.xml$',
    description:
      'A slide, by its part name as `listSlides` reports it, e.g. "/ppt/slides/slide1.xml".',
  },
  ShapeRef: {
    type: 'object',
    description:
      "A shape, by its slide and id, as `listShapes` or the user's selection reports it. `slideIndex` and `name` are informational.",
    properties: {
      slide: { $ref: '#/$defs/SlideRef' },
      shapeId: { type: 'integer', minimum: 0 },
      slideIndex: { type: 'integer', minimum: 0 },
      name: { type: 'string' },
    },
    required: ['slide', 'shapeId'],
    additionalProperties: false,
  },
  ChartRef: {
    $ref: '#/$defs/ShapeRef',
    description: 'A chart, by the ShapeRef of the graphic frame that holds it.',
  },
  CellRef: {
    type: 'object',
    description:
      "A table cell: the table's ShapeRef with the cell's 0-based `row` and `col` added.",
    properties: {
      slide: { $ref: '#/$defs/SlideRef' },
      shapeId: { type: 'integer', minimum: 0 },
      row: { type: 'integer', minimum: 0 },
      col: { type: 'integer', minimum: 0 },
      slideIndex: { type: 'integer', minimum: 0 },
      name: { type: 'string' },
    },
    required: ['slide', 'shapeId', 'row', 'col'],
    additionalProperties: false,
  },
  LayoutRef: {
    type: 'string',
    pattern: '^/ppt/slideLayouts/[^/]+\\.xml$',
    description:
      'A slide layout, by its part name as `listLayouts` reports it, e.g. "/ppt/slideLayouts/slideLayout1.xml".',
  },
  CommentRef: {
    type: 'object',
    description:
      'A comment, by its slide and its 0-based position in `listComments` for that slide.',
    properties: {
      slide: { $ref: '#/$defs/SlideRef' },
      index: { type: 'integer', minimum: 0 },
    },
    required: ['slide', 'index'],
    additionalProperties: false,
  },
  Bytes: { type: 'string', contentEncoding: 'base64', description: 'File bytes, base64-encoded.' },
  DateTime: { type: 'string', format: 'date-time', description: 'An RFC 3339 date-time.' },
};
const REF_BY_SYMBOL = {
  SlideData: 'SlideRef',
  SlideShapeData: 'ShapeRef',
  SlideChartData: 'ChartRef',
  TableCellData: 'CellRef',
  SlideLayoutData: 'LayoutRef',
  SlideCommentData: 'CommentRef',
  Uint8Array: 'Bytes',
  Date: 'DateTime',
};
// The references `core/agent-tools.ts` can turn a returned object back into.
const RETURNABLE_REFS = new Set(['SlideRef', 'ShapeRef', 'LayoutRef', 'CommentRef']);
// Defs that point at other defs: a tool's `$defs` must carry the closure.
const REF_DEPENDENCIES = {
  ShapeRef: ['SlideRef'],
  ChartRef: ['ShapeRef', 'SlideRef'],
  CellRef: ['SlideRef'],
  CommentRef: ['SlideRef'],
};

// Aliases whose TypeScript spelling is looser than what the library accepts.
// `Emu` is a branded `number`, but EMU are whole (ST_Coordinate is xsd:long).
// TypeScript cannot spell "six hex digits", so `HexColor` is `#${string}` and
// the library rejects other bodies at run time. The schema says what is
// accepted, so a model is told before the call instead of by an exception.
const ALIAS_SCHEMAS = {
  Emu: { type: 'integer', description: `A length in ${EMU_NOTE}` },
  HexColor: {
    type: 'string',
    pattern: '^#(?:[0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$',
    description: 'An sRGB color, "#RRGGBB" or "#RGB".',
  },
};

class Unsupported extends Error {}

function docOf(symbol) {
  return ts.displayPartsToString(symbol.getDocumentationComment(checker)).trim();
}

function withDescription(schema, description) {
  if (!description) return schema;
  return {
    ...schema,
    description: schema.description ? `${description}\n\n${schema.description}` : description,
  };
}

const isUndefined = (type) => (type.flags & ts.TypeFlags.Undefined) !== 0;
const isNever = (type) => (type.flags & ts.TypeFlags.Never) !== 0;

function symbolName(type) {
  return (type.aliasSymbol ?? type.getSymbol())?.getName();
}

// The public type aliases, by the type they declare. A union is flattened
// where it is used (`Color | { color?: Color }` is one union of literals,
// templates and an object), so a member list is matched against these to
// find the alias again, and a template literal is recognised by identity.
const PUBLIC_ALIASES = new Map();
for (const exported of checker.getExportsOfModule(apiModule)) {
  const symbol =
    exported.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(exported) : exported;
  if (symbol.flags & ts.SymbolFlags.TypeAlias)
    PUBLIC_ALIASES.set(checker.getDeclaredTypeOfSymbol(symbol), exported.getName());
}
// Largest first, so `Color` is taken whole before `SchemeColorToken` inside it.
const PUBLIC_UNION_ALIASES = [...PUBLIC_ALIASES.keys()]
  .filter((type) => type.isUnion())
  .sort((a, b) => b.types.length - a.types.length);

function aliasName(type) {
  return (
    PUBLIC_ALIASES.get(type) ??
    (type.aliasTypeArguments?.length ? undefined : type.aliasSymbol?.getName())
  );
}

// The name a type is hoisted into `$defs` under: a type alias or interface.
// `ChartSpec` repeats its axis, series and text types dozens of times, and
// inlining them made one tool's schema a megabyte. `Partial<TableCellBorder>`
// is named for what it is partial of.
function definitionName(type) {
  const alias = aliasName(type);
  if (alias) return alias;
  if (type.aliasSymbol) {
    const [argument] = type.aliasTypeArguments ?? [];
    const inner = argument && definitionName(argument);
    return ['Partial', 'Readonly'].includes(type.aliasSymbol.getName()) && inner
      ? `${type.aliasSymbol.getName()}${inner}`
      : null;
  }
  const symbol = type.getSymbol();
  if (!symbol || !(symbol.flags & ts.SymbolFlags.Interface)) return null;
  return checker.getTypeArguments(type)?.length ? null : symbol.getName();
}

/** Per-tool state: the `$defs` built so far and the type each name stands for. */
function schemaContext() {
  return { defs: new Map(), names: new Map(), taken: new Set() };
}

function useRef(ctx, ref) {
  ctx.defs.set(ref, REF_DEFS[ref]);
  for (const dependency of REF_DEPENDENCIES[ref] ?? [])
    ctx.defs.set(dependency, REF_DEFS[dependency]);
  return { $ref: `#/$defs/${ref}` };
}

// The schema for `type`, hoisting named types into `ctx.defs`.
function schemaOf(type, ctx, where) {
  if (type.flags & ts.TypeFlags.Object) {
    const ref = REF_BY_SYMBOL[symbolName(type)];
    if (ref) return useRef(ctx, ref);
  }
  const name =
    type.flags &
    (ts.TypeFlags.Object |
      ts.TypeFlags.Union |
      ts.TypeFlags.Intersection |
      ts.TypeFlags.TemplateLiteral)
      ? definitionName(type)
      : null;
  if (!name) return buildSchema(type, ctx, where);
  let defName = ctx.names.get(type);
  if (!defName) {
    defName = name;
    for (let n = 2; ctx.taken.has(defName) || defName in REF_DEFS; n++) defName = `${name}${n}`;
    ctx.taken.add(defName);
    ctx.names.set(type, defName);
    // Registered before it is built, so a recursive type refers to itself.
    ctx.defs.set(defName, ALIAS_SCHEMAS[name] ?? buildSchema(type, ctx, where));
  }
  return { $ref: `#/$defs/${defName}` };
}

function buildSchema(type, ctx, where) {
  if (type.isUnion()) return unionSchema(type, ctx, where);
  if (type.flags & ts.TypeFlags.String) return { type: 'string' };
  if (type.flags & ts.TypeFlags.Number) return { type: 'number' };
  if (type.flags & ts.TypeFlags.Boolean) return { type: 'boolean' };
  if (type.flags & ts.TypeFlags.Null) return { type: 'null' };
  if (type.isStringLiteral()) return { type: 'string', const: type.value };
  if (type.isNumberLiteral()) return { type: 'number', const: type.value };
  if (type.flags & ts.TypeFlags.BooleanLiteral)
    return { type: 'boolean', const: checker.typeToString(type) === 'true' };
  if (type.flags & ts.TypeFlags.TemplateLiteral) return templateSchema(type, where);
  if (type.isIntersection()) {
    // A branded primitive (`number & { [brand] }`) is that primitive.
    const primitive = type.types.find((t) => t.flags & (ts.TypeFlags.Number | ts.TypeFlags.String));
    if (primitive) return schemaOf(primitive, ctx, where);
    return objectSchema(type, ctx, where);
  }
  if (type.flags & ts.TypeFlags.Object) {
    if (checker.isTupleType(type)) {
      const items = checker.getTypeArguments(type);
      return {
        type: 'array',
        prefixItems: items.map((t, i) => schemaOf(t, ctx, `${where}[${i}]`)),
        minItems: items.length,
        maxItems: items.length,
      };
    }
    if (checker.isArrayType(type)) {
      const [item] = checker.getTypeArguments(type);
      return { type: 'array', items: schemaOf(item, ctx, `${where}[]`) };
    }
    if (type.getCallSignatures().length > 0) throw new Unsupported(`${where} is a function`);
    if (symbolName(type) === 'RegExp') throw new Unsupported(`${where} is a RegExp`);
    return objectSchema(type, ctx, where);
  }
  throw new Unsupported(`${where} has type ${checker.typeToString(type)}`);
}

function unionSchema(union, ctx, where) {
  // `string | RegExp` search patterns: a model passes the string form.
  let kept = union.types.filter(
    (t) => !isUndefined(t) && !isNever(t) && symbolName(t) !== 'RegExp',
  );
  if (kept.length === 0) throw new Unsupported(`${where} has no JSON form`);
  const branches = [];
  for (const alias of PUBLIC_UNION_ALIASES) {
    if (alias === union || !alias.types.every((t) => kept.includes(t))) continue;
    branches.push(schemaOf(alias, ctx, where));
    kept = kept.filter((t) => !alias.types.includes(t));
  }
  const isLiteral = (t) =>
    t.isStringLiteral() || t.isNumberLiteral() || (t.flags & ts.TypeFlags.BooleanLiteral) !== 0;
  const strings = kept.filter((t) => t.isStringLiteral()).map((t) => t.value);
  const numbers = kept
    .filter((t) => t.isNumberLiteral())
    .map((t) => t.value)
    .sort((a, b) => a - b);
  const booleans = kept.filter((t) => t.flags & ts.TypeFlags.BooleanLiteral);
  if (strings.length > 0)
    branches.push(
      strings.length === 1
        ? { type: 'string', const: strings[0] }
        : { type: 'string', enum: strings },
    );
  if (numbers.length > 0) {
    const type = numbers.every(Number.isInteger) ? 'integer' : 'number';
    branches.push(numbers.length === 1 ? { type, const: numbers[0] } : { type, enum: numbers });
  }
  // `boolean` is `true | false` to the checker.
  if (booleans.length === 2) branches.push({ type: 'boolean' });
  else for (const t of booleans) branches.push(buildSchema(t, ctx, where));
  for (const t of kept.filter((t) => !isLiteral(t))) branches.push(schemaOf(t, ctx, where));
  if (branches.length === 1) return branches[0];
  // `string | null` reads better as a type list than as alternatives.
  if (branches.every((b) => Object.keys(b).length === 1 && typeof b.type === 'string'))
    return { type: branches.map((b) => b.type) };
  return { anyOf: branches };
}

function escapeRegExp(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function templateSchema(type, where) {
  let pattern = `^${escapeRegExp(type.texts[0])}`;
  type.types.forEach((part, i) => {
    if (part.flags & ts.TypeFlags.String) pattern += '.*';
    else if (part.flags & ts.TypeFlags.Number) pattern += '-?\\d+(?:\\.\\d+)?';
    else throw new Unsupported(`${where} is a template literal over ${checker.typeToString(part)}`);
    pattern += escapeRegExp(type.texts[i + 1]);
  });
  return { type: 'string', pattern: `${pattern}$` };
}

function withUnit(schema, name) {
  return /Emu$/.test(name) && !schema.description?.includes('EMU')
    ? withDescription(schema, `In ${EMU_NOTE}`)
    : schema;
}

function objectSchema(type, ctx, where) {
  const properties = {};
  const required = [];
  for (const property of checker.getPropertiesOfType(type)) {
    const name = property.getName();
    // Symbol-keyed internals are not part of what a caller writes.
    if (name.startsWith('__@')) continue;
    const propertyType = checker.getTypeOfSymbol(property);
    // `paragraphEnd?: never` (read back as `undefined`) marks a property that
    // must be absent; `additionalProperties: false` says so.
    if (isNever(propertyType) || isUndefined(propertyType)) continue;
    const optional =
      (property.flags & ts.SymbolFlags.Optional) !== 0 ||
      (propertyType.isUnion() && propertyType.types.some(isUndefined));
    const schema = schemaOf(propertyType, ctx, `${where}.${name}`);
    properties[name] = withUnit(withDescription(schema, docOf(property)), name);
    if (!optional) required.push(name);
  }
  const index = checker
    .getIndexInfosOfType(type)
    .find((info) => info.keyType.flags & ts.TypeFlags.String);
  const schema = { type: 'object' };
  if (Object.keys(properties).length > 0) schema.properties = properties;
  if (required.length > 0) schema.required = required;
  schema.additionalProperties = index ? schemaOf(index.type, ctx, `${where}[key]`) : false;
  return schema;
}

function refsIn(value, out = []) {
  if (Array.isArray(value)) for (const item of value) refsIn(item, out);
  else if (value && typeof value === 'object')
    for (const [key, child] of Object.entries(value)) {
      if (key === '$ref') out.push(child.slice('#/$defs/'.length));
      else refsIn(child, out);
    }
  return out;
}

// Inlines each definition used once (and not by itself), so `$defs` holds only
// what is shared or recursive and the schema reads top-down where it can.
function finish(root, ctx, keep) {
  const defs = Object.fromEntries(ctx.defs);
  const counts = new Map();
  for (const name of refsIn([root, ...Object.values(defs)]))
    counts.set(name, (counts.get(name) ?? 0) + 1);
  const inline = new Set(
    Object.keys(defs).filter(
      (name) => !keep(name) && counts.get(name) === 1 && !refsIn(defs[name]).includes(name),
    ),
  );
  const expand = (value) => {
    if (Array.isArray(value)) return value.map(expand);
    if (!value || typeof value !== 'object') return value;
    const name = typeof value.$ref === 'string' ? value.$ref.slice('#/$defs/'.length) : null;
    if (name && inline.has(name)) {
      const { $ref: _, ...siblings } = value;
      const body = expand(defs[name]);
      // A `$ref` with a description of its own: keep both.
      return siblings.description ? withDescription(body, siblings.description) : body;
    }
    return Object.fromEntries(Object.entries(value).map(([k, v]) => [k, expand(v)]));
  };
  const result = expand(root);
  const kept = Object.fromEntries(
    Object.keys(defs)
      .filter((name) => !inline.has(name))
      .map((name) => [name, expand(defs[name])]),
  );
  shareRepeats(result, kept);
  const names = Object.keys(kept).sort();
  if (names.length > 0) result.$defs = Object.fromEntries(names.map((n) => [n, kept[n]]));
  return result;
}

// Subschemas smaller than this stay inline even when repeated: a `$ref` and a
// definition name cost about as much as they save.
const MIN_SHARED_SCHEMA_LENGTH = 60;

// Moves repeated anonymous subschemas into `defs`. The chart spec is a union of
// variants that each carry the same legend, axis and title properties through
// intersections; they have no type name to hoist by, so they are found by
// content, the largest saving first.
function shareRepeats(root, defs) {
  for (;;) {
    const seen = new Map();
    const visit = (value, key) => {
      if (!value || typeof value !== 'object') return;
      if (!Array.isArray(value)) {
        const json = JSON.stringify(value);
        if (json.length >= MIN_SHARED_SCHEMA_LENGTH) {
          const entry = seen.get(json) ?? { count: 0, keys: new Set() };
          entry.count++;
          if (key) entry.keys.add(key);
          seen.set(json, entry);
        }
      }
      for (const [childKey, child] of Object.entries(value)) {
        if (childKey !== 'properties' || Array.isArray(value)) visit(child, null);
        else for (const [name, property] of Object.entries(child)) visit(property, name);
      }
    };
    visit(root, null);
    for (const body of Object.values(defs)) visit(body, null);
    let best;
    for (const [json, entry] of seen) {
      const saving = (entry.count - 1) * json.length;
      if (entry.count > 1 && (!best || saving > best.saving)) best = { json, saving, ...entry };
    }
    if (!best) return;
    let name = Object.keys(defs).find((n) => JSON.stringify(defs[n]) === best.json);
    if (!name) {
      const schema = JSON.parse(best.json);
      const base = sharedName(schema, best.keys);
      name = base;
      for (let n = 2; name in defs || name in REF_DEFS; n++) name = `${base}${n}`;
      defs[name] = schema;
    }
    const ref = { $ref: `#/$defs/${name}` };
    const replace = (value) => {
      if (!value || typeof value !== 'object') return;
      for (const [key, child] of Object.entries(value)) {
        if (child && typeof child === 'object' && JSON.stringify(child) === best.json)
          value[key] = { ...ref };
        else replace(child);
      }
    };
    replace(root);
    for (const [defName, body] of Object.entries(defs)) if (defName !== name) replace(body);
  }
}

// A name for a shared schema: the property that holds it when that is always
// the same one, else what it is made of (`PartialTableCellBorderOrNull`).
function sharedName(schema, keys) {
  const pascal = (text) => text.charAt(0).toUpperCase() + text.slice(1);
  const part = (branch) =>
    branch.$ref?.slice('#/$defs/'.length) ??
    (typeof branch.type === 'string' ? pascal(branch.type) : null);
  const parts = schema.anyOf?.map(part);
  if (keys.size !== 1 && parts?.every(Boolean)) return parts.join('Or');
  const [key] = keys;
  return key ? pascal(key.replace(/[^A-Za-z0-9]/g, '')) : 'Shared';
}

// `@param` text by parameter name, from the declaration's JSDoc.
function paramDocs(signature) {
  const docs = {};
  for (const tag of signature.getJsDocTags()) {
    if (tag.name !== 'param' || !tag.text) continue;
    const name = tag.text.find((part) => part.kind === 'parameterName')?.text;
    const text = tag.text
      .filter((part) => part.kind === 'text')
      .map((part) => part.text)
      .join('')
      .replace(/^\s*-\s*/, '')
      .trim();
    if (name && text) docs[name] = text;
  }
  return docs;
}

function toolOf(name, symbol, signature) {
  const ctx = schemaContext();
  const docs = paramDocs(signature);
  const properties = {};
  const required = [];
  // Positional arguments: a parameter name, or `null` for the open presentation.
  const args = [];
  signature.getParameters().forEach((parameter, index) => {
    const declaration = parameter.valueDeclaration;
    const type = checker.getTypeOfSymbol(parameter);
    if (symbolName(type) === 'PresentationData') {
      if (index !== 0) throw new Unsupported(`${parameter.getName()} is another presentation`);
      args.push(null);
      return;
    }
    const paramName = ts.isIdentifier(declaration.name) ? declaration.name.text : 'options';
    const schema = schemaOf(type, ctx, paramName);
    properties[paramName] = withUnit(withDescription(schema, docs[paramName]), paramName);
    if (!checker.isOptionalParameter(declaration)) required.push(paramName);
    args.push(paramName);
  });
  const input = { type: 'object', properties };
  if (required.length > 0) input.required = required;
  input.additionalProperties = false;
  // The result is turned back into references; only their names matter there.
  const outputCtx = schemaContext();
  const returnType = signature.getReturnType();
  const output =
    returnType.flags & ts.TypeFlags.Void
      ? { type: 'null' }
      : schemaOf(returnType, outputCtx, 'result');
  const unreturnable = [...outputCtx.defs.keys()].filter(
    (def) => def in REF_DEFS && !RETURNABLE_REFS.has(def),
  );
  if (unreturnable.length > 0)
    throw new Unsupported(`returns ${unreturnable.join(', ')}, which the editor cannot report`);
  // A model chooses a tool by its description; the export's TSDoc is it.
  const description = docOf(symbol);
  if (!description) throw new Unsupported('has no TSDoc to describe it with');
  return {
    name,
    description,
    args,
    input_schema: finish(input, ctx, (def) => def in REF_DEFS),
    output: finish(output, outputCtx, (def) => def in REF_DEFS),
  };
}

// --- Generation -------------------------------------------------------------------

const capabilities = [];
const tools = [];
const failures = [];
for (const exported of checker.getExportsOfModule(apiModule)) {
  const name = exported.getName();
  if (!isMutating(name)) continue;
  const symbol =
    exported.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(exported) : exported;
  if (!(symbol.flags & (ts.SymbolFlags.Function | ts.SymbolFlags.Variable))) continue;
  const signatures = checker.getTypeOfSymbol(symbol).getCallSignatures();
  if (signatures.length === 0) continue;
  // An overloaded export is described by its first overload, the general form.
  const [signature] = signatures;
  const declaration = signature.getDeclaration();
  const file = relative(srcRoot, declaration.getSourceFile().fileName).replace(/\\/g, '/');
  const allParams = signature.getParameters().map(paramSpec);
  // The operand is the leading parameter *only* when its type is one of the
  // domain objects (PresentationData / SlideData / SlideShapeData /
  // TableCellData). Factories like `createPresentation(options)` take no
  // operand — their first parameter is a real user argument.
  const firstType = allParams.length ? allParams[0].type.split(/[<\s]/)[0] : '';
  const operandFromFirst = OPERAND_BY_TYPE[firstType];
  const takesOperand = Boolean(operandFromFirst);
  capabilities.push({
    id: name,
    operand: operandFromFirst ?? 'presentation',
    takesOperand,
    category: categoryOf(name, file),
    file,
    returns: checker.typeToString(signature.getReturnType()),
    canvas: !NON_CANVAS.has(name),
    // The user-facing argument list: the registry and forms never ask for the
    // object the user already has selected.
    params: takesOperand ? allParams.slice(1) : allParams,
  });
  if (name in NOT_TOOLS) continue;
  try {
    tools.push(toolOf(name, symbol, signature));
  } catch (error) {
    if (!(error instanceof Unsupported)) throw error;
    failures.push(`${name}: ${error.message}`);
  }
}
for (const name of Object.keys(NOT_TOOLS)) {
  if (!capabilities.some((c) => c.id === name))
    failures.push(`${name}: listed in NOT_TOOLS but not a capability`);
}
if (failures.length > 0) {
  throw new Error(
    `These capabilities cannot be described as tools; make their types JSON-expressible or list them in NOT_TOOLS:\n  ${failures.join('\n  ')}`,
  );
}
capabilities.sort((a, b) => a.id.localeCompare(b.id));
tools.sort((a, b) => a.name.localeCompare(b.name));

// `--check` compares instead of writing, for the test that keeps the committed
// files in step with the library: hand-editing them is how they drifted before.
const check = process.argv.includes('--check');
const stale = [];

function write(file, text) {
  const path = join(here, file);
  // Formatted as `format:check` wants it, so regenerating never fails that.
  const formatted = spawnSync(
    join(repoRoot, 'node_modules', '.bin', 'oxfmt'),
    [`--stdin-filepath=${path}`],
    { input: text, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 },
  );
  if (formatted.status !== 0)
    throw new Error(`oxfmt failed on ${path} (exit ${formatted.status}): ${formatted.stderr}`);
  if (!check) writeFileSync(path, formatted.stdout);
  else if (readFileSync(path, 'utf8') !== formatted.stdout) stale.push(file);
}

write(
  'capabilities.generated.json',
  JSON.stringify({
    generatedFrom: 'src/api/index.ts',
    count: capabilities.length,
    capabilities,
  }),
);
// TypeScript rather than JSON so that the compiler checks the schemas against
// the types `core/agent-tools.ts` reads them as.
write(
  'tools.generated.ts',
  `// AUTO-GENERATED by generate.mjs from the @office-kit/pptx types. Do not edit.
import type { GeneratedTools } from '../core/agent-tools.ts';

export const generatedTools: GeneratedTools = ${JSON.stringify({
    notTools: NOT_TOOLS,
    refs: REF_DEFS,
    tools,
  })};
`,
);
if (stale.length > 0) {
  console.error(
    `${stale.join(' and ')} differ from the library's API. Run \`node packages/editor/src/manifest/generate.mjs\`.`,
  );
  process.exitCode = 1;
} else if (!check) {
  console.log(`Wrote ${capabilities.length} capabilities and ${tools.length} tools.`);
}

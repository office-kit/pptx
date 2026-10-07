import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import test from 'node:test';
import { build, transform } from 'esbuild';
import { compileModule } from 'svelte/compiler';
import capabilities from '../src/manifest/capabilities.generated.json' with { type: 'json' };

// The tools run against the same library copy the test reads back with.
const result = await build({
  stdin: {
    contents: `export { editorTools, runTool } from './src/core/agent-tools.ts';
      export { generatedTools } from './src/manifest/tools.generated.ts';
      export { EditorDocument } from './src/core/document.svelte.ts';
      export * as pptx from '@office-kit/pptx';`,
    resolveDir: fileURLToPath(new URL('..', import.meta.url)),
  },
  bundle: true,
  write: false,
  platform: 'node',
  format: 'esm',
  plugins: [
    {
      name: 'editor-runes',
      setup(builder) {
        builder.onLoad({ filter: /\.svelte\.ts$/ }, async ({ path }) => {
          const source = await transform(await readFile(path, 'utf8'), { loader: 'ts' });
          return {
            contents: compileModule(source.code, { filename: path, generate: 'client' }).js.code,
          };
        });
      },
    },
  ],
});
const { editorTools, runTool, generatedTools, EditorDocument, pptx } = await import(
  `data:text/javascript;base64,${Buffer.from(result.outputFiles[0].text).toString('base64')}`
);

const READ_TOOLS = ['listComments', 'listLayouts', 'listShapes', 'listSlides'];

const PNG_1X1 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

/** A host like `mountEditor`'s: each edit one undo step named after the tool. */
function host(doc) {
  return {
    presentation: () => doc.pres,
    async apply(label, edit) {
      try {
        doc.transact(`Agent: ${label}`, () => edit(doc.pres), 'agent');
      } catch (error) {
        await doc.rollback;
        throw error;
      }
    },
  };
}

function titleRef(doc) {
  const slide = pptx.getSlides(doc.pres)[0];
  const title = pptx.getSlideShapes(slide)[0];
  return { slide: pptx.getSlidePartName(slide), shapeId: pptx.getShapeId(title) };
}

// --- Schema validity --------------------------------------------------------------

// The draft 2020-12 keywords the generator uses. The check below is the
// meta-schema's rules for exactly these, so a schema that passes is valid.
const KEYWORDS = new Set([
  '$defs',
  '$ref',
  'additionalProperties',
  'anyOf',
  'const',
  'contentEncoding',
  'description',
  'enum',
  'format',
  'items',
  'maxItems',
  'minItems',
  'minimum',
  'pattern',
  'prefixItems',
  'properties',
  'required',
  'type',
]);
const TYPES = new Set(['array', 'boolean', 'integer', 'null', 'number', 'object', 'string']);

function schemaProblems(schema, root, path, problems) {
  const fail = (message) => problems.push(`${path}: ${message}`);
  if (schema === null || typeof schema !== 'object' || Array.isArray(schema))
    return fail('a schema must be an object');
  for (const [key, value] of Object.entries(schema)) {
    if (!KEYWORDS.has(key)) fail(`unknown keyword ${key}`);
    const at = `${path}/${key}`;
    switch (key) {
      case 'type': {
        const types = Array.isArray(value) ? value : [value];
        if (types.length === 0 || new Set(types).size !== types.length)
          fail('type must be a non-empty list of distinct types');
        for (const type of types) if (!TYPES.has(type)) fail(`unknown type ${type}`);
        break;
      }
      case 'enum':
        if (!Array.isArray(value) || value.length === 0) fail('enum must be a non-empty array');
        else if (new Set(value.map((v) => JSON.stringify(v))).size !== value.length)
          fail('enum values must be distinct');
        break;
      case 'required':
        if (!Array.isArray(value) || value.some((v) => typeof v !== 'string'))
          fail('required must be an array of strings');
        else if (new Set(value).size !== value.length) fail('required must be distinct');
        else
          for (const name of value)
            if (!(name in (schema.properties ?? {}))) fail(`required ${name} is not a property`);
        break;
      case 'properties':
      case '$defs':
        if (path !== '' && key === '$defs') fail('$defs belongs at the root');
        for (const [name, child] of Object.entries(value))
          schemaProblems(child, root, `${at}/${name}`, problems);
        break;
      case 'items':
      case 'additionalProperties':
        if (typeof value !== 'boolean') schemaProblems(value, root, at, problems);
        break;
      case 'anyOf':
      case 'prefixItems':
        if (!Array.isArray(value) || value.length === 0) fail(`${key} must be a non-empty array`);
        else value.forEach((child, i) => schemaProblems(child, root, `${at}/${i}`, problems));
        break;
      case '$ref':
        if (!/^#\/\$defs\/[A-Za-z0-9]+$/.test(value) || !root.$defs?.[value.slice(8)])
          fail(`$ref ${value} does not resolve`);
        break;
      case 'pattern':
        try {
          new RegExp(value, 'u');
        } catch {
          fail(`pattern ${value} is not a valid regular expression`);
        }
        break;
      case 'minItems':
      case 'maxItems':
      case 'minimum':
        if (typeof value !== 'number' || (key !== 'minimum' && !Number.isInteger(value)))
          fail(`${key} must be a number`);
        break;
      case 'description':
      case 'contentEncoding':
      case 'format':
        if (typeof value !== 'string' || value === '') fail(`${key} must be a non-empty string`);
        break;
    }
  }
}

/** `$defs` entries nothing refers to: dead weight in a model's context. */
function unusedDefs(schema) {
  const refs = new Set(JSON.stringify(schema).match(/#\/\$defs\/[A-Za-z0-9]+/g));
  return Object.keys(schema.$defs ?? {}).filter((name) => !refs.has(`#/$defs/${name}`));
}

test('every tool schema is a valid, self-contained JSON Schema object', () => {
  const problems = [];
  for (const tool of editorTools) {
    assert.match(tool.name, /^[a-zA-Z0-9_-]{1,64}$/, 'model APIs limit tool names to this');
    assert.ok(tool.description.trim(), `${tool.name} has a description`);
    assert.equal(tool.input_schema.type, 'object', tool.name);
    assert.deepEqual(Object.keys(tool).sort(), ['description', 'input_schema', 'name']);
    const found = [];
    schemaProblems(tool.input_schema, tool.input_schema, '', found);
    for (const name of unusedDefs(tool.input_schema)) found.push(`unused $defs/${name}`);
    problems.push(...found.map((problem) => `${tool.name}${problem}`));
  }
  assert.deepEqual(problems, []);
});

test('tools are every capability that a model can call, plus the read tools, sorted', () => {
  const names = editorTools.map((tool) => tool.name);
  assert.deepEqual(
    names,
    [...names].sort((a, b) => a.localeCompare(b)),
  );
  assert.equal(new Set(names).size, names.length);
  const expected = capabilities.capabilities
    .map((c) => c.id)
    .filter((id) => !(id in generatedTools.notTools));
  assert.deepEqual(
    names.filter((name) => !READ_TOOLS.includes(name)),
    expected.sort((a, b) => a.localeCompare(b)),
  );
  for (const name of expected) assert.equal(typeof pptx[name], 'function', name);
  for (const name of READ_TOOLS) assert.equal(pptx[name], undefined, `${name} is the editor's own`);
});

test('schemas say what the signatures take', () => {
  const schema = (name) => editorTools.find((tool) => tool.name === name).input_schema;
  const { properties, required, $defs } = schema('setShapePosition');
  assert.deepEqual(required, ['shape', 'x', 'y']);
  assert.deepEqual(properties.x, { $ref: '#/$defs/Emu' });
  assert.equal($defs.Emu.type, 'integer');
  assert.match($defs.Emu.description, /914400 per inch/);
  // The presentation is the open one, not an input.
  assert.deepEqual(Object.keys(schema('addTitleSlide').properties), ['title']);
  // An optional parameter, and a `never` property that must stay absent.
  const format = schema('setShapeTextFormat');
  assert.deepEqual(format.required, ['shape', 'format']);
  const strokeDash = schema('setShapeStrokeDash').properties.dash;
  assert.ok(strokeDash.enum.includes('sysDash'));
  // `HexColor` is `#${string}` to TypeScript; the schema says what is accepted.
  assert.equal(
    schema('setShapeFill').$defs.Color.anyOf.at(-1).pattern,
    '^#(?:[0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$',
  );
  assert.deepEqual(schema('setHandoutSlidesPerPage').properties.slidesPerPage.anyOf, [
    { type: 'string', const: 'outline' },
    { type: 'integer', enum: [1, 2, 3, 4, 6, 9] },
  ]);
});

// --- Running ---------------------------------------------------------------------

test('runs an edit as one undo step named after the tool', async () => {
  const doc = new EditorDocument();
  const shape = titleRef(doc);
  assert.equal(await runTool(host(doc), 'setShapeFill', { shape, color: '#C00000' }), null);
  assert.equal(doc.undoLabel, 'Agent: setShapeFill');
  const live = () => pptx.getSlideShapes(pptx.getSlides(doc.pres)[0])[0];
  assert.equal(pptx.getShapeFillColor(live()), '#C00000');
  await doc.undo();
  assert.notEqual(pptx.getShapeFillColor(live()), '#C00000');
});

test('resolves refs in the input and returns refs for what it made', async () => {
  const doc = new EditorDocument();
  const [layout] = await runTool(host(doc), 'listLayouts', {});
  const slide = await runTool(host(doc), 'addSlide', { options: { layout: layout.layout } });
  assert.equal(slide, '/ppt/slides/slide2.xml');
  const shape = await runTool(host(doc), 'addSlideShape', {
    slide,
    opts: { preset: 'rect', x: 914400, y: 914400, w: 1828800, h: 914400, text: 'Made' },
  });
  assert.deepEqual(shape, { slideIndex: 1, slide, shapeId: shape.shapeId, name: shape.name });
  const listed = await runTool(host(doc), 'listShapes', { slide });
  const made = listed.find((entry) => entry.shape.shapeId === shape.shapeId);
  assert.equal(made.text, 'Made');
  assert.deepEqual(made.bounds, { x: 914400, y: 914400, w: 1828800, h: 914400 });

  await runTool(host(doc), 'setShapePosition', { shape: made.shape, x: 0, y: 0 });
  const moved = pptx.findShapeById(pptx.getSlides(doc.pres)[1], shape.shapeId);
  assert.deepEqual(pptx.getShapePosition(moved), { x: 0, y: 0 });

  const table = await runTool(host(doc), 'addSlideTable', {
    slide,
    opts: { x: 0, y: 0, w: 3657600, h: 914400, rows: [['a', 'b']] },
  });
  await runTool(host(doc), 'setTableCellText', { cell: { ...table, row: 0, col: 1 }, text: 'B' });
  const cells = (await runTool(host(doc), 'listShapes', { slide })).find(
    (entry) => entry.shape.shapeId === table.shapeId,
  ).cells;
  assert.deepEqual(cells, [['a', 'B']]);

  const slides = await runTool(host(doc), 'listSlides', {});
  assert.deepEqual(
    slides.map((s) => s.slide),
    ['/ppt/slides/slide1.xml', slide],
  );

  // Bytes arrive as base64, dates as RFC 3339 strings.
  await runTool(host(doc), 'touchModified', { at: '2026-01-02T03:04:05Z' });
  assert.equal(pptx.getPresentationModified(doc.pres)?.toISOString(), '2026-01-02T03:04:05.000Z');
  await runTool(host(doc), 'setThumbnail', { bytes: PNG_1X1 });
  assert.equal(Buffer.from(pptx.getThumbnail(doc.pres).bytes).toString('base64'), PNG_1X1);
});

test('rejects input that does not match the schema, naming each problem', async () => {
  const doc = new EditorDocument();
  const shape = titleRef(doc);
  const rejects = (name, input, pattern) =>
    assert.rejects(runTool(host(doc), name, input), (error) => {
      assert.match(error.message, pattern);
      return true;
    });
  await rejects('setShapeFil', { shape }, /There is no tool "setShapeFil"/);
  await rejects(
    'setShapeFill',
    { shape },
    /^Invalid input for setShapeFill:\n- \(input\): missing required "color"$/,
  );
  await rejects(
    'setShapePosition',
    { shape, x: '5', y: 1.5, z: 0 },
    /- \/x: expected integer, got string "5"\n- \/y: expected integer, got number 1\.5\n- \(input\): unknown property "z" \(allowed: shape, x, y\)/,
  );
  await rejects(
    'setShapeStrokeDash',
    { shape, dash: 'dotted' },
    /\/dash: must be one of "solid", "dot", "dash".*; got "dotted"/,
  );
  await rejects(
    'setShapeFill',
    { shape, color: '#12345' },
    /\/color: matches none of the 3 forms it could take:\n    1. must be one of "bg1"/,
  );
  await rejects(
    'setShapeText',
    { shape: { slide: 'slide1', shapeId: 2 }, value: 'x' },
    /\/shape\/slide: "slide1" does not match the pattern/,
  );
  await rejects('setShapeImage', { shape, bytes: 'not base64!' }, /\/bytes: is not valid base64/);
  await rejects('touchModified', { at: 'yesterday' }, /\/at: is not an RFC 3339 date-time/);
  // A discriminated union explains itself with the variants the value names.
  const chart = (spec) => ({ slide: shape.slide, opts: { spec, x: 0, y: 0, w: 1, h: 1 } });
  await rejects(
    'addSlideChart',
    chart({ kind: 'bars' }),
    /- \/opts\/spec\/kind: must be one of .*"bar".*; got "bars"$/,
  );
  await rejects(
    'addSlideChart',
    chart({ kind: 'pie', categories: ['a'] }),
    /- \/opts\/spec: matches none of the 3 forms it could take:\n {4}1\. missing required "series"/,
  );
  // Nothing ran: no undo step, no change.
  assert.equal(doc.undoLabel, null);
});

test('fails, editing nothing, when a ref names something that is gone', async () => {
  const doc = new EditorDocument();
  const shape = titleRef(doc);
  await assert.rejects(
    runTool(host(doc), 'setShapeText', { shape: { ...shape, shapeId: 999 }, value: 'x' }),
    /There is no shape 999 on \/ppt\/slides\/slide1\.xml; listShapes gives the current ones\./,
  );
  await assert.rejects(
    runTool(host(doc), 'listShapes', { slide: '/ppt/slides/slide9.xml' }),
    /There is no slide \/ppt\/slides\/slide9\.xml/,
  );
  assert.equal(doc.undoLabel, null);
});

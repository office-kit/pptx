import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';
import test from 'node:test';
import { build } from 'esbuild';

// The entry must not reach the editor's UI: a Svelte import fails the build.
const result = await build({
  stdin: {
    contents: `export * from './src/headless.ts';
      export { generatedTools } from './src/manifest/tools.generated.ts';
      export * as pptx from '@office-kit/pptx';`,
    resolveDir: fileURLToPath(new URL('..', import.meta.url)),
  },
  bundle: true,
  write: false,
  metafile: true,
  platform: 'node',
  format: 'esm',
  plugins: [
    {
      name: 'no-svelte',
      setup(builder) {
        builder.onResolve({ filter: /svelte/ }, ({ path, importer }) => ({
          errors: [{ text: `${importer} imports ${path}` }],
        }));
      },
    },
  ],
});
const {
  agentEditSchema,
  readForAgent,
  decodeAgentEdits,
  applyAgentEdits,
  AgentEditError,
  generatedTools,
  pptx,
} = await import(
  `data:text/javascript;base64,${Buffer.from(result.outputFiles[0].text).toString('base64')}`
);

const OPERATIONS = [
  'setShapeText',
  'setTableCellText',
  'setChartData',
  'setShapeTextFormat',
  'setShapeFill',
  'setShapeBounds',
  'removeShape',
  'moveSlide',
  'removeSlide',
];

// The whole set rides on every step of a model's conversation.
const MAX_SCHEMA_BYTES = 12 * 1024;

const EMU_PER_INCH = 914400;

/** A deck with a text box, a table and a column chart on slide 1, and a second slide. */
function deck() {
  const pres = pptx.createPresentation();
  const slide = pptx.addTitleSlide(pres, 'Quarterly');
  pptx.setSlideNotes(slide, 'Say hello');
  const box = pptx.addSlideShape(slide, {
    preset: 'rect',
    x: 0,
    y: 0,
    w: EMU_PER_INCH * 4,
    h: EMU_PER_INCH,
    text: 'Sales rose 10%',
  });
  pptx.setShapeTextFormat(box, { bold: true }, { range: { start: 0, end: 5 } });
  const table = pptx.addSlideTable(slide, {
    x: 0,
    y: EMU_PER_INCH,
    w: EMU_PER_INCH * 4,
    h: EMU_PER_INCH,
    rows: [['a', 'b']],
  });
  const chart = pptx.addSlideChart(slide, {
    x: 0,
    y: EMU_PER_INCH * 2,
    w: EMU_PER_INCH * 4,
    h: EMU_PER_INCH * 3,
    spec: { kind: 'column', categories: ['Q1', 'Q2'], series: [{ name: 'East', values: [1, 2] }] },
  });
  pptx.addTitleSlide(pres, 'Appendix');
  const ref = (shape) => ({
    slide: pptx.getSlidePartName(slide),
    shapeId: pptx.getShapeId(shape),
  });
  return { pres, slide, box: ref(box), table: ref(table), chart: ref(chart) };
}

/** Every subschema of `schema`, with its JSON Pointer. */
function* subschemas(schema, path = '') {
  yield [path, schema];
  for (const key of ['properties', '$defs'])
    for (const [name, child] of Object.entries(schema[key] ?? {}))
      yield* subschemas(child, `${path}/${key}/${name}`);
  for (const key of ['items', 'additionalProperties'])
    if (schema[key] && typeof schema[key] === 'object')
      yield* subschemas(schema[key], `${path}/${key}`);
  for (const key of ['anyOf', 'prefixItems'])
    for (const [i, child] of (schema[key] ?? []).entries())
      yield* subschemas(child, `${path}/${key}/${i}`);
}

test('the entry imports in Node without the editor UI', () => {
  const inputs = Object.keys(result.metafile.inputs);
  assert.deepEqual(
    inputs.filter((input) => /\.svelte(\.ts)?$|[/\\]svelte[/\\]/.test(input)),
    [],
  );
  assert.equal(typeof applyAgentEdits, 'function');
});

test('the schema is the nine operations, closed and self-contained', () => {
  const { edits } = agentEditSchema.properties;
  assert.deepEqual(
    edits.items.anyOf.map((branch) => branch.properties.tool.const),
    OPERATIONS,
  );
  assert.equal(edits.maxItems, 50);

  const problems = [];
  const text = JSON.stringify(agentEditSchema);
  for (const [path, schema] of subschemas(agentEditSchema)) {
    if (path !== '' && '$defs' in schema) problems.push(`${path}: $defs belongs at the root`);
    if (schema.type === 'object' && schema.additionalProperties !== false)
      problems.push(`${path}: an object must not take other properties`);
    if (typeof schema.$ref === 'string' && !agentEditSchema.$defs[schema.$ref.slice(8)])
      problems.push(`${path}: $ref ${schema.$ref} does not resolve`);
  }
  for (const name of Object.keys(agentEditSchema.$defs))
    if (!text.includes(`"#/$defs/${name}"`)) problems.push(`unused $defs/${name}`);
  assert.deepEqual(problems, []);

  // The definitions are the generated ones, gathered in one place.
  for (const tool of generatedTools.tools.filter((t) => OPERATIONS.includes(t.name)))
    for (const [name, def] of Object.entries(tool.input_schema.$defs ?? {}))
      if (name in agentEditSchema.$defs) assert.deepEqual(agentEditSchema.$defs[name], def, name);

  const input = (name) =>
    edits.items.anyOf.find((branch) => branch.properties.tool.const === name).properties.input;
  assert.deepEqual(Object.keys(input('setShapeText').properties), ['shape', 'value']);
  assert.deepEqual(Object.keys(input('setTableCellText').properties), ['cell', 'text']);
  assert.deepEqual(Object.keys(input('setShapeTextFormat').properties.format.properties), [
    'size',
    'bold',
    'italic',
    'color',
  ]);
  assert.deepEqual(input('moveSlide').properties.toIndex, { type: 'integer', minimum: 0 });

  assert.ok(
    text.length <= MAX_SCHEMA_BYTES,
    `agentEditSchema is ${text.length} bytes; keep it under ${MAX_SCHEMA_BYTES}`,
  );
});

test('decoding checks every edit and names the slides they touch', () => {
  const { pres, box, table, chart } = deck();
  const second = pptx.getSlidePartName(pptx.getSlides(pres)[1]);
  const { edits, touchedSlides } = decodeAgentEdits([
    { tool: 'moveSlide', input: { slide: second, toIndex: 0 } },
    { tool: 'setShapeText', input: { shape: box, value: 'x' } },
    { tool: 'setTableCellText', input: { cell: { ...table, row: 0, col: 0 }, text: 'y' } },
    {
      tool: 'setChartData',
      input: { chart, data: { series: [{ name: 'East', values: [3, 4] }] } },
    },
  ]);
  assert.deepEqual(touchedSlides, [second, box.slide]);
  assert.equal(edits.length, 4);
  // Formatting is kept whatever the model sends.
  assert.deepEqual(edits[1].input.options, { preserveFormatting: true });
  assert.deepEqual(edits[2].input.options, { preserveFormatting: true });

  const rejects = (input, pattern) =>
    assert.throws(
      () => decodeAgentEdits(input),
      (error) => {
        assert.ok(error instanceof AgentEditError);
        assert.equal(error.index, null);
        assert.match(error.message, pattern);
        return true;
      },
    );
  rejects([], /needs at least 1 items/);
  rejects(
    Array.from({ length: 51 }, () => ({ tool: 'removeShape', input: { shape: box } })),
    /takes at most 50 items/,
  );
  rejects(
    [{ tool: 'addSlideShape', input: {} }],
    /\/edits\/0\/tool: must be one of "setShapeText"/,
  );
  rejects(
    [
      {
        tool: 'setShapeText',
        input: { shape: box, value: 'x', options: { preserveFormatting: false } },
      },
    ],
    /\/edits\/0\/input: unknown property "options"/,
  );
  rejects(
    [{ tool: 'setShapeTextFormat', input: { shape: box, format: { font: 'Arial' } } }],
    /\/edits\/0\/input\/format: unknown property "font" \(allowed: size, bold, italic, color\)/,
  );
  rejects(
    [
      { tool: 'removeShape', input: { shape: box } },
      { tool: 'moveSlide', input: { slide: second, toIndex: 1.5 } },
    ],
    /\/edits\/1\/input\/toIndex: expected integer, got number 1\.5/,
  );
  rejects(
    [{ tool: 'moveSlide', input: { slide: second, toIndex: -1 } }],
    /\/edits\/0\/input\/toIndex: /,
  );
});

test('applies edits in order, keeping formatting, and reports the one that failed', async () => {
  const { pres, slide, box, chart } = deck();
  const { edits } = decodeAgentEdits([
    { tool: 'setShapeText', input: { shape: box, value: 'Sales rose 12%' } },
    { tool: 'setShapeFill', input: { shape: box, color: '#C00000' } },
    {
      tool: 'setChartData',
      input: {
        chart,
        data: { categories: ['Q1', 'Q2', 'Q3'], series: [{ name: 'East', values: [1, 2, 3] }] },
      },
    },
  ]);
  assert.deepEqual(await applyAgentEdits(pres, edits), [null, null, null]);
  const shape = pptx.findShapeById(slide, box.shapeId);
  assert.equal(pptx.getShapeText(shape), 'Sales rose 12%');
  assert.equal(pptx.getShapeRunFormat(shape, 0, 0)?.bold, true);
  assert.equal(pptx.getShapeFillColor(shape), '#C00000');

  const [read] = readForAgent(pres, [box.slide]).slides;
  assert.equal(read.notes, 'Say hello');
  assert.deepEqual(read.shapes.find((s) => s.kind === 'chart').chart, {
    kind: 'column',
    title: null,
    categories: ['Q1', 'Q2', 'Q3'],
    series: [{ name: 'East', values: [1, 2, 3] }],
  });

  const failing = decodeAgentEdits([
    { tool: 'removeShape', input: { shape: box } },
    { tool: 'setShapeFill', input: { shape: box, color: '#000000' } },
  ]).edits;
  await assert.rejects(applyAgentEdits(pres, failing), (error) => {
    assert.ok(error instanceof AgentEditError);
    assert.equal(error.index, 1);
    assert.match(error.message, /There is no shape/);
    assert.ok(error.cause instanceof Error);
    return true;
  });
  await assert.rejects(
    applyAgentEdits(
      pres,
      decodeAgentEdits([{ tool: 'setChartData', input: { chart, data: { categories: ['A'] } } }])
        .edits,
    ),
    (error) => error.index === 0 && /1 categories/.test(error.message),
  );
});

test('reads every slide, or the ones named', () => {
  const { pres } = deck();
  const { slides } = readForAgent(pres, null);
  assert.deepEqual(
    slides.map(({ index, title, hidden, notes }) => ({ index, title, hidden, notes })),
    [
      { index: 0, title: 'Quarterly', hidden: false, notes: 'Say hello' },
      { index: 1, title: 'Appendix', hidden: false, notes: null },
    ],
  );
  assert.deepEqual(
    readForAgent(pres, [slides[1].slide]).slides.map((s) => s.index),
    [1],
  );
  assert.throws(() => readForAgent(pres, ['/ppt/slides/slide9.xml']), /no slide/);
});

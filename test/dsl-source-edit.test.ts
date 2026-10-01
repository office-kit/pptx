import ts from 'typescript';
import { describe, expect, it } from 'vitest';
import * as api from '../src/api/index.ts';
import * as dsl from '../packages/dsl/src/index.ts';
import * as runtime from '../packages/dsl/src/jsx-runtime.ts';
import { Fragment, Text } from '../packages/dsl/src/index.ts';
import { jsx } from '../packages/dsl/src/jsx-runtime.ts';
import {
  TEXT_EDIT_MAX_LENGTH,
  planTextEdit,
  verifyTextEdit,
} from '../packages/dsl/src/source-edit.ts';

const plan = (source: string, before: string, after: string, slideText = before) =>
  planTextEdit({ files: [{ path: 'slide.tsx', source }], slideText, before, after });
const edited = (source: string, before: string, after: string) => {
  const result = plan(source, before, after);
  if (!result.ok) throw new Error(result.reason);
  expect(result.change).toMatchObject({ path: 'slide.tsx', before: source });
  return result.change.after;
};

describe('planTextEdit', () => {
  it('replaces JSX text, attribute strings, string literals and plain template literals', () => {
    expect(edited('<Text>Hello</Text>', 'Hello', 'A < B & "日本語"')).toBe(
      '<Text>{"A < B & \\"日本語\\""}</Text>',
    );
    expect(edited('<Text>\n    Hello\n    world\n  </Text>', 'Hello world', 'Hi')).toBe(
      '<Text>{"Hi"}</Text>',
    );
    expect(edited('<Shape text="Hello" />', 'Hello', 'Hi')).toBe('<Shape text={"Hi"} />');
    expect(edited("const title = 'Hello';", 'Hello', "it's")).toBe('const title = "it\'s";');
    expect(edited('const title = `Hello`;', 'Hello', 'Hi')).toBe('const title = "Hi";');
  });

  it('refuses computed text as not-found', () => {
    for (const source of [
      'const title = `Hel${"lo"}`;',
      'const word = "Hel"; const title = word;',
      'const title = "Hel" + "lo";',
    ])
      expect(plan(source, 'Hello', 'Hi')).toEqual({ ok: false, reason: 'not-found' });
  });

  it('reports repeated text as ambiguous unless an anchor picks the element', () => {
    const source = [
      '<Slide>',
      '  <Text>Churn</Text>',
      '  <Headline headline="Churn" lead="Why" />',
      '</Slide>',
    ].join('\n');
    const files = [{ path: 'slide.tsx', source }];
    const request = { files, slideText: 'Churn', before: 'Churn', after: 'Retention' };
    expect(planTextEdit(request)).toEqual({ ok: false, reason: 'ambiguous' });
    const anchored = planTextEdit({
      ...request,
      anchor: { path: 'slide.tsx', lineNumber: 3, columnNumber: 3 },
    });
    expect(anchored.ok && anchored.change.after).toBe(
      source.replace('headline="Churn"', 'headline={"Retention"}'),
    );
    // An anchor that is not where an element starts, or names another file, finds nothing.
    for (const anchor of [
      { path: 'slide.tsx', lineNumber: 3, columnNumber: 4 },
      { path: 'slide.tsx', lineNumber: 9, columnNumber: 1 },
      { path: 'other.tsx', lineNumber: 3, columnNumber: 3 },
    ])
      expect(planTextEdit({ ...request, anchor })).toEqual({ ok: false, reason: 'not-found' });
    const repeated = '<KpiRow items={[{ label: "Churn" }, { label: "Churn" }]} />';
    expect(
      planTextEdit({
        ...request,
        files: [{ path: 'slide.tsx', source: repeated }],
        anchor: { path: 'slide.tsx', lineNumber: 1, columnNumber: 1 },
      }),
    ).toEqual({ ok: false, reason: 'ambiguous' });
  });

  it('rejects invalid requests and text that repeats on the slide', () => {
    expect(plan('<Text>Hello</Text>', ' ', 'x')).toEqual({ ok: false, reason: 'invalid' });
    expect(plan('<Text>Hello</Text>', 'Hello', 'Hello')).toEqual({ ok: false, reason: 'invalid' });
    expect(plan('<Text>Hello</Text>', 'Hello', 'x'.repeat(TEXT_EDIT_MAX_LENGTH + 1))).toEqual({
      ok: false,
      reason: 'invalid',
    });
    expect(plan('<Text>Hello</Text>', 'Hello', 'Hi', 'Hello Hello')).toEqual({
      ok: false,
      reason: 'not-unique-on-slide',
    });
  });
});

describe('verifyTextEdit', () => {
  const previous = { slides: ['a', 'b'], slideTexts: ['Hello', 'Other'] };
  const check = (next: typeof previous, after = 'Hi') =>
    verifyTextEdit({ previous, next, slide: 0, before: 'Hello', after });

  it('accepts a rebuild that only changed the edited text', () => {
    expect(check({ slides: ['a2', 'b'], slideTexts: ['Hi', 'Other'] })).toEqual({ ok: true });
    expect(check({ slides: ['a2', 'b'], slideTexts: ['$&', 'Other'] }, '$&')).toEqual({ ok: true });
  });

  it('names what else changed', () => {
    expect(check({ slides: ['a2'], slideTexts: ['Hi'] })).toEqual({
      ok: false,
      reason: 'slide-count-changed',
    });
    expect(check({ slides: ['a2', 'b2'], slideTexts: ['Hi', 'Other'] })).toEqual({
      ok: false,
      reason: 'other-slides-changed',
    });
    expect(check({ slides: ['a2', 'b'], slideTexts: ['Hi Hi', 'Other'] })).toEqual({
      ok: false,
      reason: 'text-mismatch',
    });
  });
});

describe('JSX sources as anchors', () => {
  it('locates the literal behind a shape a prebuilt component made', async () => {
    const source = [
      "import { Presentation, Slide, Text } from '@office-kit/pptx-dsl';",
      "import { Headline } from './parts';",
      'export default (',
      '  <Presentation>',
      '    <Slide>',
      '      <Text x={1} y={3} width={4} height={1}>Churn</Text>',
      '      <Headline headline="Churn" lead="Why" />',
      '    </Slide>',
      '  </Presentation>',
      ');',
    ].join('\n');
    const { outputText } = ts.transpileModule(source, {
      fileName: 'slide.tsx',
      compilerOptions: {
        jsx: ts.JsxEmit.ReactJSXDev,
        jsxImportSource: '@office-kit/pptx-dsl',
        module: ts.ModuleKind.CommonJS,
      },
    });
    const box = { x: 1, y: 1, width: 4, height: 1 };
    // Prebuilt: the component's own elements carry no source.
    const parts = {
      Headline: (props: { headline: string; lead: string }) =>
        jsx(Fragment, {
          children: [
            jsx(Text, { ...box, children: props.headline }),
            jsx(Text, { ...box, children: props.lead }),
          ],
        }),
    };
    const modules: Record<string, unknown> = {
      '@office-kit/pptx-dsl': dsl,
      '@office-kit/pptx-dsl/jsx-dev-runtime': runtime,
      './parts': parts,
    };
    const exports: { default?: dsl.Node } = {};
    new Function('require', 'exports', outputText)((id: string) => modules[id], exports);
    const slide = api.getSlides(await dsl.compile(exports.default!))[0]!;
    const [, headline] = api.getSlideShapes(slide);
    const sources = dsl.getShapeJsxSources(headline!)!;
    expect(sources.map(({ lineNumber, columnNumber }) => [lineNumber, columnNumber])).toEqual([
      [5, 5],
      [7, 7],
    ]);
    const { fileName, ...position } = sources.at(-1)!;
    const result = planTextEdit({
      files: [{ path: 'slide.tsx', source }],
      slideText: 'Churn',
      before: 'Churn',
      after: 'Retention',
      anchor: { path: fileName, ...position },
    });
    expect(result.ok && result.change.after).toBe(
      source.replace('headline="Churn"', 'headline={"Retention"}'),
    );
  });
});

import { describe, expect, it } from 'vitest';
import * as api from '../src/api/index.ts';
import {
  Fragment,
  Group,
  Presentation,
  Raw,
  Shape,
  Slide,
  Text,
  compile,
  getShapeJsxSources,
  type Node,
} from '../packages/dsl/src/index.ts';
import { jsx, jsxDEV } from '../packages/dsl/src/jsx-runtime.ts';

const at = (lineNumber: number, columnNumber: number, fileName = 'slide.tsx') => ({
  fileName,
  lineNumber,
  columnNumber,
});
const box = { x: 1, y: 1, width: 4, height: 1 };
const deck = async (...children: Node[]) =>
  api.getSlideShapes(
    api.getSlides(
      await compile(
        jsx(Presentation, {
          children: jsxDEV(Slide, { children }, undefined, false, at(2, 1)),
        }),
      ),
    )[0]!,
  );

describe('getShapeJsxSources', () => {
  it('returns the enclosing JSX locations of each created shape, outermost first', async () => {
    const [text, shape] = await deck(
      jsxDEV(Text, { ...box, children: 'Title' }, undefined, false, at(3, 3)),
      jsxDEV(Shape, { ...box, preset: 'rect' }, undefined, false, at(4, 3)),
    );
    expect(getShapeJsxSources(text!)).toEqual([at(2, 1), at(3, 3)]);
    expect(getShapeJsxSources(shape!)).toEqual([at(2, 1), at(4, 3)]);
  });

  it('keeps the call site of a component on every shape it returns', async () => {
    // A prebuilt component: its own elements carry no source.
    const Headline = (props: { headline: string; lead: string }) =>
      jsx(Fragment, {
        children: [
          jsx(Text, { ...box, children: props.headline }),
          jsx(Text, { ...box, children: props.lead }),
        ],
      });
    const shapes = await deck(
      jsxDEV(Headline, { headline: 'A', lead: 'B' }, undefined, false, at(6, 5)),
    );
    expect(shapes.map(getShapeJsxSources)).toEqual([
      [at(2, 1), at(6, 5)],
      [at(2, 1), at(6, 5)],
    ]);
  });

  it('nests a located component result under its call site', async () => {
    const Title = (props: { children: string }) =>
      jsxDEV(Text, { ...box, children: props.children }, undefined, false, at(1, 10, 'parts.tsx'));
    const [text] = await deck(jsxDEV(Title, { children: 'A' }, undefined, false, at(3, 3)));
    expect(getShapeJsxSources(text!)).toEqual([at(2, 1), at(3, 3), at(1, 10, 'parts.tsx')]);
  });

  it('gives a group its own location and keeps each member stack', async () => {
    const shapes = await deck(
      jsxDEV(
        Group,
        {
          children: [
            jsxDEV(Text, { ...box, children: 'A' }, undefined, false, at(4, 5)),
            jsxDEV(Shape, { ...box, preset: 'rect' }, undefined, false, at(5, 5)),
          ],
        },
        undefined,
        false,
        at(3, 3),
      ),
    );
    const group = shapes.find((shape) => api.getShapeKind(shape) === 'group')!;
    expect(getShapeJsxSources(group)).toEqual([at(2, 1), at(3, 3)]);
    expect(api.getGroupChildren(group).map(getShapeJsxSources)).toEqual([
      [at(2, 1), at(3, 3), at(4, 5)],
      [at(2, 1), at(3, 3), at(5, 5)],
    ]);
  });

  it('is undefined without the development runtime and for Raw shapes', async () => {
    const pres = await compile(
      jsx(Presentation, {
        children: jsx(Slide, {
          children: [
            jsx(Text, { ...box, children: 'A' }),
            jsx(Raw, {
              scope: 'slide',
              apply: ({ slide }) => {
                api.addSlideTextBox(slide, {
                  x: api.inches(1),
                  y: api.inches(3),
                  w: api.inches(2),
                  h: api.inches(1),
                  text: 'raw',
                });
              },
            }),
          ],
        }),
      }),
    );
    const shapes = api.getSlideShapes(api.getSlides(pres)[0]!);
    expect(shapes).toHaveLength(2);
    expect(shapes.map(getShapeJsxSources)).toEqual([undefined, undefined]);
  });

  it('is undefined for shapes a duplicated source slide brings along', async () => {
    const source = api.createPresentation();
    api.addSlideTextBox(api.addBlankSlide(source), {
      x: api.inches(1),
      y: api.inches(1),
      w: api.inches(2),
      h: api.inches(1),
      text: 'template',
    });
    const pres = await compile(
      jsx(Presentation, {
        source: await api.savePresentation(source),
        mode: 'compose',
        children: jsxDEV(
          Slide,
          {
            from: { index: 0 },
            children: jsxDEV(Text, { ...box, children: 'new' }, undefined, false, at(3, 3)),
          },
          undefined,
          false,
          at(2, 1),
        ),
      }),
    );
    const [template, added] = api.getSlideShapes(api.getSlides(pres)[0]!);
    expect(getShapeJsxSources(template!)).toBeUndefined();
    expect(getShapeJsxSources(added!)).toEqual([at(2, 1), at(3, 3)]);
  });

  it('keeps shape ids across save and load, so a saved id → source map stays valid', async () => {
    const pres = await compile(
      jsx(Presentation, {
        children: [
          jsxDEV(
            Slide,
            {
              children: [
                jsxDEV(Text, { ...box, children: 'A' }, undefined, false, at(3, 3)),
                jsxDEV(
                  Group,
                  {
                    children: [
                      jsxDEV(Shape, { ...box, preset: 'rect' }, undefined, false, at(5, 5)),
                      jsxDEV(Shape, { ...box, preset: 'ellipse' }, undefined, false, at(6, 5)),
                    ],
                  },
                  undefined,
                  false,
                  at(4, 3),
                ),
              ],
            },
            undefined,
            false,
            at(2, 1),
          ),
          jsxDEV(
            Slide,
            { children: jsxDEV(Text, { ...box, children: 'B' }, undefined, false, at(9, 3)) },
            undefined,
            false,
            at(8, 1),
          ),
        ],
      }),
    );
    const map = (presentation: api.PresentationData) =>
      api
        .getSlides(presentation)
        .map((slide) => api.getSlideShapes(slide).map((shape) => api.getShapeId(shape)));
    const before = map(pres);
    expect(
      api
        .getSlides(pres)
        .flatMap((slide) => api.getSlideShapes(slide).map((shape) => getShapeJsxSources(shape)))
        .every(Boolean),
    ).toBe(true);
    expect(map(await api.loadPresentation(await api.savePresentation(pres)))).toEqual(before);
  });
});

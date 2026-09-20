import { describe, expect, it } from 'vitest';
import * as api from '../src/api/index.ts';
import {
  Chart,
  Fill,
  Group,
  Line,
  Media,
  Presentation,
  Raw,
  Shape,
  Slide,
  Table,
  Text,
  compile,
} from '../packages/dsl/src/index.ts';
import { jsx, jsxDEV, jsxs } from '../packages/dsl/src/jsx-runtime.ts';

const title = (value: string) =>
  Text({ x: 1, y: 1, width: 9, height: 1, size: 30, name: 'title', children: value });

describe('declarative authoring', () => {
  it('creates native editable elements through the automatic JSX runtime', async () => {
    const tree = jsx(Presentation, {
      children: jsxs(Slide, {
        children: [
          title('Revenue'),
          Chart({
            x: 5,
            y: 2,
            width: 6,
            height: 4,
            spec: {
              kind: 'column',
              categories: ['Q1'],
              series: [{ name: 'Revenue', values: [120] }],
            },
          }),
          Table({
            x: 1,
            y: 3,
            width: 3,
            height: 1,
            rows: [['Owner'], ['Aiko']],
            headerStyle: { fill: '#15171C', format: { color: '#FFFFFF' } },
          }),
        ],
      }),
    });
    const result = await api.loadPresentation(await api.savePresentation(await compile(tree)));
    const slide = api.getSlides(result)[0]!;
    expect(api.getSlideText(slide)).toContain('Revenue');
    expect(api.getSlideCharts(slide)).toHaveLength(1);
    expect(api.getSlideTables(slide)).toHaveLength(1);
    expect(api.validatePresentation(result).filter((issue) => issue.severity === 'error')).toEqual(
      [],
    );
  });

  it('reuses source slides without modifying input or accumulating watch rebuilds', async () => {
    const source = await api.savePresentation(
      await compile(Presentation({ children: Slide({ children: title('Original') }) })),
    );
    const tree = Presentation({
      source,
      mode: 'compose',
      children: Slide({
        from: { index: 0 },
        children: Fill({ target: { name: 'title' }, children: 'Updated' }),
      }),
    });
    for (let i = 0; i < 2; i++) {
      const result = await compile(tree);
      expect(api.getSlideCount(result)).toBe(1);
      expect(api.getSlideText(api.getSlides(result)[0]!)).toBe('Updated');
    }
    expect(api.getSlideText(api.getSlides(await api.loadPresentation(source))[0]!)).toBe(
      'Original',
    );
  });

  it('edits selected slides and retains unmentioned slides', async () => {
    const source = await api.savePresentation(
      await compile(
        Presentation({
          children: [Slide({ children: title('First') }), Slide({ children: title('Second') })],
        }),
      ),
    );
    const result = await compile(
      Presentation({
        source,
        children: Slide({
          target: { index: 1 },
          children: Fill({ target: { name: 'title' }, text: 'Changed' }),
        }),
      }),
    );
    expect(api.getSlides(result).map(api.getSlideText)).toEqual(['First', 'Changed']);
  });

  it('runs Raw after construction with the exact shape context', async () => {
    const result = await compile(
      Presentation({
        children: Slide({
          children: [
            Raw({
              apply: ({ slide }) => {
                expect(api.getSlideShapes(slide!)).toHaveLength(1);
              },
            }),
            Shape({
              x: 1,
              y: 1,
              width: 2,
              height: 2,
              preset: 'rect',
              children: Raw({
                apply: ({ shape }) => {
                  api.setShapeRotation(shape!, 15);
                },
              }),
            }),
          ],
        }),
      }),
    );
    expect(api.getShapeRotation(api.getSlideShapes(api.getSlides(result)[0]!)[0]!)).toBe(15);
  });

  it('adds embedded and online media with inch bounds', async () => {
    const mp3 = new Uint8Array([0x49, 0x44, 0x33, 3, 0, 0, 0, 0, 0, 0]);
    const result = await compile(
      Presentation({
        children: Slide({
          children: [
            Media({ kind: 'audio', data: mp3, x: 1, y: 1, width: 1, height: 1, name: 'jingle' }),
            Media({
              kind: 'online',
              url: 'https://youtu.be/dQw4w9WgXcQ',
              x: 3,
              y: 1,
              width: 4,
              height: 2.25,
            }),
          ],
        }),
      }),
    );
    const [audio, online] = api.findShapesWithMedia(api.getSlides(result)[0]!);
    expect(api.getShapeName(audio!)).toBe('jingle');
    expect(api.getShapeMedia(audio!)).toMatchObject({ kind: 'audio', contentType: 'audio/mpeg' });
    expect(api.getShapeBounds(audio!)).toMatchObject({ x: api.inches(1), w: api.inches(1) });
    expect(api.getShapeMedia(online!)).toEqual({
      kind: 'online',
      url: 'https://www.youtube.com/embed/dQw4w9WgXcQ?feature=oembed',
    });
  });

  it('fails for ambiguous references and incompatible parent scopes', async () => {
    await expect(compile(Presentation({ children: title('Invalid') }))).rejects.toThrow(
      'child of Slide',
    );
    await expect(
      compile(Presentation({ children: Slide({ target: { index: 0 } }) })),
    ).rejects.toThrow('mode="edit"');
    await expect(
      compile(
        Presentation({
          children: Slide({
            children: [title('A'), title('B'), Fill({ target: { name: 'title' }, text: 'C' })],
          }),
        }),
      ),
    ).rejects.toThrow('matched 2 shapes');
  });
});

it('preserves unknown package content and untouched slide bytes while editing a template', async () => {
  const original = await compile(
    Presentation({
      children: [Slide({ children: title('First') }), Slide({ children: title('Second') })],
    }),
  );
  const pkg = api._internalPackageOf(original);
  const { partName } = await import('../src/internal/opc/index.ts');
  const payload = new Uint8Array([255, 0, 19, 234]);
  pkg.addPart(partName('/vendor/opaque.bin'), 'application/x-private', payload);
  const untouched = api.getSlidePartName(api.getSlides(original)[0]!);
  const before = api.readPackagePart(original, untouched);
  const source = await api.savePresentation(original);
  const edited = await api.loadPresentation(
    await api.savePresentation(
      await compile(
        Presentation({
          source,
          mode: 'edit',
          children: Slide({
            target: { index: 1 },
            children: Fill({ target: { name: 'title' }, text: 'Changed' }),
          }),
        }),
      ),
    ),
  );
  expect(api.readPackagePart(edited, '/vendor/opaque.bin')).toEqual(payload);
  expect(api.readPackagePart(edited, untouched)).toEqual(before);
});

it('allows typed Raw inside text and rejects a mismatched requested scope', async () => {
  const result = await compile(
    Presentation({
      children: Slide({
        children: Text({
          x: 1,
          y: 1,
          width: 3,
          height: 1,
          children: [
            'Hello',
            Raw({
              scope: 'shape',
              apply: ({ shape }) => {
                api.setShapeRotation(shape, 10);
              },
            }),
          ],
        }),
      }),
    }),
  );
  const shape = api.getSlideShapes(api.getSlides(result)[0]!)[0]!;
  expect(api.getShapeText(shape)).toBe('Hello');
  expect(api.getShapeRotation(shape)).toBe(10);
  await expect(
    compile(Presentation({ children: Raw({ scope: 'shape', apply: () => {} }) })),
  ).rejects.toThrow('cannot run inside presentation');
});

describe('lines, groups and richer text and table styling', () => {
  const box = { x: 1, y: 1, width: 3, height: 1 };
  const firstSlideShapes = (pres: api.PresentationData) =>
    api.getSlideShapes(api.getSlides(pres)[0]!);

  it('draws a line between two points given in inches', async () => {
    const pres = await compile(
      Presentation({
        children: Slide({
          children: Line({ x1: 1, y1: 2, x2: 4, y2: 2, color: '#FF0000', width: 2, name: 'rule' }),
        }),
      }),
    );
    const [line] = firstSlideShapes(pres);
    expect(api.getShapeKind(line!)).toBe('connector');
    expect(api.getShapeName(line!)).toBe('rule');
    expect(api.getShapeStrokeWidth(line!)).toBe(api.pt(2));
    expect(api.getShapeStrokeColor(line!)).toBe('#FF0000');
  });

  it('groups the shapes its children create, including a nested group', async () => {
    const pres = await compile(
      Presentation({
        children: Slide({
          children: Group({
            name: 'card',
            children: [
              Shape({ ...box, preset: 'rect' }),
              Group({
                children: [Text({ ...box, children: 'a' }), Text({ ...box, children: 'b' })],
              }),
            ],
          }),
        }),
      }),
    );
    // getSlideShapes flattens groups: the outer group comes first, then its
    // descendants.
    const [card, ...descendants] = firstSlideShapes(pres);
    expect(descendants).toHaveLength(4);
    expect(api.getShapeName(card!)).toBe('card');
    const children = api.getGroupChildren(card!);
    expect(children.map((child) => api.getShapeKind(child))).toEqual(['shape', 'group']);
    expect(api.getGroupChildren(children[1]!)).toHaveLength(2);
  });

  it('rejects a group of fewer than two shapes', async () => {
    await expect(
      compile(
        Presentation({
          children: Slide({ children: Group({ children: Shape({ ...box, preset: 'rect' }) }) }),
        }),
      ),
    ).rejects.toThrow('at least 2 shapes');
  });

  it('applies bullets and spacing to every paragraph of a Text', async () => {
    const pres = await compile(
      Presentation({
        children: Slide({
          children: Text({
            ...box,
            bullets: 'bullet',
            paragraphSpacing: { before: 2, after: 9 },
            children: 'one\ntwo\nthree',
          }),
        }),
      }),
    );
    const [text] = firstSlideShapes(pres);
    expect(api.getShapeParagraphCount(text!)).toBe(3);
    for (let i = 0; i < 3; i++) {
      expect(api.getParagraphBullet(text!, i)).toBe('bullet');
      expect(api.getParagraphSpacing(text!, i)).toMatchObject({ beforePts: 2, afterPts: 9 });
    }
  });

  it('positions the text of a Shape, and refuses to without text', async () => {
    const pres = await compile(
      Presentation({
        children: Slide({
          children: Shape({
            ...box,
            preset: 'ellipse',
            text: '1',
            align: 'center',
            anchor: 'center',
          }),
        }),
      }),
    );
    const [shape] = firstSlideShapes(pres);
    expect(api.getParagraphAlignment(shape!, 0)).toBe('ctr');
    expect(api.getShapeTextAnchor(shape!)).toBe('center');
    await expect(
      compile(
        Presentation({
          children: Slide({ children: Shape({ ...box, preset: 'rect', anchor: 'center' }) }),
        }),
      ),
    ).rejects.toThrow('set text as well');
  });

  it('merges styleCell over the cell, header and stripe styles', async () => {
    const pres = await compile(
      Presentation({
        children: Slide({
          children: Table({
            x: 1,
            y: 1,
            width: 6,
            height: 2,
            rows: [
              ['Name', 'Status'],
              ['A', 'ok'],
              ['B', 'late'],
            ],
            cellStyle: {
              fill: '#FFFFFF',
              align: 'left',
              borders: { bottom: { color: '#CCCCCC', width: 0.75 } },
            },
            headerStyle: { fill: '#000000' },
            stripeFill: '#EEEEEE',
            styleCell: ({ row, value }) =>
              row > 0 && value === 'late'
                ? { fill: '#FF0000', align: 'center', borders: { left: { color: '#FF0000' } } }
                : undefined,
          }),
        }),
      }),
    );
    const [table] = firstSlideShapes(pres);
    const cell = (r: number, c: number) => api.getTableCell(table!, r, c);
    expect(api.getTableCellFill(cell(0, 0))).toBe('#000000');
    expect(api.getTableCellFill(cell(1, 0))).toBe('#FFFFFF');
    // Row 2 is striped, and its own style still wins over the stripe.
    expect(api.getTableCellFill(cell(2, 0))).toBe('#EEEEEE');
    expect(api.getTableCellFill(cell(2, 1))).toBe('#FF0000');
    expect(api.getTableCellAlignment(cell(2, 0))).toBe('l');
    expect(api.getTableCellAlignment(cell(2, 1))).toBe('ctr');
    const borders = api.getTableCellBorders(pres, cell(2, 1));
    expect(borders.bottom).toMatchObject({ color: '#CCCCCC', widthEmu: api.pt(0.75) });
    expect(borders.left).toMatchObject({ color: '#FF0000' });
    expect(api.getTableCellBorders(pres, cell(1, 0)).left).toBeNull();
  });
});

// packages/dsl/README.md documents this prefix as what the dev JSX transform buys.
it('prefixes a compile error with the source location of each enclosing element', async () => {
  const at = (lineNumber: number, columnNumber: number) => ({
    fileName: 'deck.tsx',
    lineNumber,
    columnNumber,
  });
  const fill = jsxDEV(
    Fill,
    { target: { name: 'Titel 1' }, children: 'x' },
    undefined,
    false,
    at(7, 9),
  );
  const slide = jsxDEV(Slide, { children: fill }, undefined, false, at(6, 7));
  await expect(compile(jsx(Presentation, { children: slide }))).rejects.toThrow(
    'deck.tsx:6:7 <Slide>: deck.tsx:7:9 <Fill>: Shape reference matched 0 shapes: {"name":"Titel 1"}',
  );
  const plain = jsx(Slide, { children: jsx(Fill, { target: { name: 'Titel 1' }, children: 'x' }) });
  await expect(compile(jsx(Presentation, { children: plain }))).rejects.toThrow(
    /^Shape reference matched 0 shapes/,
  );
});

// Office InkML subset: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/096dacae-0d2c-4861-bc4d-c8e4c6405ad3
import { emuCoordinate, emuExtent, lineWidthEmu } from '../bounds.ts';
import { attr, elem, qname, text, type XmlElement } from '../xml/index.ts';

export const INKML_NAMESPACE = 'http://www.w3.org/2003/InkML';
const name = (local: string) => qname('inkml', local, INKML_NAMESPACE);
const attribute = (key: string, value: string | number) => attr(qname('', key, ''), String(value));
const id = (value: string) =>
  attr(qname('xml', 'id', 'http://www.w3.org/XML/1998/namespace'), value);

export interface InkStroke {
  /** Slide coordinates and pen diameter in EMU. */
  points: ReadonlyArray<{ x: number; y: number }>;
  widthEmu: number;
  color: string;
}

export interface BuiltInk {
  root: XmlElement;
  bounds: { x: number; y: number; cx: number; cy: number };
}

/** Build constant-width pen traces; packaging and slide contentPart are separate. */
export const buildInk = (strokes: ReadonlyArray<InkStroke>): BuiltInk => {
  if (!strokes.length) throw new RangeError('ink requires at least one stroke');
  let left = Infinity,
    top = Infinity,
    right = -Infinity,
    bottom = -Infinity;
  const normalized = strokes.map((stroke) => {
    if (!stroke.points.length) throw new RangeError('ink stroke requires at least one point');
    if (!/^#[0-9a-f]{6}$/i.test(stroke.color)) throw new RangeError('ink color must be #RRGGBB');
    const width = lineWidthEmu(stroke.widthEmu, 'ink width');
    if (!width) throw new RangeError('ink width must be positive');
    const points = stroke.points.map((point) => {
      const x = emuCoordinate(point.x, 'ink x');
      const y = emuCoordinate(point.y, 'ink y');
      left = Math.min(left, x - width / 2);
      top = Math.min(top, y - width / 2);
      right = Math.max(right, x + width / 2);
      bottom = Math.max(bottom, y + width / 2);
      return { x, y };
    });
    return { points, width, color: stroke.color.toUpperCase() };
  });
  const x = emuCoordinate(Math.floor(left), 'ink bounds x');
  const y = emuCoordinate(Math.floor(top), 'ink bounds y');
  const bounds = {
    x,
    y,
    cx: emuExtent(Math.ceil(right) - x, 'ink bounds width'),
    cy: emuExtent(Math.ceil(bottom) - y, 'ink bounds height'),
  };
  // One integer sample unit = one EMU, with explicit resolution so Office
  // does not infer a device-dependent scale. All traces share this context.
  const context = elem(name('context'), {
    attrs: [id('ctx0')],
    children: [
      elem(name('inkSource'), {
        attrs: [id('src0')],
        children: [
          elem(name('traceFormat'), {
            children: ['X', 'Y'].map((axis) =>
              elem(name('channel'), {
                attrs: [
                  attribute('name', axis),
                  attribute('type', 'integer'),
                  attribute('units', 'cm'),
                ],
              }),
            ),
          }),
          elem(name('channelProperties'), {
            children: ['X', 'Y'].map((axis) =>
              elem(name('channelProperty'), {
                attrs: [
                  attribute('channel', axis),
                  attribute('name', 'resolution'),
                  attribute('value', 360000),
                  attribute('units', '1/cm'),
                ],
              }),
            ),
          }),
        ],
      }),
    ],
  });
  const brushes = normalized.map((stroke, index) =>
    elem(name('brush'), {
      attrs: [id(`br${index}`)],
      children: [
        ...['width', 'height'].map((key) =>
          elem(name('brushProperty'), {
            attrs: [
              attribute('name', key),
              attribute('value', stroke.width / 360000),
              attribute('units', 'cm'),
            ],
          }),
        ),
        elem(name('brushProperty'), {
          attrs: [attribute('name', 'color'), attribute('value', stroke.color)],
        }),
        elem(name('brushProperty'), {
          attrs: [attribute('name', 'ignorePressure'), attribute('value', '1')],
        }),
      ],
    }),
  );
  const traces = normalized.map((stroke, index) =>
    elem(name('trace'), {
      attrs: [attribute('contextRef', '#ctx0'), attribute('brushRef', `#br${index}`)],
      // Explicit absolute markers avoid InkML's first/second difference state.
      children: [text(stroke.points.map((point) => `!${point.x - x} ${point.y - y}`).join(', '))],
    }),
  );
  return {
    bounds,
    root: elem(name('ink'), {
      prefixDecls: new Map([['inkml', INKML_NAMESPACE]]),
      children: [elem(name('definitions'), { children: [context, ...brushes] }), ...traces],
    }),
  };
};

import { expect, it } from 'vitest';
import { buildInk, INKML_NAMESPACE } from '../src/internal/presentationml/ink-builder.ts';
import {
  allChildElements,
  firstChildElement,
  getAttrValue,
  parseXml,
  qname,
  serializeFragment,
  textContent,
} from '../src/internal/xml/index.ts';

const name = (local: string) => qname('inkml', local, INKML_NAMESPACE);
const value = (element: Parameters<typeof getAttrValue>[0], local: string) =>
  getAttrValue(element, qname('', local, ''));

it('serializes independently referenced brushes and explicit physical coordinates', () => {
  const ink = buildInk([
    {
      points: [
        { x: -360000, y: 720000 },
        { x: 0, y: 360000 },
      ],
      widthEmu: 3600,
      color: '#ff0000',
    },
    { points: [{ x: 360000, y: 0 }], widthEmu: 7200, color: '#00ff00' },
  ]);
  expect(ink.bounds).toEqual({ x: -361800, y: -3600, cx: 725400, cy: 725400 });
  const root = parseXml(serializeFragment(ink.root)).root;
  const definitions = firstChildElement(root, name('definitions'))!;
  const source = firstChildElement(
    firstChildElement(definitions, name('context'))!,
    name('inkSource'),
  )!;
  const properties = allChildElements(
    firstChildElement(source, name('channelProperties'))!,
    name('channelProperty'),
  );
  expect(
    properties.map((property) => [
      value(property, 'channel'),
      value(property, 'value'),
      value(property, 'units'),
    ]),
  ).toEqual([
    ['X', '360000', '1/cm'],
    ['Y', '360000', '1/cm'],
  ]);
  const brushes = allChildElements(definitions, name('brush'));
  expect(
    brushes.map((brush) =>
      allChildElements(brush, name('brushProperty')).map((property) => value(property, 'value')),
    ),
  ).toEqual([
    ['0.01', '0.01', '#FF0000', '1'],
    ['0.02', '0.02', '#00FF00', '1'],
  ]);
  const traces = allChildElements(root, name('trace'));
  expect(
    traces.map((trace) => [
      value(trace, 'contextRef'),
      value(trace, 'brushRef'),
      textContent(trace),
    ]),
  ).toEqual([
    ['#ctx0', '#br0', '!1800 723600, !361800 363600'],
    ['#ctx0', '#br1', '!721800 3600'],
  ]);
});

it('gives a stationary pen mark a nonzero bounding box including the brush radius', () => {
  const ink = buildInk([{ points: [{ x: 20.4, y: 30.6 }], widthEmu: 3, color: '#123456' }]);
  expect(ink.bounds).toEqual({ x: 18, y: 29, cx: 4, cy: 4 });
});

it('rejects empty, nonfinite, out-of-range and malformed input before serialization', () => {
  const stroke = { points: [{ x: 0, y: 0 }], widthEmu: 3600, color: '#FF0000' };
  expect(() => buildInk([])).toThrow();
  for (const patch of [
    { points: [] },
    { points: [{ x: NaN, y: 0 }] },
    { points: [{ x: 1e20, y: 0 }] },
    { widthEmu: 0 },
    { widthEmu: -1 },
    { widthEmu: Infinity },
    { color: 'red' },
    { color: '#ff0000<bad>' },
  ])
    expect(() => buildInk([{ ...stroke, ...patch }])).toThrow();
});

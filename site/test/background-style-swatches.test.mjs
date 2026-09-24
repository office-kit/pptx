import assert from 'node:assert/strict';
import test from 'node:test';
import { backgroundStyleSwatches } from '../src/lib/editor/core/background-style-swatches.ts';

test('background swatches retain rectangular contours, focus and resolved stop colors', () => {
  const styles = [
    {
      style: 1,
      selected: false,
      fill: { kind: 'gradient' },
      gradient: {
        path: 'rect',
        focus: { left: 0.25, top: 0.25, right: 0.25, bottom: 0.25 },
        stops: [
          {
            offset: 0,
            color: 'scheme:bg1',
            resolvedColor: '#BCBCBC',
            colorTransforms: [{ kind: 'tint', value: 0.5 }],
          },
          { offset: 1, color: 'scheme:bg1', resolvedColor: '#000000', opacity: 0.5 },
        ],
      },
    },
  ];
  const before = structuredClone(styles);
  const image = backgroundStyleSwatches(styles).get(1);
  const svg = decodeURIComponent(image.slice('data:image/svg+xml,'.length));
  assert.equal((svg.match(/<polygon /g) ?? []).length, 4);
  assert.match(svg, /width="0.5" height="0.5" fill="#BCBCBC"/);
  assert.match(svg, /stop-color="#000000" stop-opacity="0.5"/);
  assert.doesNotMatch(svg, /radialGradient|scheme:/);
  assert.deepEqual(styles, before);
});

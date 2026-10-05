import { describe, expect, it } from 'vitest';
import { inches } from '../src/api/index.ts';
import { planPropEdit, type PropEdit } from '../packages/dsl/src/source-edit.ts';

const EMU = 914400;
const anchorOf = (source: string, marker = '<') => {
  const offset = source.indexOf(marker);
  const before = source.slice(0, offset).split('\n');
  return { path: 'slide.tsx', lineNumber: before.length, columnNumber: before.at(-1)!.length + 1 };
};
const plan = (source: string, edits: PropEdit[], marker?: string) =>
  planPropEdit({ file: { path: 'slide.tsx', source }, anchor: anchorOf(source, marker), edits });
const edited = (source: string, edits: PropEdit[], marker?: string) => {
  const result = plan(source, edits, marker);
  if (!result.ok) throw new Error(result.reason);
  return result.change.after;
};
const move = (prop: PropEdit['prop'], from: number, to: number): PropEdit => ({
  prop,
  from: prop === 'rotation' ? Math.round(from * 60000) : inches(from),
  to: prop === 'rotation' ? Math.round(to * 60000) : inches(to),
});

describe('planPropEdit', () => {
  it('replaces literals with the shortest decimal that reproduces the wire value', () => {
    expect(edited('<Shape x={1} y={2} />', [move('x', 1, 1.5), move('y', 2, 0.25)])).toBe(
      '<Shape x={1.5} y={0.25} />',
    );
    expect(edited('<Shape x={-1} />', [move('x', -1, 2)])).toBe('<Shape x={2} />');
    // An arbitrary drag lands on whole EMU; the written decimal rounds back to it.
    const written = edited('<Shape x={1} />', [{ prop: 'x', from: EMU, to: 1234567 }]);
    const value = Number(/x=\{([^}]+)\}/.exec(written)![1]);
    expect(inches(value)).toBe(1234567);
    expect(edited('<Shape rotation={10} />', [move('rotation', 10, 12.5)])).toBe(
      '<Shape rotation={12.5} />',
    );
  });

  it('keeps an expression and adds a delta, folding later edits into it', () => {
    const once = edited('<Shape x={col * 2.5} />', [
      { prop: 'x', from: inches(5), to: inches(5.3) },
    ]);
    expect(once).toBe('<Shape x={col * 2.5 + 0.3} />');
    const twice = edited(once, [{ prop: 'x', from: inches(5.3), to: inches(5.1) }]);
    expect(twice).toBe('<Shape x={col * 2.5 + 0.1} />');
    const back = edited(twice, [{ prop: 'x', from: inches(5.1), to: inches(5) }]);
    expect(back).toBe('<Shape x={col * 2.5} />');
    expect(
      edited('<Shape width={W} />', [{ prop: 'width', from: inches(3), to: inches(2.5) }]),
    ).toBe('<Shape width={W - 0.5} />');
    expect(
      edited('<Shape x={wide ? 1 : 2} />', [{ prop: 'x', from: inches(1), to: inches(1.25) }]),
    ).toBe('<Shape x={(wide ? 1 : 2) + 0.25} />');
  });

  it('turns rotation deltas the short way round', () => {
    expect(edited('<Shape rotation={angle} />', [move('rotation', 350, 10)])).toBe(
      '<Shape rotation={angle + 20} />',
    );
  });

  it('inserts a missing prop and edits only the anchored element', () => {
    const source = '<>\n  <Shape x={1} />\n  <Text x={1}>Hi</Text>\n</>';
    expect(edited(source, [move('x', 1, 2), move('rotation', 0, 30)], '<Text')).toBe(
      '<>\n  <Shape x={1} />\n  <Text x={2} rotation={30}>Hi</Text>\n</>',
    );
  });

  it('refuses what it cannot edit in place', () => {
    expect(plan('<Shape {...props} />', [move('x', 1, 2)])).toEqual({
      ok: false,
      reason: 'spread',
    });
    expect(plan('<KpiCard x={1} />', [move('x', 1, 2)])).toEqual({
      ok: false,
      reason: 'not-shape-element',
    });
    expect(plan('<Shape x="1" />', [move('x', 1, 2)])).toEqual({
      ok: false,
      reason: 'not-numeric',
    });
    expect(plan('<Shape x={1} />', [{ prop: 'x', from: 1.5, to: 2 }])).toEqual({
      ok: false,
      reason: 'invalid',
    });
    expect(plan('const a = 1; <Shape />', [move('x', 1, 2)], 'const')).toEqual({
      ok: false,
      reason: 'not-element',
    });
  });
});

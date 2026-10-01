import assert from 'node:assert/strict';
import test from 'node:test';
import { OutlineSelectionModel } from '../src/lib/editor/core/outline-selection.ts';

globalThis.Node = { TEXT_NODE: 3 };
globalThis.HTMLElement = class {};
globalThis.HTMLBRElement = class extends HTMLElement {};

function createFields(values, { raf } = {}) {
  return values.map((value, order) => {
    const root = {
      childNodes: [],
      firstChild: null,
      isConnected: true,
      compareDocumentPosition(other) {
        return order < other.order ? 4 : 2;
      },
      ownerDocument: {
        defaultView: raf
          ? { requestAnimationFrame: (callback) => (raf.push(callback), raf.length) }
          : null,
        getSelection: () => null,
      },
      order,
    };
    return {
      key: String(order),
      root,
      text: () => value,
      copy: (start, end) => ({ text: value.slice(start, end), formats: [] }),
      flush: () => [],
      apply: () => {},
      transact: (_label, callback) => callback(),
      focus: () => {},
      setRange: () => {},
    };
  });
}

function register(model, fields) {
  for (const field of fields) model.register(field);
  return model.fields();
}

test('extend includes the complete last field when bridging forward', () => {
  const model = new OutlineSelectionModel();
  const fields = register(model, createFields(['A', 'BC']));
  model.update(fields[0], 1, 1);
  assert.equal(model.extend(fields[0], 1), true);
  assert.equal(model.extend(fields[1], 1), true);
  assert.deepEqual(model.current(), {
    start: { key: '0', offset: 1 },
    end: { key: '1', offset: 2 },
  });
});

test('extend includes the complete first field when bridging backward', () => {
  const model = new OutlineSelectionModel();
  const fields = register(model, createFields(['AB', 'C']));
  model.update(fields[1], 0, 0);
  assert.equal(model.extend(fields[1], -1), true);
  assert.equal(model.extend(fields[0], -1), true);
  assert.deepEqual(model.current(), {
    start: { key: '0', offset: 0 },
    end: { key: '1', offset: 0 },
  });
});

test('a later replacement invalidates an earlier queued caret restore', () => {
  const raf = [];
  const model = new OutlineSelectionModel();
  const fields = register(model, createFields(['abc'], { raf }));
  const calls = [];
  fields[0].setRange = (offset) => calls.push(['range', offset]);
  fields[0].focus = (offset) => calls.push(['focus', offset]);
  model.setCaret(fields[0], 0);
  assert.equal(model.replace('x'), true);
  assert.equal(model.replace('y'), true);
  assert.equal(raf.length, 2);
  raf.splice(0).forEach((callback) => callback());
  assert.deepEqual(calls, [
    ['range', 2],
    ['focus', 2],
  ]);
});

test('clear invalidates a queued replacement caret restore', () => {
  const raf = [];
  const model = new OutlineSelectionModel();
  const fields = register(model, createFields(['abc'], { raf }));
  const calls = [];
  fields[0].setRange = (offset) => calls.push(['range', offset]);
  fields[0].focus = (offset) => calls.push(['focus', offset]);
  model.setCaret(fields[0], 0);
  assert.equal(model.replace('x'), true);
  model.clear();
  raf.splice(0).forEach((callback) => callback());
  assert.deepEqual(calls, []);
});

test('replacement clears the bridge anchor preservation mode', () => {
  const model = new OutlineSelectionModel();
  const fields = register(model, createFields(['A', 'BC']));
  model.update(fields[0], 1, 1);
  assert.equal(model.extend(fields[0], 1), true);
  assert.equal(model.replace('x'), true);
  model.update(fields[0], 0, 0);
  assert.deepEqual(model.current(), {
    start: { key: '0', offset: 0 },
    end: { key: '0', offset: 0 },
  });
});

test('re-registering a key with a new root clears the old logical range', () => {
  const model = new OutlineSelectionModel();
  const field = createFields(['A'])[0];
  const unregister = model.register(field);
  model.setCaret(field, 1);
  const replacement = createFields(['A'])[0];
  model.register(replacement);
  assert.equal(model.current(), null);
  unregister();
  assert.deepEqual(model.fields(), [replacement]);
});

test('collapse moves a cross-field range to the logical edge', () => {
  const raf = [];
  const model = new OutlineSelectionModel();
  const fields = register(model, createFields(['A', 'CD'], { raf }));
  const calls = [];
  fields[0].focus = (offset) => calls.push(['first', offset]);
  fields[1].focus = (offset) => calls.push(['last', offset]);
  model.update(fields[0], 1, 1);
  assert.equal(model.extend(fields[0], 1), true);
  assert.equal(model.extend(fields[1], 1), true);
  assert.equal(model.collapse(-1), true);
  assert.deepEqual(model.current(), {
    start: { key: '0', offset: 1 },
    end: { key: '0', offset: 1 },
  });
  assert.deepEqual(calls, [
    ['last', 0],
    ['first', 1],
  ]);
  assert.equal(model.collapse(1), true);
});

test('collapse invalidates a queued replacement caret restore', () => {
  const raf = [];
  const model = new OutlineSelectionModel();
  const fields = register(model, createFields(['abc'], { raf }));
  const calls = [];
  fields[0].focus = (offset) => calls.push(offset);
  model.setCaret(fields[0], 0);
  model.replace('x');
  model.collapse(-1);
  assert.equal(raf.length, 1);
  raf.splice(0).forEach((callback) => callback());
  assert.deepEqual(calls, [1]);
});

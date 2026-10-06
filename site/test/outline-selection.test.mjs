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
      formats: () => [],
      applyFormat: () => {},
      changeCase: (_start, end) => end,
      transact: (_label, callback) => callback(),
      focus: () => {},
      setRange: () => {},
    };
  });
}

test('formats every selected field in one transaction and leaves the edges untouched', () => {
  const transactions = [];
  const ranges = [];
  const fields = createFields(['abcd', 'EF']);
  for (const field of fields) {
    field.transact = (label, callback) => {
      transactions.push(label);
      callback();
    };
    field.applyFormat = (start, end, format, reset) =>
      ranges.push({ key: field.key, start, end, format, reset });
  }
  const model = new OutlineSelectionModel();
  register(model, fields);
  model.update(fields[0], 2, 4);
  assert.equal(model.extend(fields[0], 1), true);
  assert.equal(model.extend(fields[1], 1), true);
  assert.equal(model.format({ bold: true }), true);
  assert.deepEqual(transactions, ['Format selected text']);
  assert.deepEqual(ranges, [
    { key: '0', start: 2, end: 4, format: { bold: true }, reset: false },
    { key: '1', start: 0, end: 2, format: { bold: true }, reset: false },
  ]);
});

test('change case expands a caret to the current word and restores the transformed caret', () => {
  const model = new OutlineSelectionModel();
  const fields = register(model, createFields(['one two']));
  const calls = [];
  fields[0].changeCase = (start, end, value, caret) => {
    calls.push({ start, end, value, caret });
    return end + 1;
  };
  model.setCaret(fields[0], 1);
  assert.equal(model.changeCase('upper'), true);
  assert.deepEqual(calls, [{ start: 0, end: 3, value: 'upper', caret: 1 }]);
  assert.deepEqual(model.current(), {
    start: { key: '0', offset: 4 },
    end: { key: '0', offset: 4 },
  });
});

test('change case applies one transaction across selected outline fields', () => {
  const model = new OutlineSelectionModel();
  const fields = register(model, createFields(['Heading', 'Body']));
  const calls = [];
  fields.forEach((field) => {
    field.changeCase = (start, end, value) => {
      calls.push({ key: field.key, start, end, value });
      return end + 1;
    };
  });
  model.setCaret(fields[0], 2);
  model.update(fields[1], 0, 4, true);
  assert.equal(model.changeCase('title'), true);
  assert.deepEqual(calls, [
    { key: '0', start: 2, end: 7, value: 'title' },
    { key: '1', start: 0, end: 4, value: 'title' },
  ]);
  assert.deepEqual(model.current(), {
    start: { key: '0', offset: 2 },
    end: { key: '1', offset: 5 },
  });
});

test('format callback reads formats after pending edits are flushed', () => {
  const transactions = [];
  const seen = [];
  const fields = createFields(['ab', 'cd']);
  fields[0].formats = () => [{ italic: true }];
  fields[1].formats = () => [{ bold: true }];
  fields[0].transact = (_label, callback) => {
    transactions.push(1);
    callback();
  };
  fields[0].applyFormat = (_start, _end, format) => seen.push(format);
  fields[1].applyFormat = (_start, _end, format) => seen.push(format);
  const model = new OutlineSelectionModel();
  register(model, fields);
  model.update(fields[0], 1, 2);
  assert.equal(model.extend(fields[0], 1), true);
  assert.equal(model.extend(fields[1], 1), true);
  assert.equal(
    model.format((formats) => ({ bold: formats.some((format) => format.bold) })),
    true,
  );
  assert.deepEqual(transactions, [1]);
  assert.deepEqual(seen, [{ bold: true }, { bold: true }]);
});

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

test('a ribbon focus change cancels a queued replacement caret restore', () => {
  const raf = [];
  const model = new OutlineSelectionModel();
  const fields = register(model, createFields(['abc'], { raf }));
  const calls = [];
  fields[0].setRange = (offset) => calls.push(['range', offset]);
  fields[0].focus = (offset) => calls.push(['focus', offset]);
  model.setCaret(fields[0], 0);
  assert.equal(model.replace('x'), true);
  fields[0].root.ownerDocument.activeElement = { ribbon: true };
  raf.splice(0).forEach((callback) => callback());
  assert.deepEqual(calls, []);
});

test('a queued cross-field replacement restores while the old focus is unchanged', () => {
  const raf = [];
  const model = new OutlineSelectionModel();
  const fields = register(model, createFields(['ab', 'cd'], { raf }));
  const calls = [];
  fields[0].setRange = (offset) => calls.push(['range', offset]);
  fields[0].focus = (offset) => calls.push(['focus', offset]);
  model.update(fields[0], 2, 2);
  assert.equal(model.extend(fields[0], 1), true);
  assert.equal(model.extend(fields[1], 1), true);
  // Cross-field selection leaves the native active element on the end field
  // while replacement is committed to the first field.
  fields[0].root.ownerDocument.activeElement = fields[1].root;
  assert.equal(model.replace('x'), true);
  raf.splice(0).forEach((callback) => callback());
  assert.deepEqual(calls, [
    ['range', 3],
    ['focus', 3],
  ]);
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

import assert from 'node:assert/strict';
import test from 'node:test';

class FakeElement {
  constructor(tagName, children = [], attrs = []) {
    this.tagName = tagName;
    this.childNodes = children;
    this.attrs = new Set(attrs);
    this.firstChild = children[0] ?? null;
    for (const [index, child] of children.entries()) {
      child.parentNode = this;
      child.nextSibling = children[index + 1] ?? null;
    }
  }
  hasAttribute(name) {
    return this.attrs.has(name);
  }
}
class FakeBr extends FakeElement {
  constructor(attrs = []) {
    super('BR', [], attrs);
  }
}
class FakeText {
  constructor(text) {
    this.nodeType = 3;
    this.textContent = text;
  }
}
globalThis.Node = { TEXT_NODE: 3 };
globalThis.HTMLElement = FakeElement;
globalThis.HTMLBRElement = FakeBr;

const { richTextPoint, richTextValue, selectRichText } =
  await import('../src/core/rich-text-dom.ts');

function text(value) {
  return new FakeText(value);
}
function root(...children) {
  return new FakeElement('DIV', children);
}
function roundTrip(input) {
  const value = richTextValue(input);
  for (let target = 0; target <= value.length; target++) {
    const point = richTextPoint(input, target);
    assert.equal(richTextValue(input, point), value.slice(0, target), `target=${target}`);
  }
}

test('richTextPoint round-trips implicit block separators', () => {
  const input = root(
    new FakeElement('DIV', [text('one')]),
    new FakeElement('DIV', [text('two')]),
    new FakeElement('DIV', [text('three')]),
  );
  assert.equal(richTextValue(input), 'one\ntwo\nthree');
  roundTrip(input);
});

test('empty editing placeholders resolve to offset zero', () => {
  for (const input of [root(new FakeBr()), root(new FakeBr(['data-caret-end']))]) {
    const point = richTextPoint(input, 0);
    assert.equal(richTextValue(input, point), '');
  }
});

test('richTextPoint round-trips authored BR boundaries', () => {
  const input = root(text('one'), new FakeBr(), text('two'));
  assert.equal(richTextValue(input), 'one\ntwo');
  roundTrip(input);
});

test('richTextPoint round-trips leading BR and empty block boundaries', () => {
  for (const input of [
    root(new FakeBr(), text('two')),
    root(new FakeElement('DIV'), text('two')),
  ]) {
    roundTrip(input);
  }
});

test('richTextPoint round-trips block boundaries around empty blocks', () => {
  const input = root(
    new FakeElement('DIV', [text('one')]),
    new FakeElement('DIV'),
    new FakeElement('DIV', [text('two')]),
  );
  const value = richTextValue(input);
  assert.equal(value, 'one\ntwo');
  roundTrip(input);
});

test('richTextPoint round-trips nested blocks and surrogate pairs', () => {
  const input = root(
    new FakeElement('DIV', [new FakeElement('P', [text('😀')])]),
    new FakeElement('DIV', [text('x')]),
  );
  assert.equal(richTextValue(input), '😀\nx');
  roundTrip(input);
});

test('selection restoration counts authored line breaks', () => {
  const first = text('one');
  const second = text('two');
  const input = root(first, new FakeBr(), second);
  let restored;
  const document = {
    createTreeWalker: () => {
      const nodes = [first, second];
      return { nextNode: () => nodes.shift() ?? null };
    },
    getSelection: () => ({
      setBaseAndExtent(anchorNode, anchorOffset, focusNode, focusOffset) {
        restored = {
          start: { node: anchorNode, offset: anchorOffset },
          end: { node: focusNode, offset: focusOffset },
        };
      },
    }),
  };
  input.ownerDocument = document;
  globalThis.document = document;
  globalThis.NodeFilter = { SHOW_TEXT: 4 };
  selectRichText(input, 4, 6);
  assert.equal(richTextValue(input, restored.start), 'one\n');
  assert.equal(richTextValue(input, restored.end), 'one\ntw');
});

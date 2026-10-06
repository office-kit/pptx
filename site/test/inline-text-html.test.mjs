import assert from 'node:assert/strict';
import test from 'node:test';
import {
  addBlankSlide,
  addSlideTable,
  createPresentation,
  getTableCellParagraphs,
  getTableCells,
  setPresentationTheme,
  setTableCellTextFormat,
  inches,
} from '@office-kit/pptx';
import {
  resolveEditingTextColor,
  resolveEditingTextFormatColors,
} from '../src/lib/editor/core/inline-text-html.ts';
import { textClipboardHtml } from '../src/lib/editor/core/html-text-clipboard.ts';

test('inline HTML exports character outline, shadow, and glow with zoom-scaled CSS', () => {
  const previousDocument = globalThis.document;
  const makeElement = () => {
    const style = {
      setProperty(name, value) {
        this[name] = value;
      },
    };
    return {
      style,
      dataset: {},
      children: [],
      textContent: '',
      append(...nodes) {
        this.children.push(...nodes);
      },
      replaceChildren() {},
      get outerHTML() {
        const css = Object.entries(style)
          .filter(([name]) => name !== 'setProperty' && typeof name === 'string')
          .map(([name, value]) => `${name}:${value}`)
          .join(';');
        const content =
          this.textContent || this.children.map((child) => child.outerHTML ?? '').join('');
        return `<span style="${css}">${content}</span>`;
      },
    };
  };
  globalThis.document = { createElement: makeElement };
  try {
    const html = textClipboardHtml({
      text: 'A',
      formats: [
        {
          start: 0,
          end: 1,
          format: {
            outline: { color: '#123456', widthEmu: 9525 },
            shadow: { color: '#000000', offsetEmu: 9525, blurEmu: 19050, angleDeg: 0 },
            glow: { color: '#00FF00', radiusEmu: 9525 },
          },
        },
      ],
    });
    assert.match(html, /-webkit-text-stroke:calc\(1px \* var\(--text-zoom, 1\)\) #123456/);
    assert.match(html, /textShadow:/);
    assert.match(html, /var\(--text-zoom, 1\)/);
  } finally {
    globalThis.document = previousDocument;
  }
});

test('table editing resolves scheme run colors through the slide color map', () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const table = addSlideTable(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(1),
    rows: [['Text']],
  });
  const cell = getTableCells(table)[0][0];
  setPresentationTheme(pres, { dark1: '#123456' });
  setTableCellTextFormat(cell, {
    color: 'scheme:tx1',
    underline: 'dotted',
    underlineColor: 'scheme:accent1',
  });

  const rawColor = getTableCellParagraphs(cell)[0].elements[0].format.color;
  assert.equal(rawColor, 'tx1');
  assert.equal(resolveEditingTextColor(pres, table, rawColor), '#123456');
  const rawUnderlineColor = getTableCellParagraphs(cell)[0].elements[0].format.underlineColor;
  assert.equal(rawUnderlineColor, 'accent1');
  setPresentationTheme(pres, { accent1: '#654321' });
  assert.equal(resolveEditingTextColor(pres, table, rawUnderlineColor), '#654321');
  assert.equal(resolveEditingTextColor(pres, table, '#123456'), '#123456');
});

test('text effects resolve scheme colors before HTML export', () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const table = addSlideTable(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(1),
    rows: [['Text']],
  });
  setPresentationTheme(pres, {
    accent1: '#112233',
    accent2: '#223344',
    accent3: '#334455',
    accent4: '#445566',
  });
  const format = resolveEditingTextFormatColors(pres, table, {
    color: 'accent1',
    outline: { color: 'accent2', widthEmu: 9525 },
    shadow: { color: 'accent3' },
    glow: { color: 'accent4' },
  });
  assert.equal(format.color, '#112233');
  assert.equal(format.outline?.color, '#223344');
  assert.equal(format.shadow?.color, '#334455');
  assert.equal(format.glow?.color, '#445566');
});

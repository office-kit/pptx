import { test } from 'node:test';
import assert from 'node:assert/strict';
import { tableCellShape, expandTableSelection } from '../src/table-cells.ts';

test('cell overlays scale unequal columns and rows, cover merge anchors and rotate about table center', () => {
  const cell = {
    text: 'Cell',
    span: { gridSpan: 1, rowSpan: 1, hMerge: false, vMerge: false },
    paragraphs: [
      {
        align: 'center',
        elements: [{ kind: 'r', text: 'Cell', format: { bold: true } }],
        endFormat: null,
      },
    ],
    margins: { left: null, right: null, top: null, bottom: null },
    anchor: 'center',
    direction: null,
  };
  const shape = {
    id: 2,
    bounds: { x: 100, y: 200, w: 600, h: 300 },
    rotation: 0,
    table: {
      columnWidths: [100, 200],
      rowHeights: [50, 100],
      cells: [
        [cell, cell],
        [cell, cell],
      ],
    },
  };
  const result = tableCellShape(shape, 1, 1);
  assert.deepEqual(result.bounds, { x: 300, y: 300, w: 400, h: 200 });
  assert.equal(result.runs[0].format.bold, true);
  assert.equal(result.align, 'ctr');
  const styled = tableCellShape(shape, 1, 1, {
    fill: '#4472C4',
    color: '#FFFFFF',
    font: 'Aptos',
  });
  assert.equal(styled.fill, '#4472C4');
  assert.deepEqual(styled.runs[0].format, { bold: true, color: '#FFFFFF', font: 'Aptos' });
  assert.equal(styled.format.color, '#FFFFFF');
  cell.paragraphs[0].elements[0].format.color = '#123456';
  assert.equal(
    tableCellShape(shape, 1, 1, {
      fill: '#4472C4',
      color: '#FFFFFF',
      font: 'Aptos',
    }).runs[0].format.color,
    '#123456',
  );

  shape.rotation = 90;
  const rotated = tableCellShape(shape, 1, 1);
  assert.ok(Math.abs(rotated.bounds.x - 150) < 1e-8);
  assert.ok(Math.abs(rotated.bounds.y - 350) < 1e-8);
  shape.rotation = 0;
  shape.table.cells[0][0] = { ...cell, span: { ...cell.span, gridSpan: 2 } };
  shape.table.cells[0][1] = { ...cell, span: { ...cell.span, hMerge: true } };
  assert.deepEqual(tableCellShape(shape, 0, 0).bounds, { x: 100, y: 200, w: 600, h: 100 });
  assert.equal(tableCellShape(shape, 0, 1), null);
});

test('selection expansion reaches a fixed point across intersecting merged cells', () => {
  const cells = Array.from({ length: 3 }, () =>
    Array.from({ length: 3 }, () => ({
      span: { rowSpan: 1, gridSpan: 1, hMerge: false, vMerge: false },
    })),
  );
  cells[0][0].span.gridSpan = 2;
  cells[0][1].span.hMerge = true;
  cells[1][1].span.rowSpan = 2;
  cells[2][1].span.vMerge = true;
  const range = { row: 0, column: 0, rows: 2, columns: 1 };
  assert.deepEqual(expandTableSelection({ cells }, range), {
    row: 0,
    column: 0,
    rows: 3,
    columns: 2,
  });
  assert.deepEqual(range, { row: 0, column: 0, rows: 2, columns: 1 });
  assert.deepEqual(expandTableSelection({ cells }, { row: 0, column: 2, rows: 3, columns: 1 }), {
    row: 0,
    column: 2,
    rows: 3,
    columns: 1,
  });
});

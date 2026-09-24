import type { EditorShape } from './editor.ts';

/** Project a visible cell into the existing direct-text editor's coordinate model. */
export function tableCellShape(
  shape: EditorShape,
  row: number,
  column: number,
  appearance?: { fill: string; color: string; font: string },
): EditorShape | null {
  const table = shape.table,
    box = shape.bounds;
  const cell = table?.cells[row]?.[column];
  if (!table || !box || !cell || cell.span.hMerge || cell.span.vMerge) return null;
  const sum = (values: number[]) => values.reduce((a, b) => a + b, 0);
  const sx = sum(table.columnWidths) > 0 ? box.w / sum(table.columnWidths) : 1,
    sy = sum(table.rowHeights) > 0 ? box.h / sum(table.rowHeights) : 1;
  const x = sum(table.columnWidths.slice(0, column)) * sx;
  const y = sum(table.rowHeights.slice(0, row)) * sy;
  const w = sum(table.columnWidths.slice(column, column + cell.span.gridSpan)) * sx;
  const h = sum(table.rowHeights.slice(row, row + cell.span.rowSpan)) * sy;
  const angle = (shape.rotation * Math.PI) / 180;
  const dx = x + w / 2 - box.w / 2,
    dy = y + h / 2 - box.h / 2;
  const runs: EditorShape['runs'] = [],
    paragraphs: EditorShape['paragraphs'] = [];
  const defaults = appearance ? { color: appearance.color, font: appearance.font } : {};
  let offset = 0;
  for (const paragraph of cell.paragraphs) {
    const start = offset;
    for (const element of paragraph.elements) {
      const length = element.kind === 'br' ? 1 : element.text.length;
      runs.push({
        start: offset,
        end: offset + length,
        format: { ...defaults, ...element.format },
      });
      offset += length;
    }
    paragraphs.push({
      start,
      end: offset,
      align: paragraph.align,
      bullets: paragraph.properties?.bullet ?? null,
      properties: {
        align: paragraph.align,
        level: paragraph.properties?.level ?? 0,
        marL: paragraph.properties?.marL ?? null,
        marR: paragraph.properties?.marR ?? null,
        indent: paragraph.properties?.indent ?? null,
        lineSpacing: paragraph.properties?.lineSpacing ?? null,
        spcBefPts: paragraph.properties?.spcBefPts ?? null,
        spcAftPts: paragraph.properties?.spcAftPts ?? null,
        rtl: paragraph.properties?.rtl ?? null,
        bullet: paragraph.properties?.bullet ?? null,
      },
    });
    offset++;
  }
  return {
    ...shape,
    textable: true,
    fill: appearance?.fill ?? shape.fill,
    text: cell.text,
    textLinks: cell.textLinks,
    link: null,
    runs,
    paragraphs,
    bounds: {
      x: box.x + box.w / 2 + dx * Math.cos(angle) - dy * Math.sin(angle) - w / 2,
      y: box.y + box.h / 2 + dx * Math.sin(angle) + dy * Math.cos(angle) - h / 2,
      w,
      h,
    },
    format: { ...defaults, ...(runs[0]?.format ?? cell.paragraphs[0]?.endFormat) },
    align:
      cell.paragraphs[0]?.align === 'center'
        ? 'ctr'
        : cell.paragraphs[0]?.align === 'right'
          ? 'r'
          : 'l',
    anchor: cell.anchor ?? 'top',
    anchorCenter: false,
    autoFit: 'none',
    autoFitParams: null,
    textDirection: cell.direction ?? 'horz',
    textFrame: {
      margins: {
        left: cell.margins.left ?? 91440,
        right: cell.margins.right ?? 91440,
        top: cell.margins.top ?? 45720,
        bottom: cell.margins.bottom ?? 45720,
      },
      anchor: cell.anchor ?? 'top',
      wrap: true,
      columns: null,
    },
  };
}

/** Expand a rectangular selection until every intersecting merged cell is whole. */
export function expandTableSelection(
  table: NonNullable<EditorShape['table']>,
  range: { row: number; column: number; rows: number; columns: number },
) {
  let { row, column } = range;
  let bottom = row + range.rows,
    right = column + range.columns;
  let expanded: boolean;
  do {
    expanded = false;
    table.cells.forEach((cells, r) =>
      cells.forEach((cell, c) => {
        if (cell.span.hMerge || cell.span.vMerge) return;
        const b = r + cell.span.rowSpan,
          end = c + cell.span.gridSpan;
        if (
          r < bottom &&
          b > row &&
          c < right &&
          end > column &&
          (r < row || c < column || b > bottom || end > right)
        ) {
          row = Math.min(row, r);
          column = Math.min(column, c);
          bottom = Math.max(bottom, b);
          right = Math.max(right, end);
          expanded = true;
        }
      }),
    );
  } while (expanded);

  return { row, column, rows: bottom - row, columns: right - column };
}

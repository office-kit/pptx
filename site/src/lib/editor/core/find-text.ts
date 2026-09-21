import {
  getSlideShapes,
  getSlides,
  getShapeId,
  getShapeText,
  isTableShape,
  getTableCells,
  getTableCellText,
  type PresentationData,
  type SlideShapeData,
  type TableCellData,
} from '@office-kit/pptx';

export function searchPattern(query: string, matchCase: boolean): RegExp {
  return new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), matchCase ? 'gu' : 'giu');
}

export interface TextMatch {
  slideIndex: number;
  shapeId: number;
  target:
    | { kind: 'shape'; shape: SlideShapeData }
    | { kind: 'cell'; cell: TableCellData; row: number; col: number };
  start: number;
  end: number;
  text: string;
}

export function findText(
  pres: PresentationData,
  query: string,
  matchCase: boolean,
  slideIndex?: number,
): TextMatch[] {
  if (!query) return [];
  const pattern = searchPattern(query, matchCase);
  const matches: TextMatch[] = [];
  for (const [index, slide] of getSlides(pres).entries()) {
    if (slideIndex !== undefined && index !== slideIndex) continue;
    for (const shape of getSlideShapes(slide)) {
      const shapeId = getShapeId(shape);
      const collect = (text: string, target: TextMatch['target']) => {
        for (const match of text.matchAll(pattern)) {
          matches.push({
            slideIndex: index,
            shapeId,
            target,
            text,
            start: match.index,
            end: match.index + match[0].length,
          });
        }
      };
      if (isTableShape(shape)) {
        getTableCells(shape).forEach((row, r) =>
          row.forEach((cell, col) =>
            collect(getTableCellText(cell), { kind: 'cell', cell, row: r, col }),
          ),
        );
      } else collect(getShapeText(shape), { kind: 'shape', shape });
    }
  }
  return matches;
}

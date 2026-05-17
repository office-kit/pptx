// getTableSize — sum of column widths + row heights.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTable,
  emu,
  getSlides,
  getTableSize,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getTableSize', () => {
  it('sums column widths and row heights', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addSlideTable(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(6),
      h: inches(2),
      rows: [
        ['a', 'b', 'c'],
        ['d', 'e', 'f'],
      ],
      colWidths: [emu(1000000), emu(2000000), emu(3000000)],
      rowHeights: [emu(400000), emu(600000)],
    });
    expect(getTableSize(table)).toEqual({ width: 6000000, height: 1000000 });
  });

  it('throws on a non-table shape', async () => {
    const { addSlideShape } = await import('../src/api/index.ts');
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const rect = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    expect(() => getTableSize(rect)).toThrow();
  });
});

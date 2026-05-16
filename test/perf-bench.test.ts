// Performance smoke test.
//
// Builds a synthetic deck of N slides (default 100) populated with shapes,
// text, tables, and charts, then measures the load → save round-trip.
//
// Not part of the default CI run — gated on `PERF=1`. Run with:
//
//     PERF=1 pnpm vitest run test/perf-bench.test.ts
//
// Targets (M-series Node 20+):
//   - 100-slide synthetic deck save:    < 2000ms
//   - load → save round-trip:           < 2000ms

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideChart,
  addSlideShape,
  addSlideTable,
  addSlideTextBox,
  findSlideLayout,
  inches,
  loadPresentation,
  savePresentation,
  setShapeFill,
  setShapeRunFormat,
  setSlideTitle,
} from '../src/api/index.ts';

const ENABLED = process.env.PERF === '1';

const SLIDE_COUNT = Number(process.env.PERF_SLIDES ?? '100');

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

describe.skipIf(!ENABLED)('performance bench', () => {
  it(`builds + saves ${SLIDE_COUNT}-slide deck under 2s`, async () => {
    const tBuildStart = performance.now();
    const pres = await loadPresentation(await readFile(fixturePath));
    const layout = findSlideLayout(pres, 'Title and Content') ?? findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('no layout');

    for (let i = 0; i < SLIDE_COUNT; i++) {
      const slide = addSlide(pres, { layout });
      setSlideTitle(slide, `Slide ${i + 1}`);

      // Mix of content kinds to simulate a real deck.
      const mod = i % 4;
      if (mod === 0) {
        const box = addSlideTextBox(slide, {
          x: inches(0.7),
          y: inches(1.5),
          w: inches(9),
          h: inches(5),
          text: `Body line 1\nBody line 2\nBody line 3 — slide ${i + 1}`,
        });
        setShapeRunFormat(box, 0, 0, { size: 20, bold: true });
      } else if (mod === 1) {
        for (let s = 0; s < 6; s++) {
          const r = Math.floor(s / 3);
          const c = s % 3;
          const shape = addSlideShape(slide, {
            preset: 'roundRect',
            x: (inches(0.7) + c * inches(3)) as ReturnType<typeof inches>,
            y: (inches(1.5) + r * inches(2.2)) as ReturnType<typeof inches>,
            w: inches(2.7),
            h: inches(2),
            text: `Card ${s + 1}`,
          });
          setShapeFill(
            shape,
            ['#2E75B6', '#548235', '#C00000', '#7030A0', '#BF8F00', '#0070C0'][s]!,
          );
        }
      } else if (mod === 2) {
        addSlideTable(slide, {
          x: inches(0.7),
          y: inches(1.5),
          w: inches(9),
          h: inches(4),
          rows: [
            ['Col A', 'Col B', 'Col C', 'Col D'],
            ['1', '2', '3', '4'],
            ['5', '6', '7', '8'],
            ['9', '10', '11', '12'],
          ],
          firstRow: true,
          bandRow: true,
        });
      } else {
        addSlideChart(slide, {
          x: inches(0.7),
          y: inches(1.5),
          w: inches(9),
          h: inches(4.5),
          spec: {
            kind: 'column',
            categories: ['Q1', 'Q2', 'Q3', 'Q4'],
            series: [
              { name: 'A', values: [10, 20, 30, 40] },
              { name: 'B', values: [15, 25, 35, 45] },
            ],
            title: `Chart ${i + 1}`,
          },
        });
      }
    }
    const tBuildEnd = performance.now();
    const buildMs = tBuildEnd - tBuildStart;

    const tSaveStart = performance.now();
    const bytes = await savePresentation(pres);
    const tSaveEnd = performance.now();
    const saveMs = tSaveEnd - tSaveStart;

    const tLoadStart = performance.now();
    await loadPresentation(bytes);
    const tLoadEnd = performance.now();
    const loadMs = tLoadEnd - tLoadStart;

    const mb = (bytes.byteLength / (1024 * 1024)).toFixed(2);
    // biome-ignore lint/suspicious/noConsole: bench output.
    console.log(
      `[perf] ${SLIDE_COUNT} slides — build ${buildMs.toFixed(0)}ms — save ${saveMs.toFixed(0)}ms — load ${loadMs.toFixed(0)}ms — size ${mb}MB`,
    );

    expect(saveMs).toBeLessThan(2000);
    expect(loadMs).toBeLessThan(2000);
  });
});

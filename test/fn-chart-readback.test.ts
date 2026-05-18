// `getSlideCharts` — parse charts back from a deck.
//
// Round-trip the same `ChartSpec` through `addSlideChart` →
// `savePresentation` → `loadPresentation` → `getSlideCharts` and
// verify the spec survives.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideChart,
  getSlideCharts,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideCharts', () => {
  it('round-trips a column chart spec', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['Q1', 'Q2', 'Q3', 'Q4'],
        series: [
          { name: 'Revenue', values: [10, 20, 15, 30] },
          { name: 'Cost', values: [5, 7, 9, 11] },
        ],
        title: 'Quarterly',
      },
    });

    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const charts = getSlideCharts(getSlides(reloaded)[0]!);
    expect(charts).toHaveLength(1);
    const spec = charts[0]!.spec;
    expect(spec).not.toBeNull();
    expect(spec!.kind).toBe('column');
    expect(spec!.categories).toEqual(['Q1', 'Q2', 'Q3', 'Q4']);
    expect(spec!.series.map((s) => s.name)).toEqual(['Revenue', 'Cost']);
    expect(spec!.series[0]!.values).toEqual([10, 20, 15, 30]);
    expect(spec!.series[1]!.values).toEqual([5, 7, 9, 11]);
    expect(spec!.title).toBe('Quarterly');
  });

  it('round-trips titleStyle through builder', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2] }],
        title: 'Styled',
        titleStyle: { sizePt: 18, bold: true, color: '#FF0000' },
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const charts = getSlideCharts(getSlides(reloaded)[0]!);
    const spec = charts[0]!.spec!;
    expect(spec.title).toBe('Styled');
    expect(spec.titleStyle?.sizePt).toBe(18);
    expect(spec.titleStyle?.bold).toBe(true);
    expect(spec.titleStyle?.color).toBe('#FF0000');
  });

  it('round-trips valueAxis extras (logBase / displayUnits / tickMark)', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [100, 200] }],
        valueAxis: { logBase: 10, displayUnits: 'thousands' },
        valueAxisMajorTickMark: 'cross',
        categoryAxisMajorTickMark: 'none',
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.valueAxis?.logBase).toBe(10);
    expect(spec.valueAxis?.displayUnits).toBe('thousands');
    expect(spec.valueAxisMajorTickMark).toBe('cross');
    expect(spec.categoryAxisMajorTickMark).toBe('none');
  });

  it('round-trips chart-level dataLabels (toggles + numFmt + dLblPos + separator)', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2] }],
        dataLabels: {
          showValue: true,
          showCategory: false,
          showSeriesName: false,
          showPercent: false,
          numberFormat: '0.00',
          position: 'ctr',
          separator: ', ',
        },
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.dataLabels?.showValue).toBe(true);
    expect(spec.dataLabels?.numberFormat).toBe('0.00');
    expect(spec.dataLabels?.position).toBe('ctr');
    expect(spec.dataLabels?.separator).toBe(', ');
  });

  it('round-trips chart-level dataLabels.textStyle (sz / bold / color)', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2] }],
        dataLabels: {
          showValue: true,
          showCategory: false,
          showSeriesName: false,
          showPercent: false,
          textStyle: { sizePt: 14, bold: true, color: '#FF8800' },
        },
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.dataLabels?.textStyle?.sizePt).toBe(14);
    expect(spec.dataLabels?.textStyle?.bold).toBe(true);
    expect(spec.dataLabels?.textStyle?.color).toBe('#FF8800');
  });

  it('round-trips per-series dataLabels.textStyle', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [
          {
            name: 'X',
            values: [1, 2],
            dataLabels: {
              showValue: true,
              showCategory: false,
              showSeriesName: false,
              showPercent: false,
              textStyle: { sizePt: 9, italic: true, color: '#003366' },
            },
          },
        ],
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.series[0]!.dataLabels?.textStyle?.sizePt).toBe(9);
    expect(spec.series[0]!.dataLabels?.textStyle?.italic).toBe(true);
    expect(spec.series[0]!.dataLabels?.textStyle?.color).toBe('#003366');
  });

  it('round-trips legend position + hiddenIndices + overlay + dispBlanksAs', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [
          { name: 'X', values: [1, 2] },
          { name: 'Y', values: [3, 4] },
          { name: 'Z', values: [5, 6] },
        ],
        legend: { position: 't', overlay: true, hiddenIndices: [1] },
        dispBlanksAs: 'span',
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.legend?.position).toBe('t');
    expect(spec.legend?.overlay).toBe(true);
    expect(spec.legend?.hiddenIndices).toEqual([1]);
    expect(spec.dispBlanksAs).toBe('span');
  });

  it('round-trips valueAxis scaling (min / max / majorUnit / minorUnit / numberFormat)', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [10, 50] }],
        valueAxis: {
          min: 0,
          max: 100,
          majorUnit: 20,
          minorUnit: 5,
          numberFormat: '0.00%',
        },
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.valueAxis?.min).toBe(0);
    expect(spec.valueAxis?.max).toBe(100);
    expect(spec.valueAxis?.majorUnit).toBe(20);
    expect(spec.valueAxis?.minorUnit).toBe(5);
    expect(spec.valueAxis?.numberFormat).toBe('0.00%');
  });

  it('round-trips axis titles + hidden / tickLblSkip / tickLblPos', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2] }],
        valueAxisTitle: 'Revenue',
        valueAxisTitleStyle: { sizePt: 12, bold: true },
        categoryAxisTitle: 'Quarter',
        categoryAxisHidden: true,
        categoryAxisTickLabelSkip: 2,
        categoryAxisTickMarkSkip: 5,
        categoryAxisTickLabelPos: 'low',
        categoryAxisLabelOffset: 150,
        categoryAxisLabelAlign: 'l',
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.valueAxisTitle).toBe('Revenue');
    expect(spec.valueAxisTitleStyle?.sizePt).toBe(12);
    expect(spec.valueAxisTitleStyle?.bold).toBe(true);
    expect(spec.categoryAxisTitle).toBe('Quarter');
    expect(spec.categoryAxisHidden).toBe(true);
    expect(spec.categoryAxisTickLabelSkip).toBe(2);
    expect(spec.categoryAxisTickMarkSkip).toBe(5);
    expect(spec.categoryAxisTickLabelPos).toBe('low');
    expect(spec.categoryAxisLabelOffset).toBe(150);
    expect(spec.categoryAxisLabelAlign).toBe('l');
  });

  it('round-trips chart extras (varyColors / gapWidth / overlap / firstSliceAng / holeSize / dropLines / hiLowLines / titleOverlay / majorGridlines)', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2] }],
        title: 'T',
        titleOverlay: true,
        varyColors: true,
        gapWidthPct: 75,
        overlapPct: -10,
        valueAxisMajorGridlines: true,
        valueAxisMajorGridlineColor: '#AABBCC',
      },
    });
    const slide2 = getSlides(pres)[1]!;
    addSlideChart(slide2, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'doughnut',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [40, 60] }],
        firstSliceAngleDeg: 45,
        holeSizePct: 70,
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec1 = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec1.titleOverlay).toBe(true);
    expect(spec1.varyColors).toBe(true);
    expect(spec1.gapWidthPct).toBe(75);
    expect(spec1.overlapPct).toBe(-10);
    expect(spec1.valueAxisMajorGridlines).toBe(true);
    expect(spec1.valueAxisMajorGridlineColor).toBe('#AABBCC');
    const spec2 = getSlideCharts(getSlides(reloaded)[1]!)[0]!.spec!;
    expect(spec2.firstSliceAngleDeg).toBe(45);
    expect(spec2.holeSizePct).toBe(70);
  });

  it('round-trips plot-area / chart-area fill and stroke colors', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2] }],
        plotAreaFill: '#F0F0F0',
        plotAreaStrokeColor: '#102030',
        chartAreaFill: '#FAFAFA',
        chartAreaStrokeColor: '#888888',
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.plotAreaFill).toBe('#F0F0F0');
    expect(spec.plotAreaStrokeColor).toBe('#102030');
    expect(spec.chartAreaFill).toBe('#FAFAFA');
    expect(spec.chartAreaStrokeColor).toBe('#888888');
  });

  it('round-trips axis tick-label style + rotation via <c:txPr>', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2] }],
        categoryAxisLabelStyle: { sizePt: 9, bold: true, color: '#112233' },
        categoryAxisLabelRotationDeg: 45,
        valueAxisLabelStyle: { sizePt: 10, italic: true },
        valueAxisLabelRotationDeg: -30,
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.categoryAxisLabelStyle?.sizePt).toBe(9);
    expect(spec.categoryAxisLabelStyle?.bold).toBe(true);
    expect(spec.categoryAxisLabelStyle?.color).toBe('#112233');
    expect(spec.categoryAxisLabelRotationDeg).toBe(45);
    expect(spec.valueAxisLabelStyle?.sizePt).toBe(10);
    expect(spec.valueAxisLabelStyle?.italic).toBe(true);
    expect(spec.valueAxisLabelRotationDeg).toBe(-30);
  });

  it('round-trips series-level lineWidth / lineDash / marker / smooth / invertIfNegative', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'line',
        categories: ['A', 'B'],
        series: [
          {
            name: 'Trend',
            values: [10, 20],
            lineWidthEmu: 28575, // 2.25pt
            lineDash: 'dash',
            markerSymbol: 'diamond',
            markerSizePt: 7,
            smooth: true,
          },
        ],
      },
    });
    const slide2 = getSlides(pres)[1]!;
    addSlideChart(slide2, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [-1, 2], invertIfNegative: true }],
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const s1 = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!.series[0]!;
    expect(s1.lineWidthEmu).toBe(28575);
    expect(s1.lineDash).toBe('dash');
    expect(s1.markerSymbol).toBe('diamond');
    expect(s1.markerSizePt).toBe(7);
    expect(s1.smooth).toBe(true);
    const s2 = getSlideCharts(getSlides(reloaded)[1]!)[0]!.spec!.series[0]!;
    expect(s2.invertIfNegative).toBe(true);
  });

  it('round-trips series-level trendline (type / forward / backward / period / color)', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'line',
        categories: ['A', 'B', 'C', 'D'],
        series: [
          {
            name: 'Trend',
            values: [1, 2, 3, 4],
            trendline: {
              type: 'movingAvg',
              period: 2,
              forward: 3,
              backward: 1,
              color: '#993366',
            },
          },
        ],
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const tl = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!.series[0]!.trendline!;
    expect(tl.type).toBe('movingAvg');
    expect(tl.period).toBe(2);
    expect(tl.forward).toBe(3);
    expect(tl.backward).toBe(1);
    expect(tl.color).toBe('#993366');
  });

  it('round-trips per-series dataLabels (toggles + numFmt + position)', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [
          {
            name: 'X',
            values: [1, 2],
            dataLabels: {
              showValue: true,
              showCategory: true,
              showSeriesName: false,
              showPercent: false,
              numberFormat: '$#,##0',
              position: 'inEnd',
            },
          },
          { name: 'Y', values: [3, 4] }, // chart-level only, no per-series
        ],
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.series[0]!.dataLabels?.showValue).toBe(true);
    expect(spec.series[0]!.dataLabels?.showCategory).toBe(true);
    expect(spec.series[0]!.dataLabels?.numberFormat).toBe('$#,##0');
    expect(spec.series[0]!.dataLabels?.position).toBe('inEnd');
    expect(spec.series[1]!.dataLabels).toBeUndefined();
  });

  it('round-trips per-series dPt overrides (pointColors + pointExplosions)', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'pie',
        categories: ['A', 'B', 'C'],
        series: [
          {
            name: 'X',
            values: [10, 20, 30],
            pointColors: ['#FF0000', null, '#00FF00'],
            pointExplosions: [null, 25, null],
          },
        ],
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const ser = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!.series[0]!;
    expect(ser.pointColors?.[0]).toBe('#FF0000');
    expect(ser.pointColors?.[2]).toBe('#00FF00');
    expect(ser.pointExplosions?.[1]).toBe(25);
  });

  it('round-trips legend.textStyle + axis orientation reversals', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2] }],
        legend: {
          position: 'b',
          textStyle: { sizePt: 11, bold: true, color: '#102030' },
        },
        categoryAxisOrientation: 'maxMin',
        valueAxisOrientation: 'maxMin',
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.legend?.textStyle?.sizePt).toBe(11);
    expect(spec.legend?.textStyle?.bold).toBe(true);
    expect(spec.legend?.textStyle?.color).toBe('#102030');
    expect(spec.categoryAxisOrientation).toBe('maxMin');
    expect(spec.valueAxisOrientation).toBe('maxMin');
  });

  it('round-trips trendline custom name', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'line',
        categories: ['A', 'B', 'C'],
        series: [
          {
            name: 'X',
            values: [1, 2, 4],
            trendline: {
              type: 'movingAvg',
              period: 2,
              name: 'MA(2) — smoothed',
            },
          },
        ],
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.series[0]!.trendline!.name).toBe('MA(2) — smoothed');
  });

  it('round-trips trendline displayEquation + displayRSquared', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'line',
        categories: ['A', 'B', 'C'],
        series: [
          {
            name: 'X',
            values: [1, 2, 4],
            trendline: {
              type: 'linear',
              displayEquation: true,
              displayRSquared: true,
            },
          },
        ],
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    const tl = spec.series[0]!.trendline!;
    expect(tl.displayEquation).toBe(true);
    expect(tl.displayRSquared).toBe(true);
  });

  it('round-trips chartStyle preset', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2] }],
        chartStyle: 42,
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.chartStyle).toBe(42);
  });

  it('round-trips roundedCorners=true', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2] }],
        roundedCorners: true,
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.roundedCorners).toBe(true);
  });

  it('round-trips plotVisibleCellsOnly=false', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2] }],
        plotVisibleCellsOnly: false,
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.plotVisibleCellsOnly).toBe(false);
  });

  it('omits plotVisibleCellsOnly on read when the chart uses the default (true)', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2] }],
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.plotVisibleCellsOnly).toBeUndefined();
  });

  it('round-trips axis title rotation', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2] }],
        valueAxisTitle: 'Revenue',
        valueAxisTitleRotationDeg: -90,
        categoryAxisTitle: 'Quarter',
        categoryAxisTitleRotationDeg: 45,
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.valueAxisTitleRotationDeg).toBe(-90);
    expect(spec.categoryAxisTitleRotationDeg).toBe(45);
  });

  it('round-trips all 4 gridline colors', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'bar',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2] }],
        categoryAxisMajorGridlines: true,
        categoryAxisMinorGridlines: true,
        valueAxisMajorGridlines: true,
        valueAxisMinorGridlines: true,
        categoryAxisMajorGridlineColor: '#111111',
        categoryAxisMinorGridlineColor: '#222222',
        valueAxisMajorGridlineColor: '#333333',
        valueAxisMinorGridlineColor: '#444444',
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.categoryAxisMajorGridlineColor).toBe('#111111');
    expect(spec.categoryAxisMinorGridlineColor).toBe('#222222');
    expect(spec.valueAxisMajorGridlineColor).toBe('#333333');
    expect(spec.valueAxisMinorGridlineColor).toBe('#444444');
  });

  it('round-trips category-axis gridlines (major + minor)', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'bar',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2] }],
        categoryAxisMajorGridlines: true,
        categoryAxisMinorGridlines: true,
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.categoryAxisMajorGridlines).toBe(true);
    expect(spec.categoryAxisMinorGridlines).toBe(true);
  });

  it('round-trips axis line colors', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2] }],
        categoryAxisLineColor: '#FF0000',
        valueAxisLineColor: '#00AA00',
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.categoryAxisLineColor).toBe('#FF0000');
    expect(spec.valueAxisLineColor).toBe('#00AA00');
  });

  it('round-trips categoryAxisNumberFormat', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['2024-01', '2024-02', '2024-03'],
        series: [{ name: 'X', values: [1, 2, 3] }],
        categoryAxisNumberFormat: 'mmm-yyyy',
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.categoryAxisNumberFormat).toBe('mmm-yyyy');
  });

  it('round-trips minor tick mark on both axes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2] }],
        categoryAxisMinorTickMark: 'out',
        valueAxisMinorTickMark: 'cross',
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.categoryAxisMinorTickMark).toBe('out');
    expect(spec.valueAxisMinorTickMark).toBe('cross');
  });

  it('round-trips valueAxisCrossBetween', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2] }],
        valueAxisCrossBetween: 'midCat',
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.valueAxisCrossBetween).toBe('midCat');
  });

  it('round-trips valueAxisCrosses enum + numeric forms', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2] }],
        valueAxisCrosses: 'max',
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.valueAxisCrosses).toBe('max');

    // Now the numeric form on a second chart.
    const pres2 = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    addSlideChart(getSlides(pres2)[0]!, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [-5, 5] }],
        valueAxisCrosses: { at: -3 },
      },
    });
    const b2 = await savePresentation(pres2);
    const r2 = await loadPresentation(b2);
    const s2 = getSlideCharts(getSlides(r2)[0]!)[0]!.spec!;
    expect(s2.valueAxisCrosses).toEqual({ at: -3 });
  });

  it('distinguishes bar from column on the same barChart wire format', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(3),
      spec: {
        kind: 'bar',
        categories: ['A', 'B'],
        series: [{ name: 'S', values: [1, 2] }],
      },
    });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.kind).toBe('bar');
  });

  it('reads pie and doughnut kinds', async () => {
    for (const kind of ['pie', 'doughnut'] as const) {
      const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
      const slide = getSlides(pres)[0]!;
      addSlideChart(slide, {
        x: inches(0),
        y: inches(0),
        w: inches(4),
        h: inches(3),
        spec: {
          kind,
          categories: ['X', 'Y', 'Z'],
          series: [{ name: 'S', values: [1, 2, 3] }],
        },
      });
      const reloaded = await loadPresentation(await savePresentation(pres));
      expect(getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!.kind).toBe(kind);
    }
  });

  it('returns custom series colors verbatim', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(3),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2], color: '#112233' }],
      },
    });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const spec = getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!;
    expect(spec.series[0]!.color).toBe('#112233');
  });

  it('returns empty array for slides without charts', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(getSlideCharts(slide)).toEqual([]);
  });

  it('parses an unmodified template chart loaded fresh from disk', async () => {
    // Add a chart, save, load again — proves the reader doesn't depend
    // on in-memory state created by addSlideChart.
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    addSlideChart(getSlides(pres)[0]!, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(3),
      spec: {
        kind: 'line',
        categories: ['Jan', 'Feb', 'Mar'],
        series: [{ name: 'A', values: [3, 5, 7] }],
      },
    });
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const charts = getSlideCharts(getSlides(reloaded)[0]!);
    expect(charts[0]!.spec!.kind).toBe('line');
  });
});

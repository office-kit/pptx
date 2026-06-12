// Chart reader: verifies the recently-added optional fields surface
// correctly when their underlying XML is authored. Each block is a
// focused unit test that constructs a minimal `<c:chartSpace>` with
// just enough markup to exercise one feature, parses via
// `readChartSpec`, and asserts the expected `ChartSpec` shape.

import { describe, expect, it } from 'vitest';
import { readChartSpec } from '../src/internal/chartml/index.ts';
import { parseXml } from '../src/internal/xml/index.ts';

const wrap = (
  innerChart: string,
): string => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>${innerChart}</c:chart>
</c:chartSpace>`;

const MIN_PLOT_AREA = `
  <c:plotArea>
    <c:layout/>
    <c:barChart>
      <c:barDir val="col"/>
      <c:grouping val="clustered"/>
      <c:ser>
        <c:idx val="0"/>
        <c:order val="0"/>
        <c:tx><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:tx>
        <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
      </c:ser>
    </c:barChart>
    <c:valAx>
      <c:axId val="1"/>
      <c:crossAx val="2"/>
    </c:valAx>
    <c:catAx>
      <c:axId val="2"/>
      <c:crossAx val="1"/>
    </c:catAx>
  </c:plotArea>`;

describe('chart reader: extras', () => {
  it('reads `<c:title><c:tx><c:rich><a:r><a:rPr>` titleStyle', () => {
    const xml = wrap(`
      <c:title>
        <c:tx>
          <c:rich>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p>
              <a:r>
                <a:rPr sz="2400" b="1">
                  <a:solidFill><a:srgbClr val="2A5CAA"/></a:solidFill>
                </a:rPr>
                <a:t>Quarterly</a:t>
              </a:r>
            </a:p>
          </c:rich>
        </c:tx>
      </c:title>
      ${MIN_PLOT_AREA}
    `);
    const spec = readChartSpec(parseXml(xml).root)!;
    expect(spec.title).toBe('Quarterly');
    expect(spec.titleStyle).toEqual({ sizePt: 24, bold: true, color: '#2A5CAA' });
  });

  it('reads `<c:title><c:tx><c:strRef>` cell-reference titles', () => {
    const xml = wrap(`
      <c:title>
        <c:tx>
          <c:strRef>
            <c:f>Sheet1!$A$1</c:f>
            <c:strCache>
              <c:ptCount val="1"/>
              <c:pt idx="0"><c:v>Linked Title</c:v></c:pt>
            </c:strCache>
          </c:strRef>
        </c:tx>
      </c:title>
      ${MIN_PLOT_AREA}
    `);
    const spec = readChartSpec(parseXml(xml).root)!;
    expect(spec.title).toBe('Linked Title');
  });

  it('reads `<c:valAx><c:scaling><c:logBase>` log scale', () => {
    const xml = wrap(`
      <c:plotArea>
        <c:layout/>
        <c:barChart>
          <c:barDir val="col"/>
          <c:grouping val="clustered"/>
          <c:ser>
            <c:idx val="0"/>
            <c:order val="0"/>
            <c:tx><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:tx>
            <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
          </c:ser>
        </c:barChart>
        <c:valAx>
          <c:axId val="1"/>
          <c:scaling><c:logBase val="10"/></c:scaling>
          <c:crossAx val="2"/>
        </c:valAx>
        <c:catAx><c:axId val="2"/><c:crossAx val="1"/></c:catAx>
      </c:plotArea>
    `);
    const spec = readChartSpec(parseXml(xml).root)!;
    expect(spec.valueAxis?.logBase).toBe(10);
  });

  it('reads `<c:dispUnits><c:builtInUnit>` displayUnits', () => {
    const xml = wrap(`
      <c:plotArea>
        <c:layout/>
        <c:barChart>
          <c:barDir val="col"/>
          <c:grouping val="clustered"/>
          <c:ser>
            <c:idx val="0"/>
            <c:order val="0"/>
            <c:tx><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:tx>
            <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
          </c:ser>
        </c:barChart>
        <c:valAx>
          <c:axId val="1"/>
          <c:dispUnits><c:builtInUnit val="millions"/></c:dispUnits>
          <c:crossAx val="2"/>
        </c:valAx>
        <c:catAx><c:axId val="2"/><c:crossAx val="1"/></c:catAx>
      </c:plotArea>
    `);
    const spec = readChartSpec(parseXml(xml).root)!;
    expect(spec.valueAxis?.displayUnits).toBe('millions');
  });

  it('reads `<c:legend><c:legendEntry><c:delete>` hidden series indices', () => {
    const xml = wrap(`
      ${MIN_PLOT_AREA}
      <c:legend>
        <c:legendPos val="r"/>
        <c:legendEntry>
          <c:idx val="2"/>
          <c:delete val="1"/>
        </c:legendEntry>
        <c:legendEntry>
          <c:idx val="0"/>
          <c:delete val="1"/>
        </c:legendEntry>
      </c:legend>
    `);
    const spec = readChartSpec(parseXml(xml).root)!;
    expect(spec.legend?.hiddenIndices).toEqual([2, 0]);
  });

  it('reads `<c:catAx><c:majorTickMark>` tickMark mode', () => {
    const xml = wrap(`
      <c:plotArea>
        <c:layout/>
        <c:barChart>
          <c:barDir val="col"/>
          <c:grouping val="clustered"/>
          <c:ser>
            <c:idx val="0"/>
            <c:order val="0"/>
            <c:tx><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:tx>
            <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
          </c:ser>
        </c:barChart>
        <c:valAx>
          <c:axId val="1"/>
          <c:majorTickMark val="cross"/>
          <c:crossAx val="2"/>
        </c:valAx>
        <c:catAx>
          <c:axId val="2"/>
          <c:majorTickMark val="none"/>
          <c:crossAx val="1"/>
        </c:catAx>
      </c:plotArea>
    `);
    const spec = readChartSpec(parseXml(xml).root)!;
    expect(spec.valueAxisMajorTickMark).toBe('cross');
    expect(spec.categoryAxisMajorTickMark).toBe('none');
  });

  it('reads `<c:catAx><c:numRef>` numeric/date categories', () => {
    const xml = wrap(`
      <c:plotArea>
        <c:layout/>
        <c:barChart>
          <c:barDir val="col"/>
          <c:grouping val="clustered"/>
          <c:ser>
            <c:idx val="0"/>
            <c:order val="0"/>
            <c:tx><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:tx>
            <c:cat>
              <c:numRef>
                <c:f>Sheet1!$A$1:$A$3</c:f>
                <c:numCache>
                  <c:formatCode>General</c:formatCode>
                  <c:pt idx="0"><c:v>2024</c:v></c:pt>
                  <c:pt idx="1"><c:v>2025</c:v></c:pt>
                  <c:pt idx="2"><c:v>2026</c:v></c:pt>
                </c:numCache>
              </c:numRef>
            </c:cat>
            <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
          </c:ser>
        </c:barChart>
        <c:valAx><c:axId val="1"/><c:crossAx val="2"/></c:valAx>
        <c:catAx><c:axId val="2"/><c:crossAx val="1"/></c:catAx>
      </c:plotArea>
    `);
    const spec = readChartSpec(parseXml(xml).root)!;
    expect(spec.categories).toEqual(['2024', '2025', '2026']);
  });

  it('reads `<c:trendline><c:forward>/<c:backward>` extensions', () => {
    const xml = wrap(`
      <c:plotArea>
        <c:layout/>
        <c:lineChart>
          <c:grouping val="standard"/>
          <c:ser>
            <c:idx val="0"/>
            <c:order val="0"/>
            <c:tx><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:tx>
            <c:trendline>
              <c:trendlineType val="linear"/>
              <c:forward val="3"/>
              <c:backward val="1"/>
            </c:trendline>
            <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
          </c:ser>
        </c:lineChart>
        <c:valAx><c:axId val="1"/><c:crossAx val="2"/></c:valAx>
        <c:catAx><c:axId val="2"/><c:crossAx val="1"/></c:catAx>
      </c:plotArea>
    `);
    const spec = readChartSpec(parseXml(xml).root)!;
    const tl = spec.series[0]?.trendline;
    expect(tl?.type).toBe('linear');
    expect(tl?.forward).toBe(3);
    expect(tl?.backward).toBe(1);
  });

  it('reads `<c:varyColors>` flag', () => {
    const xml = wrap(`
      <c:plotArea>
        <c:layout/>
        <c:barChart>
          <c:varyColors val="1"/>
          <c:barDir val="col"/>
          <c:grouping val="clustered"/>
          <c:ser>
            <c:idx val="0"/>
            <c:order val="0"/>
            <c:tx><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:tx>
            <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
          </c:ser>
        </c:barChart>
        <c:valAx><c:axId val="1"/><c:crossAx val="2"/></c:valAx>
        <c:catAx><c:axId val="2"/><c:crossAx val="1"/></c:catAx>
      </c:plotArea>
    `);
    const spec = readChartSpec(parseXml(xml).root)!;
    expect(spec.varyColors).toBe(true);
  });
});

describe('chart reader: scatter / radar / bubble', () => {
  it('detects scatter, reads xy tuples + scatterStyle', () => {
    const xml = wrap(`
      <c:plotArea>
        <c:layout/>
        <c:scatterChart>
          <c:scatterStyle val="lineMarker"/>
          <c:ser>
            <c:idx val="0"/>
            <c:order val="0"/>
            <c:tx><c:strLit><c:pt idx="0"><c:v>S1</c:v></c:pt></c:strLit></c:tx>
            <c:xVal>
              <c:numLit>
                <c:pt idx="0"><c:v>1</c:v></c:pt>
                <c:pt idx="1"><c:v>2</c:v></c:pt>
                <c:pt idx="2"><c:v>3</c:v></c:pt>
              </c:numLit>
            </c:xVal>
            <c:yVal>
              <c:numLit>
                <c:pt idx="0"><c:v>10</c:v></c:pt>
                <c:pt idx="1"><c:v>20</c:v></c:pt>
                <c:pt idx="2"><c:v>15</c:v></c:pt>
              </c:numLit>
            </c:yVal>
          </c:ser>
          <c:axId val="1"/>
          <c:axId val="2"/>
        </c:scatterChart>
        <c:valAx><c:axId val="1"/><c:crossAx val="2"/></c:valAx>
        <c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>
      </c:plotArea>
    `);
    const spec = readChartSpec(parseXml(xml).root)!;
    expect(spec.kind).toBe('scatter');
    expect(spec.scatterStyle).toBe('lineMarker');
    // For scatter, `values` is the y-channel; `xValues` the paired x.
    expect(spec.series[0]!.xValues).toEqual([1, 2, 3]);
    expect(spec.series[0]!.values).toEqual([10, 20, 15]);
  });

  it('detects bubble, reads bubbleSize + bubbleScale + sizeRepresents', () => {
    const xml = wrap(`
      <c:plotArea>
        <c:layout/>
        <c:bubbleChart>
          <c:ser>
            <c:idx val="0"/>
            <c:order val="0"/>
            <c:tx><c:strLit><c:pt idx="0"><c:v>B1</c:v></c:pt></c:strLit></c:tx>
            <c:xVal>
              <c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit>
            </c:xVal>
            <c:yVal>
              <c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numLit>
            </c:yVal>
            <c:bubbleSize>
              <c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt><c:pt idx="1"><c:v>9</c:v></c:pt></c:numLit>
            </c:bubbleSize>
          </c:ser>
          <c:bubbleScale val="80"/>
          <c:sizeRepresents val="area"/>
          <c:axId val="1"/>
          <c:axId val="2"/>
        </c:bubbleChart>
        <c:valAx><c:axId val="1"/><c:crossAx val="2"/></c:valAx>
        <c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>
      </c:plotArea>
    `);
    const spec = readChartSpec(parseXml(xml).root)!;
    expect(spec.kind).toBe('bubble');
    expect(spec.bubbleScale).toBe(80);
    expect(spec.bubbleSizeRepresents).toBe('area');
    expect(spec.series[0]!.xValues).toEqual([1, 2]);
    expect(spec.series[0]!.values).toEqual([10, 20]);
    expect(spec.series[0]!.bubbleSizes).toEqual([4, 9]);
  });

  it("maps sizeRepresents='w' to 'width'", () => {
    const xml = wrap(`
      <c:plotArea>
        <c:layout/>
        <c:bubbleChart>
          <c:ser>
            <c:idx val="0"/>
            <c:order val="0"/>
            <c:tx><c:strLit><c:pt idx="0"><c:v>B</c:v></c:pt></c:strLit></c:tx>
            <c:yVal><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:yVal>
            <c:bubbleSize><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:bubbleSize>
          </c:ser>
          <c:sizeRepresents val="w"/>
          <c:axId val="1"/>
          <c:axId val="2"/>
        </c:bubbleChart>
        <c:valAx><c:axId val="1"/><c:crossAx val="2"/></c:valAx>
        <c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>
      </c:plotArea>
    `);
    const spec = readChartSpec(parseXml(xml).root)!;
    expect(spec.bubbleSizeRepresents).toBe('width');
  });

  it('detects radar, reads cat/val + radarStyle', () => {
    const xml = wrap(`
      <c:plotArea>
        <c:layout/>
        <c:radarChart>
          <c:radarStyle val="filled"/>
          <c:ser>
            <c:idx val="0"/>
            <c:order val="0"/>
            <c:tx><c:strLit><c:pt idx="0"><c:v>R1</c:v></c:pt></c:strLit></c:tx>
            <c:cat>
              <c:strLit>
                <c:pt idx="0"><c:v>A</c:v></c:pt>
                <c:pt idx="1"><c:v>B</c:v></c:pt>
                <c:pt idx="2"><c:v>C</c:v></c:pt>
              </c:strLit>
            </c:cat>
            <c:val>
              <c:numLit>
                <c:pt idx="0"><c:v>3</c:v></c:pt>
                <c:pt idx="1"><c:v>5</c:v></c:pt>
                <c:pt idx="2"><c:v>2</c:v></c:pt>
              </c:numLit>
            </c:val>
          </c:ser>
          <c:axId val="1"/>
          <c:axId val="2"/>
        </c:radarChart>
        <c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx>
        <c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>
      </c:plotArea>
    `);
    const spec = readChartSpec(parseXml(xml).root)!;
    expect(spec.kind).toBe('radar');
    expect(spec.radarStyle).toBe('filled');
    expect(spec.categories).toEqual(['A', 'B', 'C']);
    expect(spec.series[0]!.values).toEqual([3, 5, 2]);
    // Radar carries no x-channel / bubble sizes.
    expect(spec.series[0]!.xValues).toBeUndefined();
    expect(spec.series[0]!.bubbleSizes).toBeUndefined();
  });
});
